"""命令行适配层：解析参数、装配依赖、调用工作流并设置进程退出码。

控制台脚本 ``doc-renamer`` 和 ``python -m doc_renamer`` 最终都会调用本模块的
``main``。本层只负责把命令行输入转换为 Python 参数，不负责抽取文档、请求模型
或直接改名；这些业务行为分别位于 providers、extractors 和 workflow。

职责可概括为四项：

    1. 用 argparse 定义子命令、参数类型、默认值和帮助文本。
    2. 把命令行覆盖值与环境变量组合成 ProviderSettings。
    3. 调用 workflow 中与子命令对应的入口函数。
    4. 把结果以 JSON/提示文字输出，并用退出码区分成功、中断和失败。

四个子命令组成一个显式操作闭环：

    plan        扫描并写方案；不修改被扫描文件
    probe-api   用内置虚拟文档测试模型 API；不读取用户文件，但会访问网络
    apply       只处理方案中 status=approved 的行；会改名并写日志
    undo        按日志恢复名称；会修改文件名

``plan`` 产生的正常建议当前默认标为 approved，但仍只有显式运行 ``apply`` 才会
改名。将有副作用的动作拆成独立子命令，是命令层最重要的安全边界，也方便未来
的 GUI 或批处理代码直接复用底层工作流。
"""

from __future__ import annotations

import argparse
import json
import sys
from dataclasses import asdict
from pathlib import Path

from dotenv import load_dotenv

from .providers import ProviderSettings, build_provider
from .workflow import apply_plan, create_plan, undo_log


def _provider_args(parser: argparse.ArgumentParser) -> None:
    """给 ``plan`` 和 ``probe-api`` 注册同一组模型选择参数。

    抽成辅助函数可防止两个子解析器的参数契约逐渐分叉。三个参数都可省略：
    显式传入时覆盖对应环境变量，省略时由 ``ProviderSettings.from_env`` 读取
    ``.env`` / 进程环境。

    ``--provider`` 用 choices 在解析阶段拦截拼写错误。``--api-base`` 只填写基础
    地址，具体的 ``/responses`` 或 ``/chat/completions`` 由 provider 追加。
    API Key 刻意没有命令行参数，因为命令行可能留在 shell 历史、进程列表和运维
    审计日志中；密钥只能通过环境变量提供。
    """
    parser.add_argument("--provider", choices=("rules", "responses", "chat"))
    parser.add_argument("--api-base", help="API 基础地址，例如 https://api.openai.com/v1")
    parser.add_argument("--model", help="模型 ID；密钥必须放在环境变量中")


def build_parser() -> argparse.ArgumentParser:
    """构建完整的命令行解析器。

    单独构造 parser，测试或未来的界面层就能检查参数契约，而不必真正执行命令。
    ``prog`` 固定帮助信息中的程序名；``required=True`` 要求用户显式选择动作，
    空命令会由 argparse 打印用法并以参数错误结束。

    子命令契约：

    * ``plan``：root 为扫描根目录；``--output`` 默认 ``rename_plan.txt``，后缀
      决定写出 TSV 还是 CSV。``--limit`` 截取排序后的前 N 个文件，适合小批试跑。
      ``--max-chars`` 限制正文节选字符数；``--max-pages`` 只影响 PDF/PPTX。
    * ``probe-api``：不接受 root，确保只使用 provider 内置的虚拟文档。
    * ``apply``：接收 root、plan 和可选日志路径；``--overwrite`` 只允许覆盖旧日志，
      不会放宽“目标文件不能已存在”的保护。
    * ``undo``：接收 root 和日志；恢复冲突仍由 workflow 拒绝。

    ``type=Path`` 只做类型转换，不验证路径存在性或边界；这些依赖实际文件系统的
    检查必须延迟到 workflow。两个 ``--overwrite`` 都是 ``store_true``，因此只有
    用户明确输入开关时才允许覆盖对应的方案或日志。
    """
    parser = argparse.ArgumentParser(
        prog="doc-renamer",
        description="根据恢复文档的内容生成候选文件名；默认绝不修改原文件。",
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    plan = subparsers.add_parser("plan", help="扫描文件并生成改名方案（TXT/TSV 或 CSV）")
    plan.add_argument("root", type=Path, help="恢复文档所在目录")
    plan.add_argument("--output", type=Path, default=Path("rename_plan.txt"))
    plan.add_argument("--limit", type=int, help="只处理前 N 个文件，用于小批量试跑")
    plan.add_argument("--max-chars", type=int, default=6000, help="每个文档最多发送的字符数")
    plan.add_argument("--max-pages", type=int, default=5, help="PDF/PPTX 最多读取的页数")
    plan.add_argument("--overwrite", action="store_true", help="明确覆盖已经存在的方案文件")
    _provider_args(plan)

    probe = subparsers.add_parser("probe-api", help="用一段虚拟内容测试模型 API")
    _provider_args(probe)

    apply_cmd = subparsers.add_parser("apply", help="执行方案中 status=approved 的改名")
    apply_cmd.add_argument("root", type=Path)
    apply_cmd.add_argument("plan", type=Path)
    apply_cmd.add_argument("--log", type=Path, default=Path("rename_log.txt"))
    apply_cmd.add_argument("--overwrite", action="store_true", help="明确覆盖已经存在的日志文件")

    undo = subparsers.add_parser("undo", help="根据改名日志撤销")
    undo.add_argument("root", type=Path)
    undo.add_argument("log", type=Path)
    return parser


def _settings(args: argparse.Namespace) -> ProviderSettings:
    """从解析好的参数中取出提供方配置，交给 ProviderSettings 统一处理。

    当前仅 ``plan`` 和 ``probe-api`` 会调用本函数。使用带默认值的 ``getattr``，
    是为了让辅助函数不依赖某个子解析器的 Namespace 形状；若未来复用到其他命令，
    缺少可选属性时仍会自然回退到环境变量。

    参数为 None 表示“命令行没有覆盖”。API Key 不进入 Namespace，始终只能来自
    环境变量，从数据结构层面降低调试打印或参数日志泄露密钥的风险。
    """
    return ProviderSettings.from_env(
        provider=getattr(args, "provider", None),
        api_base=getattr(args, "api_base", None),
        model=getattr(args, "model", None),
    )


def main(argv: list[str] | None = None) -> int:
    """程序主入口，返回进程退出码。

    ``argv=None`` 时 argparse 使用真实的 ``sys.argv[1:]``；测试可以传入独立
    列表，在当前进程内验证分发逻辑。``load_dotenv`` 从当前目录向上查找
    ``.env``，且默认不覆盖进程中已有的环境变量，所以优先级是：显式命令行参数、
    已存在的进程环境变量、``.env``、代码默认值。

    四个分支保持展开，是因为它们的副作用和结果结构不同。``plan`` 写方案并返回
    total/planned/duplicates/errors；``probe-api`` 返回 NameSuggestion；``apply``
    写日志并返回 approved/renamed/skipped/failed；``undo`` 返回恢复统计。

    JSON 使用 ``ensure_ascii=False`` 保留中文、使用缩进便于人读。产生方案或日志
    时另行打印绝对路径，避免用户在错误目录寻找相对路径产物。统计命令正常退出
    不等于每一行都成功，调用方仍应检查 ``errors`` 或 ``failed``。

    退出码约定：成功为 0；Ctrl+C 返回 130；运行期异常返回 1。argparse 的参数
    错误发生在 try 块之前，由 argparse 按惯例返回 2。普通异常只把“类型 + 消息”
    写到 stderr，不向最终用户倾倒 traceback；开发排错应运行测试或直接调用底层
    函数取得完整堆栈。
    """
    load_dotenv()
    args = build_parser().parse_args(argv)
    try:
        if args.command == "plan":
            provider = build_provider(_settings(args))
            result = create_plan(
                args.root,
                args.output,
                provider,
                limit=args.limit,
                max_chars=args.max_chars,
                max_pages=args.max_pages,
                overwrite=args.overwrite,
            )
            print(json.dumps(result, ensure_ascii=False, indent=2))
            print(f"方案已写入：{args.output.resolve()}")
        elif args.command == "probe-api":
            # probe() 返回 NameSuggestion dataclass；标准 json 编码器不直接认识
            # dataclass，因此先用 asdict 转为只含基础类型的普通字典。
            provider = build_provider(_settings(args))
            print(json.dumps(asdict(provider.probe()), ensure_ascii=False, indent=2))
        elif args.command == "apply":
            result = apply_plan(args.root, args.plan, args.log, overwrite=args.overwrite)
            print(json.dumps(result, ensure_ascii=False, indent=2))
            print(f"日志已写入：{args.log.resolve()}")
        elif args.command == "undo":
            print(json.dumps(undo_log(args.root, args.log), ensure_ascii=False, indent=2))
        return 0
    except KeyboardInterrupt:
        # 中断消息走 stderr，使 stdout 保持为可供脚本消费的正常 JSON/路径输出。
        print("已中止。", file=sys.stderr)
        return 130
    except Exception as exc:
        print(f"错误：{type(exc).__name__}: {exc}", file=sys.stderr)
        return 1
