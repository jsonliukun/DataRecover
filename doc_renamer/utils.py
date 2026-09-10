"""跨模块使用的无状态工具：内容指纹、文本规整、名称净化和路径约束。

这些函数没有 provider 或命令行依赖，因此 extractors、providers 和 workflow 都能
复用。它们分别承担不同层次的保证：

``sha256_file``
    为“生成方案时的文件”和“执行时的文件”建立内容一致性校验；它不是数字签名，
    不能证明文件来源可信。
``compact_text``
    折叠空白并按字符数截断，用于控制标题、节选和错误说明的体积。
``sanitize_stem``
    把不可信的模型标题转换成 Windows 可接受的文件名主干。
``safe_resolve_below``
    约束方案中的相对路径始终落在用户指定的根目录内，是 apply/undo 的路径边界。
``unique_name``
    同时检查现有文件和本批次预留名称，在生成方案时避免目标名碰撞。

调用方仍需组合这些原语：名称合法不代表路径安全，路径位于根目录内也不代表文件
内容未变化。因此 apply 会依次执行名称、路径、存在性和哈希检查。
"""

from __future__ import annotations

import hashlib
import re
from pathlib import Path


# Windows 普通文件名不能包含 < > : " / \ | ? * 或 U+0000～U+001F 控制字符。
# 正则只负责字符级过滤；尾部空格/点和保留设备名在 sanitize_stem 的后续步骤处理。
INVALID_WINDOWS_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')

# Python 3 的 Unicode 正则中，\s 会覆盖常见空格、制表符、换行及多种 Unicode
# 空白。统一成半角空格可让标题和节选稳定地保持单行、便于审核和提示词计量；
# csv 模块本身能引用含换行字段，因此这不是表格正确性的唯一保护。
WHITESPACE = re.compile(r"\s+")

# Windows 传统设备名不区分大小写，即使再加扩展名也不能可靠地作为普通文件名。
# sanitize_stem 接收的是不含扩展名的主干，因此直接对大写主干做集合成员判断。
# 星号解包两个生成器，得到 COM1～COM9 与 LPT1～LPT9。
RESERVED_WINDOWS_NAMES = {
    "CON",
    "PRN",
    "AUX",
    "NUL",
    *(f"COM{i}" for i in range(1, 10)),
    *(f"LPT{i}" for i in range(1, 10)),
}


def sha256_file(path: Path, chunk_size: int = 1024 * 1024) -> str:
    """分块计算文件 SHA-256，返回 64 位小写十六进制摘要。

    ``path.read_bytes()`` 会一次性占用与文件大小相当的内存；恢复文件可能达到
    GB 级，因此使用固定大小块增量更新摘要。默认块为 1 MiB，内存占用基本不随
    文件大小增长。``chunk_size`` 由内部调用方保证为正数。

    海象运算符在判断 EOF 的同时保存本次字节块；空字节表示读取结束。文件不存在、
    无权限或读取中发生 I/O 错误时不吞异常，由上层决定是写进方案还是终止操作。

    摘要在这里用于一致性和重复检测，不承担身份认证：攻击者若能同时篡改方案中的
    哈希，SHA-256 本身无法证明方案可信。
    """
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        while chunk := handle.read(chunk_size):
            digest.update(chunk)
    return digest.hexdigest()


def compact_text(text: str, max_chars: int) -> str:
    """把文本规整为单行单空格形式，并按 Unicode 字符数截断。

    先折叠全部连续空白并清理首尾，再取前 ``max_chars`` 个 Python 字符。字符数
    不等于模型 token 数，也不等于 UTF-8 字节数；这里只提供稳定、与模型无关的
    输入上限。段落边界会丢失，因此该函数适合搜索摘要/提示词，不适合保真导出。

    截断后的 ``rstrip`` 清掉恰好落在边界上的尾部空格。调用方应传入非负上限；
    本函数故意保持轻量，不额外验证业务参数。
    """
    text = WHITESPACE.sub(" ", text).strip()
    return text[:max_chars].rstrip()


def sanitize_stem(value: str, fallback: str = "未识别文档", max_length: int = 120) -> str:
    """把模型/规则标题规整成 Windows 文件名主干，不包含扩展名。

    处理顺序很重要，每一步都针对一种具体的失败场景：
      1. 非法字符替换为下划线——否则文件创建会直接抛 OSError。
      2. 压缩空白，并去掉首尾空格、点和下划线。Windows 对尾部空格/点处理不可靠；
         去掉下划线则是为了减少替换非法字符留下的视觉噪音。
      3. 把连续下划线压成一个，避免标题里的一串非法字符变成 `_____` 这种噪声。
      4. 若清洗后为空（例如原标题全是非法字符），退回 fallback。
      5. 若主干以 = + - @ 开头，补下划线。这里同时保护最终文件名和方案中的
         ``proposed_name``；workflow 还会对其他表格列执行可逆的公式转义。
      6. 避开 CON、NUL 等 Windows 保留设备名，同样用前缀下划线破坏匹配。
      7. 截断到 ``max_length`` 并再次清理尾部，避免边界留下空格、点或下划线。

    ``fallback`` 和 ``max_length`` 是内部扩展点：调用方若自定义，应提供本身合法的
    fallback，并避免把 max_length 设得过小。函数只生成主干；扩展名保留、目录
    边界和目标冲突分别由 workflow / safe_resolve_below / unique_name 负责。
    """
    value = INVALID_WINDOWS_CHARS.sub("_", value)
    value = WHITESPACE.sub(" ", value).strip(" ._")
    value = re.sub(r"_+", "_", value)
    if not value:
        value = fallback
    # 这是 proposed_name 的第一层保护；表格其他列由 workflow 统一可逆转义。
    if value[0] in "=+-@":
        value = f"_{value}"
    if value.upper() in RESERVED_WINDOWS_NAMES:
        value = f"_{value}"
    value = value[:max_length].rstrip(" ._")
    return value or fallback


def safe_resolve_below(root: Path, relative_path: str) -> Path:
    """解析方案路径，并拒绝当前解析结果位于 ``root`` 之外的情况。

    这是本工具最重要的安全防线。relative_path 来自方案文件（可被人工编辑），
    属于外部输入；若直接使用，像 `../../Windows/System32/xxx` 这样的路径会让
    改名/撤销操作跑到目标目录之外，造成不可挽回的损失（目录穿越 / Path Traversal）。

    ``resolve`` 会规范化 ``..`` 并按当前文件系统状态解析符号链接；即使传入绝对
    ``relative_path`` 导致 Path 拼接忽略 root，后续祖先判断也会拦截外部路径。
    候选允许等于 root，本函数只判断边界，调用方随后还会用 ``is_file`` 拒绝目录。

    越界时抛 ``ValueError``。本函数不检查存在性、文件类型或哈希，也不能消除
    “校验后由其他进程替换符号链接”的竞态；本工具假定工作副本不被恶意并发修改，
    并在实际改名前紧接着完成剩余检查。
    """
    root = root.resolve()
    candidate = (root / relative_path).resolve()
    if candidate != root and root not in candidate.parents:
        raise ValueError(f"路径越界：{relative_path}")
    return candidate


def unique_name(parent: Path, proposed_name: str, reserved: set[Path]) -> str:
    """预留一个当前不冲突的文件名，必要时追加 ``_2``、``_3``……

    reserved 参数用于解决「磁盘上还不存在、但本批次内已被占用」的问题：
    生成方案时文件尚未改名，此时多个文档可能被建议成同一个名字。如果只检查
    candidate.exists()，这些同批次内的冲突会在真正执行改名时集中爆发
    （后一个改名因目标已存在而失败）。因此把已分配的目标路径记入 reserved，
    在内存中提前去重。

    现有磁盘对象用 ``exists`` 检查；尚未落盘的同批目标用 ``reserved`` 检查。
    规范化 Path 可消除 ``.`` / ``..`` 等表示差异；在 Windows 上 Path 比较还遵循
    Windows 的大小写归一规则，比裸字符串集合更符合目标平台语义。

    返回的是纯文件名（candidate.name）而非完整路径，因为调用方需要把它写进
    方案文件的 proposed_name 列，该列要求不含目录成分。

    前置条件是 ``proposed_name`` 已由 ``sanitize_stem`` 等逻辑变成单独文件名；
    本函数不会再次做非法字符或路径穿越检查。返回纯文件名，并立即把绝对候选加入
    ``reserved``，所以调用本身具有“批内占位”副作用。并发进程仍可能在方案生成后
    抢占该名称，apply 因此还会重新检查目标是否存在。
    """
    candidate = parent / proposed_name
    counter = 2
    while candidate.exists() or candidate.resolve() in reserved:
        # 用 Path.stem / suffix 拆分，保证原扩展名始终保留在末尾
        # （如 `报告_2.docx` 而不是 `报告.docx_2`）。
        candidate = parent / f"{Path(proposed_name).stem}_{counter}{Path(proposed_name).suffix}"
        counter += 1
    reserved.add(candidate.resolve())
    return candidate.name
