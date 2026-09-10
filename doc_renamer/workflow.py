"""核心流程编排：方案生成 -> 人工审核 -> 执行改名 -> 按日志撤销。

本文件的作用
------------
这是整个工具的心脏，把 extractors（读内容）和 providers（想名字）串成一条
可审核、可回滚的操作流水线。面向数据恢复场景，所有设计都围绕一个前提：
「文件是用户仅存的副本，任何自动操作都必须是可逆、可追溯、可中止的」。

三个公开函数构成完整生命周期：
    create_plan  扫描目录，生成改名方案表（只读，绝不改动任何原文件）
    apply_plan   读取方案中 status=approved 的行，执行实际改名，同时写日志
    undo_log     只依据日志中 status=renamed 的行，把文件改回原名

为什么叫「方案」（plan）而不是「结果」：因为中间必须插入人工审核。方案文件
本质是一张表格。成功生成建议的行默认标成 approved；用户在 Excel/WPS 中审核，
把不想执行的行改成 pending。这里有两组容易混淆、但分属不同层级的词：

    approved / duplicate / error  是方案表每一行的 status，即工作流状态；
    planned / duplicates / errors 是 create_plan 返回的统计键，即本次生成结果的计数。

尤其是 planned：它表示「成功得到命名建议的文件数」，不是会被写进 status 列的
状态值。正常行虽然默认写成 approved，统计时仍增加 counts["planned"]。apply_plan
重新读取人工审核后的表格，才用 counts["approved"] 统计这次实际入选的行数。

本文件还处理了几个容易被忽视的工程细节：
    - 输出格式按扩展名选择，读取时再独立探测编码和分隔符。
    - 表格公式注入防护，以及读回时与之配对的可逆转义。
    - 执行前预解析整个方案（避免后半段格式错误造成半批次执行）。
    - 每个执行结果写入日志后立即 flush + fsync，尽量缩小崩溃时的记录缺口。
"""

from __future__ import annotations

import csv
import os
from contextlib import contextmanager
from datetime import datetime, timezone
from pathlib import Path
from typing import Callable, Iterator

from .extractors import DISCOVERED_EXTENSIONS, extract_document
from .providers import NamingProvider
from .utils import safe_resolve_below, sanitize_stem, sha256_file, unique_name


# 方案文件的列定义。列表顺序就是写出后的列顺序，不能随意打乱：status、
# original_name、proposed_name、confidence 靠前，便于人工审核；sha256、
# embedded_title、extraction_method 等靠后，保留「依据什么生成名称」以及
# 「执行时如何确认仍是同一份文件」的审计信息。
#
# status 是持久化在每行中的工作流状态，不是 create_plan 返回字典的统计键。
# 正常建议默认写 approved；duplicate/error 会覆盖该默认值。与之对应的生成期
# 统计键分别叫 planned/duplicates/errors，其中 planned 刻意使用过去分词，
# 表示「建议已生成」，避免把统计数据误当成方案状态再次写回。
PLAN_FIELDS = [
    "status",
    "original_relpath",
    "original_name",
    "proposed_name",
    "confidence",
    "category",
    "reason",
    "sha256",
    "duplicate_of",
    "embedded_title",
    "extraction_method",
    "extraction_error",
    "size",
    "modified_at",
]

# 改名日志的列定义。undo_log 不反推方案，而是以这份日志为事实来源，因此要
# 记录改名前后路径、方案中的内容哈希和本次尝试结果。成功行的 message 为空；
# 失败行保留异常类型与消息，但 undo 只消费 status 精确等于 renamed 的行。
LOG_FIELDS = ["timestamp", "status", "before_relpath", "after_relpath", "sha256", "message"]

# 读取方案/日志时必须存在的列（缺任意一列都说明文件不是本工具生成的，
# 或已被人为改坏），缺失时直接报错而不是继续跑。
# 这里只覆盖关键列，其余列缺失不影响执行，保证向前兼容性。
PLAN_REQUIRED_FIELDS = {"status", "original_relpath", "proposed_name", "sha256"}
LOG_REQUIRED_FIELDS = {"status", "before_relpath", "after_relpath", "sha256"}


def _output_delimiter(path: Path) -> str:
    """按输出文件的扩展名决定分隔符：.txt/.tsv 用制表符，其余用逗号。

    这里决定的是「新文件如何写出」，不是「已有文件如何读入」：

    - 输出名以 .txt 或 .tsv 结尾时写 TSV（Tab-Separated Values）；
    - 其他后缀（通常是 .csv）写逗号分隔内容。

    .txt + 制表符是为了减少中文 Windows 上 Excel/WPS 对 CSV 分隔符、区域设置
    和另存格式处理不一致的问题；它并不承诺关闭表格软件自身的日期/数字类型
    推断。读取端不会相信扩展名，而由 _detect_delimiter 根据表头重新判断，
    因而也兼容早期版本产生的「后缀是 .txt、内容仍用逗号分隔」的文件。
    """
    return "\t" if path.suffix.lower() in {".txt", ".tsv"} else ","


def _looks_like_spreadsheet_formula(value: str) -> bool:
    """判断字符串是否可能被表格软件（Excel/WPS）当作公式执行。

    =、+、-、@ 是常见的公式起始字符。判断前仅去掉空格、制表符和换行，
    是为了覆盖有意用前导空白隐藏危险字符的输入；不调用无参数 strip，避免
    意外扩大这里约定的字符集合。`bool(stripped)` 同时保护空串，防止访问 [0]
    时抛出 IndexError。

    这是面向分隔文本被表格软件打开时的保守检查，并不是公式语法解析器。
    宁可把少量看起来像负数或普通文本的值按文本写出，也不能让来自文件名、
    文档标题或模型输出的内容有机会作为公式求值。
    """
    stripped = value.lstrip(" \t\r\n")
    return bool(stripped) and stripped[0] in "=+-@"


def _spreadsheet_safe(value: object) -> object:
    """写出前转义：给公式样式的字符串加前导单引号，使其被当作纯文本。

    在原始 CSV/TSV 单元格前增加单引号，是常见的公式注入缓解措施：Excel/WPS
    通常把这类内容按文本处理，而不是执行 `=HYPERLINK(...)` 等公式。这里对
    方案和日志的每个字符串列统一处理，因为文件名、内嵌标题、模型输出乃至
    错误消息都不是可以直接信任的表格内容。

    可逆性规则：
      - 公式样式字符串：`=1+1` 写成 `'=1+1`；
      - 原本以单引号开头：`'报告` 写成 `''报告`，避免读回时误删原始引号；
      - 普通字符串、空串和非字符串对象原样返回。

    “可逆”特指本工具的 _spreadsheet_safe -> _spreadsheet_restore 序列化往返。
    若外部表格软件在保存时主动删除、添加或规范化单引号，任何文本格式都无法
    保证逐字符还原；路径边界和源文件哈希仍会照常校验，但它们不能猜回被外部
    软件改写的单元格文本，因此人工审核仍不可省略。
    """
    if not isinstance(value, str) or not value:
        return value
    # 开头的单引号也加倍，使转义在 apply/undo 读取时保持可逆。
    if value.startswith("'") or _looks_like_spreadsheet_formula(value):
        return "'" + value
    return value


def _spreadsheet_restore(value: object) -> object:
    """读入后还原 _spreadsheet_safe 添加的单引号。

    这一步必须精确：只有当第一个单引号「确实是转义标记」时才去掉它，
    否则会破坏本来就以单引号开头的正常内容（如文件名 `'年度报告`）。

    判断依据与 _spreadsheet_safe 的编码规则严格配对：
    - 原值以单引号开头 -> 写出时变成两个单引号，因此 remainder 也以单引号开头；
    - 原值是公式样式   -> 写出时前缀一个单引号，因此 remainder 是公式样式。
    两者满足其一才认定是转义，去掉第一个引号；否则原样返回。

    例如用户后来手工填入单个普通值 `'年度报告` 时，remainder 为“年度报告”，
    两项条件都不满足，因此这个单引号会保留。函数一次最多移除一层由本工具
    增加的前缀，原值自身的单引号仍属于数据。
    """
    if not isinstance(value, str) or not value.startswith("'"):
        return value
    remainder = value[1:]
    if remainder.startswith("'") or _looks_like_spreadsheet_formula(remainder):
        return remainder
    return value


def _safe_row(row: dict[str, object]) -> dict[str, object]:
    """对整个字典逐值做公式注入防护，用于写出前的批量处理。

    用字典推导式生成新字典而不是原地修改，避免日志/方案写出层反过来污染
    调用方仍在使用的领域对象。csv.DictWriter 会再负责必要的引号和分隔符转义；
    本函数处理的是单元格内容安全，两者职责不同。
    """
    return {key: _spreadsheet_safe(value) for key, value in row.items()}


def _restore_row(row: dict[str | None, object]) -> dict[str | None, object]:
    """对整个字典逐值还原转义，用于读入后的批量处理。

    键类型是 `str | None` 而非 `str`，这是 csv.DictReader 的实际行为：
    当某行的字段数多于表头列数时，多出来的值会以 None 为键存入列表。
    _spreadsheet_restore 只转换字符串，因此这种列表会原样保留；这里的目标只是
    撤销本工具的单元格转义，并不在这一层偷偷丢弃或修复结构异常的数据。
    """
    return {key: _spreadsheet_restore(value) for key, value in row.items()}


def _detect_encoding(path: Path) -> str:
    """探测表格文件的文本编码，供 csv 模块正确读取。

    这个问题在中文 Windows 上非常现实：Excel 和 WPS 另存为 CSV 时，
    默认编码往往是 ANSI（在简体中文系统上通常是 GBK/GB18030）而不是 UTF-8。
    用错编码会使中文路径和名称变成乱码，轻则找不到源文件，重则让人工误审，
    因此读取端要优先使用明确标记，并在无法确认格式时停止执行。

    判断顺序与优先级：
      1. 空文件直接报错——没有表头就无法解读表格，继续下去毫无意义。
      2. BOM 优先。BOM 是文件开头声明编码和（对 UTF-16 而言）字节序的标记，
         在这里属于确定性证据：
             \\xff\\xfe  小端 UTF-16
             \\xfe\\xff  大端 UTF-16（两者字节序相反）
             \\xef\\xbb\\xbf  UTF-8 BOM（记事本另存 UTF-8 时常见）
      3. 无 BOM 时按 utf-8 -> gb18030 -> big5 的优先级试解码，第一个能完整
         解码的候选即采用。utf-8 放在前面是为了优先保护现代文件；gb18030
         兼容 GBK，是简体中文 Windows 的主要兜底；big5 用于部分繁体旧文件。
      4. 全部失败则抛出带解决方案的异常，提示用户另存为 UTF-8 或 GB18030。

    无 BOM 的编码探测本质上是启发式判断：某些短字节串可能同时符合多个编码，
    “能够解码”不等于一定选中了作者原本使用的编码。此函数通过固定优先级获得
    可预测结果，而不声称能消除这种信息论上的歧义。如果表头因此变成乱码，
    _open_dict_reader 随后的必需列检查会失败，阻止改名继续进行。

    这里一次读取全部字节。无 BOM 时，各候选编码会对整份数据试解码，避免文件
    前部恰好是 ASCII、真正的非法字节到后半段才出现；有 BOM 时直接按标记返回，
    实际解码错误会在调用方消费 DictReader 时暴露。方案和日志应远小于待处理
    文档，因此用少量额外内存换取更完整的判断是可接受的。
    """
    data = path.read_bytes()
    if not data:
        raise ValueError(f"表格文件为空：{path}")
    if data.startswith((b"\xff\xfe", b"\xfe\xff")):
        return "utf-16"
    if data.startswith(b"\xef\xbb\xbf"):
        return "utf-8-sig"

    # “ANSI”不是固定编码名；简体中文 Windows 中通常对应 GBK/GB18030。
    # 必须对完整字节串试解码，不能只看开头恰好全部为 ASCII 的表头。
    for encoding in ("utf-8", "gb18030", "big5"):
        try:
            data.decode(encoding)
            return encoding
        except UnicodeDecodeError:
            continue
    raise UnicodeError(f"无法识别表格文件编码：{path}；请保存为 UTF-8、Unicode 文本或 GB18030")


def _detect_delimiter(sample: str, path: Path) -> str:
    """从表头行推断分隔符（制表符 / 逗号 / 分号）。

    做法：统计第一个非空物理行中三种候选字符的出现次数，取最多者。方案和
    日志的表头由固定、纯 ASCII 且不含这三种字符的字段名构成，所以对本工具的
    文件而言，这比让 Sniffer 猜测更可控。分号只用于兼容部分受区域设置影响的
    CSV；本工具自己不会写出分号格式。

    先跳过空行（`if line.strip()`）是因为文件开头可能有空行，用它当表头会
    统计出全 0 的结果。若三种分隔符都没出现（counts 最大值为 0），说明这
    根本不是表格文件，抛出明确错误而不是随便选一个继续，否则后续解析会
    产生一堆莫名其妙的列名。

    next(生成器, "") 取得第一个非空行；默认值 "" 让全空白文件落入同一个
    “无法识别”错误，而不是泄漏 StopIteration。若候选计数相同，max 会按字典
    插入顺序选择；正常的固定表头不会出现这种歧义，真正是否选对还会由后续
    必需列检查验证。

    注意读取时不依据 .txt/.csv 后缀。扩展名只能影响 _output_delimiter 的写出
    选择，不能作为外部软件另存后的可信格式证据。
    """
    first_line = next((line for line in sample.splitlines() if line.strip()), "")
    counts = {delimiter: first_line.count(delimiter) for delimiter in ("\t", ",", ";")}
    delimiter = max(counts, key=counts.get)
    if counts[delimiter] == 0:
        raise ValueError(f"无法识别表格分隔符：{path}；支持制表符、逗号或分号")
    return delimiter


@contextmanager
def _open_dict_reader(path: Path, required_fields: set[str]) -> Iterator[csv.DictReader]:
    """打开表格文件并返回配置好的 DictReader，同时校验必需列是否存在。

    @contextmanager 把「打开文件 - 处理 - 关闭文件」的样板代码封装成一个
    `with` 语句。生成器在 yield 之前的部分相当于 __enter__，yield 之后
    （此处没有）相当于 __exit__；无论调用方是否抛异常，with 语句退出时
    文件都会被正确关闭。这避免了忘记 close 导致的文件句柄泄漏。

    关键细节：
    - resolve() 先把方案/日志路径规范化；这不是 root 边界校验，因为这些文件
      本来就允许放在数据目录之外，真正受 root 约束的是表格内的文件路径。
    - `newline=""` 把换行识别交给 csv 模块，避免文本层提前转换换行而干扰
      csv 对带引号字段、不同平台行尾的处理。
    - 先 read(65536) 取得表头样本用于探测分隔符，再 seek(0) 回到开头，
      让正式 DictReader 从表头重新读取。这里的 64 KiB 只用于找第一个非空
      表头行，不承担完整解析；尤其对带 BOM 的文件，后续消费 reader 才会完成
      整份文本的解码验证。
    - strict=True 要求 csv 对它能够识别的格式错误（例如不合法的引号结构）
      抛出 csv.Error，而不是宽松恢复。它不会替代业务校验，也不会自动拒绝
      多余列、空值或非法路径；这些仍由字段检查和 apply/undo 的逐行逻辑负责。
    - 访问 reader.fieldnames 会读取表头。用集合差集算出缺失的必要列并排序，
      排序仅用于让错误消息稳定。此处只校验“执行所需列存在”，不强制完整等于
      当前版本的全部列，因此旧方案/日志仍有一定向前兼容空间。

    本上下文管理器只负责打开方式、表头和逐行解码。调用方若要保证“任何改名
    发生前整份文件都已通过 csv 解析”，还必须像 apply_plan/undo_log 那样在
    with 内把 reader 消费为列表；仅取得 reader 并不会预读所有数据行。
    """
    path = path.resolve()
    encoding = _detect_encoding(path)
    with path.open("r", encoding=encoding, newline="") as handle:
        sample = handle.read(65536)
        handle.seek(0)
        delimiter = _detect_delimiter(sample, path)
        reader = csv.DictReader(handle, delimiter=delimiter, strict=True)
        fieldnames = set(reader.fieldnames or [])
        missing = sorted(required_fields - fieldnames)
        if missing:
            raise ValueError(f"表格缺少必要列：{', '.join(missing)}")
        yield reader


def iter_documents(root: Path):
    """递归遍历目录，产出所有「需要处理」的文件路径（惰性生成器）。

    函数接口使用 yield，调用方可以逐项消费结果；但内部为了获得稳定顺序调用了
    sorted(root.rglob("*"))，排序会先收集所有目录条目。因此它不是严格的常量
    内存流式扫描，内存占用与目录条目数相关。create_plan 随后还会显式构造 files
    列表，以便计算 total、应用 limit 并显示 `[当前/总数]` 进度。

    - rglob("*") 递归匹配所有条目；sorted() 让处理顺序稳定可复现
      （否则不同机器、不同时刻的扫描顺序可能不同，方案文件难以对比）。
      排序是按完整路径字符串比较，因此同一目录下的文件会聚在一起。
    - 两个过滤条件：必须是文件（排除目录），且扩展名在 DISCOVERED_EXTENSIONS
      里。注意 DISCOVERED 包含「不支持但需提示」的格式，这些也会被产出，
      以便在方案里留下一行错误记录。
    - suffix.lower() 保证 .DOCX 这类大写扩展名同样被识别。
    """
    for path in sorted(root.rglob("*")):
        if path.is_file() and path.suffix.lower() in DISCOVERED_EXTENSIONS:
            yield path


def create_plan(
    root: Path,
    output: Path,
    provider: NamingProvider,
    *,
    limit: int | None = None,
    max_chars: int = 6000,
    max_pages: int = 5,
    overwrite: bool = False,
    progress: Callable[[str], None] = print,
) -> dict[str, int]:
    """扫描目录并生成改名方案文件，全程只读，不修改任何原文件。

    参数（`*` 之后的均为「仅关键字参数」，强制调用方写明参数名，
    避免位置参数传错顺序导致难以察觉的 bug）：
        root       待扫描的根目录。
        output     方案文件输出路径。
        provider   命名提供方（规则 / 模型）。
        limit      只处理前 N 个文件，用于小批量试跑验证配置是否正确。
        max_chars  每份文档抽取的最大字符数。
        max_pages  PDF/PPTX 最多读取页数。
        overwrite  是否允许覆盖已存在的方案文件。
        progress   进度回调，默认 print。设计成可注入的参数，便于测试时
                   静默收集、或在 GUI 场景改成进度条。

    返回各类文件的统计计数，供命令行打印：

        total       纳入本批次的文档总数（应用 limit 之后）；
        planned     成功获得命名建议的数量；这些行当前默认 status=approved；
        duplicates  内容哈希与更早文件重复、行状态写成 duplicate 的数量；
        errors      无可用抽取内容或命名调用失败、行状态写成 error 的数量。

    planned 是“生成阶段结果”，approved 是“方案行审批状态”。两者在默认配置下
    数量通常相等，但语义不能混用：用户可在方案生成后把 approved 改为 pending，
    create_plan 当时的 planned 历史统计不会随表格编辑变化；apply_plan 则会重新
    计算它看到的 approved 数量。

    执行流程与关键决策：
      1. 规范化 root 为绝对路径，并校验确实是目录。resolve() 顺带消除了
         `..` 和符号链接带来的路径歧义，让后续 relative_to 计算可靠。
      2. 若输出文件已存在且未指定 overwrite，抛 FileExistsError。
         这是防误覆盖：方案文件可能包含用户已审核的 status 调整和名称修改，
         被静默覆盖意味着审核工作白做，所以必须显式确认。
      3. 排除输出文件自身（`path.resolve() != output`）。方案文件若写在
         扫描目录内且扩展名恰好在支持列表（如 .txt），下一轮扫描会把它
         自己当作文档——这就是为什么要显式剔除。
      4. limit 截断在排序之后进行，保证「前 N 个」是确定的文件集合。
      5. seen_hashes 与 reserved 两个字典/集合贯穿全程，前者做重复检测，
         后者做文件名去重（详见对应位置注释）。
      6. 输出统一使用 utf-8-sig。UTF-8 BOM 能帮助中文 Windows 上的 Excel/WPS
         正确识别中文；读取时使用 utf-8-sig 会自动消费 BOM，不让它混入首列名。
         列分隔符再由输出后缀决定：.txt/.tsv 是制表符，其余是逗号。
      7. 每处理完一个文件就 handle.flush()，把 Python 用户态缓冲推给操作系统，
         让长批次中途也能看到方案增长。这里没有 fsync，也没有临时文件换名，
         所以进程中断时可能留下“表头和前若干行”的部分方案；apply_plan 的全量
         预解析与逐行校验仍会保护原文件，但用户应重新生成完整方案再执行。

    每个文件按以下优先级落入互斥的三类之一：
        duplicate —— sha256 已出现过。记为重复而不改名，因为恢复出来的
                     数据里常含大量副本，对它们逐个调用模型纯属浪费。
                     注意 sha256 为空（文件读不出）时不做重复判断，
                     否则所有读取失败的文件会因空哈希相等而被误判成互相重复。
                     setdefault 只记录首次出现的路径，因此 duplicate_of
                     指向的是「最早出现的那份」，语义清晰。
                     duplicate 判断先于抽取错误判断，因此同内容的后续副本
                     会稳定归入 duplicate，不会重复报告相同抽取问题。
        error     —— 抽取失败且没有任何可用内容（既无节选也无内嵌标题），
                     或者进入命名阶段后 provider.suggest 抛出异常。
                     后者会把错误追加到已有 extraction_error 之后，
                     用分号连接的写法保留了「先抽取失败、后调用失败」的完整链路。
        approved  —— 正常生成了建议，默认纳入后续 apply；执行前仍应人工审核。
                     同时增加的是 counts["planned"]，而不是名为 approved 的
                     生成期计数。approved 只存在于方案 status 列。
    """
    root = root.resolve()
    if not root.is_dir():
        raise ValueError(f"输入目录不存在：{root}")
    output = output.resolve()
    if output.exists() and not overwrite:
        raise FileExistsError(f"方案文件已存在：{output}；请更换名称或使用 --overwrite")
    files = [path for path in iter_documents(root) if path.resolve() != output]
    if limit is not None:
        files = files[:limit]

    seen_hashes: dict[str, str] = {}
    reserved: set[Path] = set()
    # 统计字典描述“create_plan 做出了什么”，不是方案 status 的枚举。
    # planned 表示成功拿到建议；正常行虽然默认 status=approved，这里仍必须使用
    # counts["planned"]。apply_plan 才另设 counts["approved"] 统计审核后入选的行。
    counts = {"total": len(files), "planned": 0, "duplicates": 0, "errors": 0}
    output.parent.mkdir(parents=True, exist_ok=True)
    with output.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(
            handle,
            fieldnames=PLAN_FIELDS,
            delimiter=_output_delimiter(output),
        )
        writer.writeheader()
        # start=1 让进度从 1 开始计数，对用户更自然。
        for index, path in enumerate(files, start=1):
            # 写进表格的必须是相对路径：绝对路径既冗长，又会让整个方案
            # 无法随数据目录一起搬迁（换台机器路径就全失效了）。
            relative = str(path.relative_to(root))
            progress(f"[{index}/{len(files)}] {relative}")
            document = extract_document(path, max_chars=max_chars, max_pages=max_pages)
            duplicate_of = seen_hashes.get(document.sha256, "") if document.sha256 else ""
            if document.sha256:
                seen_hashes.setdefault(document.sha256, relative)
            # 这是“行状态”的默认值：只要后续成功获得 suggestion，就原样写成
            # approved，使 apply 可以直接选中；duplicate/error 分支会覆盖它。
            # 它与上方 counts["planned"] 的统计键职责不同，不能把统计键也改成
            # approved，否则返回结构的消费者会找不到约定好的 planned 字段。
            status = "approved"
            suggestion = None
            if duplicate_of:
                status = "duplicate"
                counts["duplicates"] += 1
            elif document.extraction_error and not document.excerpt and not document.embedded_title:
                status = "error"
                counts["errors"] += 1
            else:
                try:
                    suggestion = provider.suggest(document)
                except Exception as exc:
                    # 单次模型调用失败只影响这一行，不中断整批扫描。
                    # 用分号拼接而非覆盖，保留此前的抽取错误信息。
                    document.extraction_error = (
                        f"{document.extraction_error}; " if document.extraction_error else ""
                    ) + f"模型调用失败：{type(exc).__name__}: {exc}"
                    status = "error"
                    counts["errors"] += 1
                else:
                    # 进入 else 说明 provider.suggest 正常返回，故增加“建议已生成”
                    # 的 planned 计数。统计更新特意放在 try/except 外的 else：只有
                    # provider 调用本身属于“模型调用失败”；若统计键拼错等本地 bug
                    # 发生，应立即暴露，不能被宽泛捕获后伪装成某个文件的模型错误。
                    counts["planned"] += 1

            proposed_name = ""
            if suggestion:
                # 两步走：先净化名字主干（去非法字符、防保留名），
                # 再在目标目录内查重并追加 _2、_3 后缀。
                # 注意扩展名统一转小写，避免同一批文件里出现 .DOCX 和 .docx
                # 两种写法。
                stem = sanitize_stem(suggestion.title)
                proposed_name = unique_name(path.parent, f"{stem}{path.suffix.lower()}", reserved)
            writer.writerow(
                _safe_row(
                    {
                        "status": status,
                        "original_relpath": relative,
                        "original_name": path.name,
                        "proposed_name": proposed_name,
                        # 没有建议时写空串而不是 None，保证表格列类型一致。
                        "confidence": suggestion.confidence if suggestion else "",
                        "category": suggestion.category if suggestion else "",
                        "reason": suggestion.reason if suggestion else "",
                        "sha256": document.sha256,
                        "duplicate_of": duplicate_of,
                        "embedded_title": document.embedded_title,
                        "extraction_method": document.extraction_method,
                        "extraction_error": document.extraction_error,
                        "size": document.size,
                        "modified_at": document.modified_at,
                    }
                )
            )
            # 这里只要求方案内容尽快对用户可见；flush 不等于物理落盘。真正会
            # 驱动撤销的执行日志在 apply_plan 中还会额外调用 os.fsync。
            handle.flush()
    return counts


def apply_plan(root: Path, plan: Path, log_path: Path, *, overwrite: bool = False) -> dict[str, int]:
    """执行方案中 status=approved 的改名操作，并写入可撤销的日志。

    这是本工具唯一会修改用户文件的函数，因此安全措施最密集：
        1. 日志文件已存在时默认拒绝执行（除非 --overwrite），避免覆盖掉
           上一次的撤销依据。
        2. 先把整个方案完整解码并按 CSV/TSV 语法解析成列表，再开始改名。
        3. 对每个 approved 行重新校验名称、扩展名、路径边界、存在性和哈希。
        4. 每次 approved 尝试写完日志后立即 flush + fsync，缩小崩溃窗口。

    返回 {"approved", "renamed", "skipped", "failed"} 四项计数：

        approved  从人工编辑后的方案中选中、准备尝试执行的行数；
        renamed   通过全部校验并成功改名的行数；
        failed    已选中但校验或改名失败的行数；
        skipped   status 不是 approved，因而完全未尝试的行数。

    正常结束时 renamed + failed == approved，approved + skipped == 方案数据行数。
    这里的 approved 与 create_plan 返回的 planned 不同：planned 记录生成建议成功，
    approved 记录方案经过人工编辑后，本次 apply 实际选择了多少行。

    第 2 点值得展开：如果一边从磁盘读取方案、一边改名，直到第 50 行才发现
    引号没有闭合或后半段字节不是当前编码，前 49 个文件可能已经改完。这里先
    消费完整个 DictReader，把“文件级编码错误、分隔符/表头错误、csv 语法错误”
    尽量在第一次 rename 之前暴露出来。方案通常只有几千行，放入内存的成本
    很小。这个阶段是“语法预解析”而不是事务，也不表示每行业务值已经验证：
    名称、路径、文件当前状态和哈希仍必须紧邻每次 rename 重新检查。

    为什么日志用 flush + os.fsync：
        flush() 只把 Python 文本/字节缓冲交给操作系统；fsync() 再要求系统把该
        文件描述符此前的写入同步到稳定存储。每次 approved 尝试（成功或失败）
        都写一行后执行这两步，因此“已经成功完成 fsync 的日志行”具有更强的
        崩溃持久性。方案生成时的 flush 只为可见性，不能提供同等级保证。

        这仍不是跨文件系统操作和日志文件的原子事务：代码顺序是先 rename，
        再 writerow -> flush -> fsync。进程若恰好在 rename 成功后、日志 fsync
        完成前终止，仍可能出现文件已改名但日志行缺失或未持久化的极短窗口；
        日志写入自身失败也会让函数中止。fsync 的作用是缩小并明确持久性边界，
        不能把两个独立文件系统动作变成全有或全无的数据库事务。

    对 status=approved 的正常结构行，校验链如下；链内异常会记为 failed 并继续：
        a. proposed_name 必须非空，且 Path(name).name == name。
           如果字符串含目录分隔符或本身是路径（如 `../x.txt`），取 .name 后
           通常只剩末尾文件名，与原文不等，从而拒绝把 proposed_name 当路径。
        b. 扩展名必须与原文件一致。防止方案被改错后把 .docx 改成 .exe 之类。
        c. 源、目标路径分别经 safe_resolve_below 解析 `..` 和符号链接，并确认
           规范化后的结果仍位于 root 内。即使某一列被恶意修改，任一路径越界
           都会在真正访问文件前失败。
        d. 源文件必须存在、目标不能已存在（避免静默覆盖他人文件）。
        e. 重新计算源文件哈希并与方案中的 sha256 比对。这是最后一道防线：
           如果文件在「生成方案」和「执行改名」之间被修改过（比如用户又
           用恢复工具导出覆盖了一次），说明方案里的判断已经过时，
           此时必须拒绝执行。哈希核对的是文件字节内容，不核对修改时间、权限
           等元数据，也不是对并发修改的文件锁。
        f. 全部通过后才调用 source.rename(target)。
           本场景中源和目标只改同一父目录的文件名，通常由文件系统作为单次
           rename 完成，不需要复制文件内容。目标存在检查仍必须显式执行，
           既避免覆盖，也让跨平台行为一致可理解。

    存在性检查、哈希计算和 rename 之间没有持有文件系统锁；若另一个进程同时
    修改目录，仍存在检查后状态变化（TOCTOU）的可能。本工具面向人工控制的本地
    批处理，因此以“执行期间不要让其他程序改动目标目录”为运行前提。

    日志语义：
      - 只为 approved 行写日志；pending、duplicate、error 等行只计入 skipped；
      - before_relpath 总是来自方案；after_relpath 在名称/扩展名校验前为空，
        一旦候选路径算出便会记录，所以某些 failed 行也可能带 after_relpath；
      - status=renamed 才表示文件实际改名成功，failed 行绝不能仅凭非空的
        after_relpath 判断为成功；
      - sha256 记录方案中的期望值，message 在失败时记录异常类型和说明。
    """
    root = root.resolve()
    counts = {"approved": 0, "renamed": 0, "skipped": 0, "failed": 0}
    log_path = log_path.resolve()
    if log_path.exists() and not overwrite:
        raise FileExistsError(f"日志文件已存在：{log_path}；请更换名称或使用 --overwrite")
    log_path.parent.mkdir(parents=True, exist_ok=True)

    # 在修改任何文件前消费完整个 reader。这样后半段的解码/csv 语法错误会在
    # rename 前暴露；名称、路径、存在性、哈希等业务校验仍按 approved 行执行。
    with _open_dict_reader(plan, PLAN_REQUIRED_FIELDS) as reader:
        rows = [_restore_row(row) for row in reader]

    with log_path.open("w", encoding="utf-8-sig", newline="") as log_handle:
        writer = csv.DictWriter(
            log_handle,
            fieldnames=LOG_FIELDS,
            delimiter=_output_delimiter(log_path),
        )
        writer.writeheader()
        for row in rows:
            # 人工审核的约定：只处理 status 明确为 approved 的行。create_plan
            # 默认写 approved，但用户可以在此之前改为 pending；duplicate/error
            # 也会自然落入 skipped。这里统计的是 apply 看到的最终审批结果，
            # 与 create_plan 曾返回多少 planned 无关。
            # 用 lower() + strip() 容忍表格软件保存时的多余空格和大小写差异
            # （用户手填 "Approved " 也能生效）。
            if row.get("status", "").strip().lower() != "approved":
                counts["skipped"] += 1
                continue
            counts["approved"] += 1
            before_rel = row["original_relpath"]
            after_rel = ""
            timestamp = datetime.now(timezone.utc).isoformat()
            # 先假定失败，成功后再改写为 renamed。这种「悲观默认值」的做法
            # 保证任何未预料到的异常路径都不会把失败误记成成功。
            status = "failed"
            message = ""
            try:
                proposed_name = row.get("proposed_name", "").strip()
                if not proposed_name or Path(proposed_name).name != proposed_name:
                    raise ValueError("proposed_name 必须是单独的文件名，不能包含目录")
                if Path(proposed_name).suffix.lower() != Path(before_rel).suffix.lower():
                    raise ValueError("建议名称必须保留原文件扩展名")
                # with_name 只替换文件名部分，保留原目录结构；随后源和目标都
                # 独立经过 safe_resolve_below，不能只校验其中一边。
                after_rel = str(Path(before_rel).with_name(proposed_name))
                source = safe_resolve_below(root, before_rel)
                target = safe_resolve_below(root, after_rel)
                if not source.is_file():
                    raise FileNotFoundError(f"源文件不存在：{source}")
                if target.exists():
                    raise FileExistsError(f"目标文件已存在：{target}")
                actual_hash = sha256_file(source)
                if actual_hash != row["sha256"]:
                    raise ValueError("文件哈希与方案不一致，文件可能已被修改")
                source.rename(target)
                status = "renamed"
                counts["renamed"] += 1
            except Exception as exc:
                # 记录异常类型名很重要：FileExistsError 和 ValueError 对应的
                # 处理方式完全不同（前者改个名字重试，后者要重新生成方案）。
                message = f"{type(exc).__name__}: {exc}"
                counts["failed"] += 1
            writer.writerow(
                _safe_row(
                    {
                        "timestamp": timestamp,
                        "status": status,
                        "before_relpath": before_rel,
                        "after_relpath": after_rel,
                        "sha256": row["sha256"],
                        "message": message,
                    }
                )
            )
            log_handle.flush()
            # fsync 接收操作系统文件描述符，要求把 writerow 和 flush 已交付的
            # 数据同步到稳定存储。它增强“本行已 fsync 之后”的持久性，但不能
            # 消除上方 rename 成功到本次 fsync 完成之间的窗口，也不能让改名与
            # 日志写入成为同一个原子事务。逐行同步有性能成本，换来更小的风险窗。
            os.fsync(log_handle.fileno())
    return counts


def undo_log(root: Path, log_path: Path) -> dict[str, int]:
    """依据改名日志，把所有已完成的改名还原。

    这是数据恢复场景的「后悔药」，也是整个工具敢于自动化改名的底气所在。

    undo 的数据源是 apply 写出的日志，而不是当前方案：方案后来即使被编辑，
    也不能改变历史上实际完成了哪些 rename。日志会先完整解码/解析到内存，
    避免后半段格式错误出现时已经开始撤销；公式防护前缀也会在路径校验前还原。

    关键设计：
    - **逆序遍历**（`reversed(rows)`）。日志按执行顺序写入，撤销按后进先出
      处理。若后一次改名使用了前一次操作刚腾出的名称，先撤销较新的操作才能
      释放路径，再撤销较旧操作。例如执行顺序是 B->C、随后 A->B，撤销顺序
      必须是 B->A、随后 C->B。即使常规方案生成器会尽量避免名称冲突，日志
      仍按操作栈语义处理人工编辑或其他合法依赖。
    - 只处理 status **精确等于** "renamed" 的行。apply 自己写出的值是规范
      小写，因此 undo 不像审核入口那样 strip/lower；failed 或被手工改成其他
      状态的行均计入 skipped，非空 after_relpath 不能替代成功状态判断。
    - 两条路径都经 safe_resolve_below 约束在 root 内。撤销目标（原名称）必须
      尚未被占用，当前名称必须是普通文件，且内容 SHA-256 必须等于执行方案时
      记录的期望哈希。哈希只证明字节内容一致，不证明权限、时间戳等元数据；
      undo 也只恢复名称/路径，不恢复元数据。
    - 任一检查或 rename 抛异常时，只增加 failed 并继续处理更早记录。这样某个
      文件被删除、手工改名或内容变化不会阻断其他独立文件；代价是返回值只给
      失败数量，不保存逐项撤销错误详情。
    - undo 不覆盖已有文件、不带“强制撤销”选项，也不生成第二份撤销日志。
      因此它本身不是可再次 undo 的事务。若需要重新应用名称，可在当前文件
      状态仍满足原方案校验时重新运行 apply，并使用新的日志路径。

    返回计数：restored 是成功恢复原路径的 renamed 行，skipped 是日志中所有
    非 renamed 行，failed 是本应撤销但未通过校验或改名失败的行。正常遍历完成
    时三者之和等于日志数据行数。

    与 apply 一样，这不是数据库事务：撤销没有新日志，也没有逐步 fsync；进程
    中断可能留下部分已撤销、部分尚未撤销的状态。重新运行同一日志时，已经恢复
    的项目通常会因“原名称已被占用”计为 failed，其余项目仍会继续尝试，用户
    应结合目录现状核对最终结果，而不能把 failed=0 当作跨多次运行的唯一标准。
    """
    root = root.resolve()
    with _open_dict_reader(log_path, LOG_REQUIRED_FIELDS) as reader:
        rows = [_restore_row(row) for row in reader]
    counts = {"restored": 0, "skipped": 0, "failed": 0}
    # 日志按 apply 的执行顺序排列；从尾到头相当于按栈回退，能先释放较晚操作
    # 占用的名称。注意这里反转的是全量日志，非 renamed 行随后会被跳过。
    for row in reversed(rows):
        # 与 apply 对人工输入的 approved 判断不同，这里读取程序自己生成的审计
        # 状态，要求精确匹配 renamed；failed 行即使 after_relpath 非空也没改名。
        if row.get("status") != "renamed":
            counts["skipped"] += 1
            continue
        try:
            # before/after 都来自可编辑的文本日志，必须分别做 root 边界校验。
            # 变量名从“撤销方向”理解：renamed 是当前源，original 是恢复目标。
            original = safe_resolve_below(root, row["before_relpath"])
            renamed = safe_resolve_below(root, row["after_relpath"])
            if original.exists():
                raise FileExistsError(f"原名称已被占用：{original}")
            if not renamed.is_file():
                raise FileNotFoundError(f"改名后的文件不存在：{renamed}")
            # 防止把改名后又被编辑/替换的文件搬回旧名称。该检查验证内容，
            # 不负责恢复修改时间、权限等文件系统元数据。
            if sha256_file(renamed) != row["sha256"]:
                raise ValueError("文件哈希不匹配，不执行撤销")
            renamed.rename(original)
            counts["restored"] += 1
        except Exception:
            # 继续回退其余记录，但当前接口只返回计数，不另写撤销错误日志。
            counts["failed"] += 1
    return counts
