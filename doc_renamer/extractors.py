"""文档内容抽取层：为“根据内容建议文件名”提取有限的文字线索。

本模块不是完整的 Office/PDF 渲染器，也不承诺还原原始版式。它把一个磁盘文件
转换成 :class:`types.DocumentContent`，供命名规则或模型使用；输出只包含文件
元数据、可选的内嵌标题、经过压缩截断的正文节选、抽取方式和错误说明。

支持范围与实际读取策略：
    1. OOXML Office：``.docx/.docm``、``.xlsx/.xlsm``、``.pptx/.pptm``。
       这些格式是 ZIP 容器，本模块只用 ``zipfile`` 和 ``ElementTree`` 读取其中
       指定的 XML 部件：Word 的正文及数字编号页眉/页脚、Excel 的工作表名称和
       部分单元格存储值、PowerPoint 的前 ``max_pages`` 个幻灯片文本，以及可用
       时的核心属性标题。它不渲染版式，不展开嵌入对象或外部链接。
    2. PDF：用 pypdf 读取 ``/Title`` 元数据和前若干页已有的文字层；不做 OCR，
       因而纯扫描 PDF 通常只能得到明确的抽取错误。
    3. 文本类：``txt/md/csv/tsv/json/xml/html/htm/log/rtf`` 只按字节读取并启发式
       猜测编码。CSV、TSV、JSON 不做结构化解析；HTML、XML、RTF 也不会渲染，
       只用简单正则削减标记后生成节选。

安全边界：这里从不启动 Office/PDF 应用，也不调用宏、公式计算引擎或脚本引擎。
``.docm/.xlsm/.pptm`` 中的 VBA 二进制部件不会被读取或执行；Excel 公式不会求值，
最多读取文件中已经保存的 ``<v>`` 值。这里的“支持”仅表示可尝试读取上述静态文字，
不表示文件内主动内容安全或文档已被完整解析。

旧版二进制 Office（``.doc/.xls/.ppt``）和图片会被扫描发现，但不会尝试解析；
入口函数把原因写入 ``DocumentContent.extraction_error``，避免静默遗漏。通常的
格式解析异常也会降级为单文件错误，让批处理继续；但该保证属于
``extract_document()`` 公共入口，直接调用以下私有抽取器仍可能收到异常。
"""

from __future__ import annotations

import re
import zipfile
from datetime import datetime, timezone
from pathlib import Path
from xml.etree import ElementTree as ET

from .types import DocumentContent
from .utils import compact_text, sha256_file


# 可尝试抽取静态文字的扩展名。宏格式与对应 OOXML 格式共享本模块需要的 XML
# 文字部件，所以复用解析器；这不表示两种格式的全部包结构完全相同。解析器不会
# 打开、解释或执行 vbaProject.bin 等宏部件。
SUPPORTED_EXTENSIONS = {
    ".docx",
    ".docm",
    ".xlsx",
    ".xlsm",
    ".pptx",
    ".pptm",
    ".pdf",
    ".txt",
    ".md",
    ".csv",
    ".tsv",
    ".json",
    ".xml",
    ".html",
    ".htm",
    ".log",
    ".rtf",
}

# 明确不支持但仍需“被看见”的扩展名。扫描器把它们纳入结果，由公共入口返回
# 带 extraction_error 的记录；当前既没有旧版 Office 二进制解析器，也没有图片
# OCR。这样方案中能留下待人工处理的 error 行，而不会因扩展名不支持而静默漏项。
KNOWN_UNSUPPORTED_EXTENSIONS = {
    ".doc",
    ".xls",
    ".ppt",
    ".jpg",
    ".jpeg",
    ".png",
    ".bmp",
    ".tif",
    ".tiff",
}

# 扫描阶段使用此并集：可解析文件和需要显式报错的文件都进入结果。它不是
# “所有可能文档扩展名”的清单；未列出的扩展名不会被常规目录扫描发现。
DISCOVERED_EXTENSIONS = SUPPORTED_EXTENSIONS | KNOWN_UNSUPPORTED_EXTENSIONS


def _xml_text(data: bytes, tag_suffix: str | None = None) -> str:
    """收集一段 XML 中符合条件的节点文本，并用换行拼接。

    ``data`` 是完整 XML 部件的原始字节，交给 ``ET.fromstring`` 一次性解析；
    XML 不完整、命名空间声明损坏或内容不合法时会抛 ``ParseError``，由公共入口
    统一降级为 ``extraction_error``。这里不做 XML 修复，也不会解析实体以外的
    文档部件关系。

    ``tag_suffix`` 是 ElementTree 展开后的标签名后缀过滤器。OOXML 的文本标签
    通常呈 ``{命名空间 URI}t``，调用方传 ``"}t"`` 可以兼容 Transitional/Strict
    等不同命名空间 URI，而无需写死完整 URI；传 ``None`` 则收集所有具有
    ``node.text`` 的节点。后缀匹配只按局部名称筛选，不校验该 ``t`` 标签究竟
    属于哪个 OOXML 命名空间，因此这是面向文字线索的宽松读取，不是模式验证。

    ``root.iter()`` 按 XML 文档顺序遍历全部后代。每个文本节点之间直接插入换行，
    不尝试恢复段落、表格单元格、空格保留规则或视觉布局；随后调用方还会通过
    ``compact_text`` 把空白压缩，所以返回值适合命名分析，不适合内容归档。
    """
    root = ET.fromstring(data)
    values: list[str] = []
    for node in root.iter():
        if node.text and (tag_suffix is None or node.tag.endswith(tag_suffix)):
            values.append(node.text)
    return "\n".join(values)


def _core_title(archive: zipfile.ZipFile) -> str:
    """尽力读取 OOXML ``docProps/core.xml`` 中的 Dublin Core 标题。

    返回值是压缩并截断到 200 字符的标题；找不到核心属性部件、核心属性 XML
    损坏、没有 ``{...}title`` 节点或标题为空时都返回空字符串。这里只把
    ``KeyError`` 和 ``ET.ParseError`` 视为“无可用标题”，ZIP 容器本身损坏等
    其他异常仍交给上层记录，避免把真正的文件损坏误装成普通的标题缺失。

    标题是作者/软件保存的元数据，通常是高价值命名线索，但并不保证可信、最新，
    也不保证与正文一致。函数只读取 XML 文本，不跟随外部关系、不加载自定义属性，
    更不会执行宏。使用 ``endswith("}title")`` 是为了兼容命名空间 URI 差异，
    同时意味着这里只做宽松的局部名称匹配。
    """
    try:
        root = ET.fromstring(archive.read("docProps/core.xml"))
    except (KeyError, ET.ParseError):
        return ""
    for node in root.iter():
        if node.tag.endswith("}title") and node.text:
            return compact_text(node.text, 200)
    return ""


def _extract_docx(path: Path, max_chars: int, max_pages: int) -> tuple[str, str, str]:
    """抽取 ``.docx/.docm`` 的核心属性标题及有限的 WordprocessingML 文本。

    返回 ``(embedded_title, excerpt, "docx-xml")``。正文来源限定为
    ``word/document.xml``，并追加名称严格匹配 ``word/header数字.xml`` 或
    ``word/footer数字.xml`` 的页眉页脚；关系文件不会误入。脚注、尾注、批注、
    文档属性之外的元数据以及独立嵌入对象不会另行读取。正文部件内部所有局部名为
    ``t`` 的节点都会按 XML 顺序收集，因此普通段落、表格和某些文本框中的文字
    可能出现，但原有段落/表格结构不会保留。

    Word 的分页是排版结果，并不直接对应这里读取的 XML 节点，所以
    ``max_pages`` 被有意忽略；``max_chars`` 在合并正文和页眉页脚后通过
    ``compact_text`` 生效。若恢复出的 ZIP 缺少 ``word/document.xml``，代码会
    跳过该部件而继续尝试页眉页脚；若最终标题和正文都为空，公共入口再记录错误。

    对 ``.docm`` 也只读取上述静态 XML。VBA、OLE 对象、字段代码及外部关系不会
    被运行；本函数从不启动 Word/WPS。抽取成功仅说明拿到了文字线索，不代表文档
    完整、无恶意内容或能够正常由 Office 打开。
    """
    del max_pages
    with zipfile.ZipFile(path) as archive:
        title = _core_title(archive)
        parts = ["word/document.xml"]
        parts.extend(
            name
            for name in archive.namelist()
            if re.fullmatch(r"word/(header|footer)\d+\.xml", name)
        )
        # 恢复出的 OOXML 可能缺少主文档部件。存在性检查允许继续读取尚存的
        # 页眉/页脚；它只处理“条目缺失”，ZIP 或 XML 损坏仍由公共入口记录。
        text = "\n".join(
            _xml_text(archive.read(name), "}t") for name in parts if name in archive.namelist()
        )
    return title, compact_text(text, max_chars), "docx-xml"


def _extract_xlsx(path: Path, max_chars: int, max_pages: int) -> tuple[str, str, str]:
    """抽取 ``.xlsx/.xlsm`` 的标题、工作表名和一部分已存储单元格值。

    返回 ``(embedded_title, excerpt, "xlsx-xml")``。处理范围如下：

    1. ``xl/sharedStrings.xml`` 存在时，先建立共享字符串列表。每个 ``<si>``
       条目可能包含多个富文本 ``<t>`` 片段；这里只拼接文字，不保留字体、颜色、
       换行样式等格式。``t="s"`` 的单元格把 ``<v>`` 当作非负下标回查该列表；
       非数字或越界下标不会回查，而会保留原始 ``<v>`` 文本作为线索。
    2. ``xl/workbook.xml`` 存在时，读取全部 ``<sheet name="...">`` 名称并放在
       节选开头。这里不解析 relationship 来核对每个名称对应哪个 sheet XML，
       工作表 XML 则按 ``sheet数字.xml`` 的数字文件名稳定排序读取。因此结果
       适合判断工作簿主题，不应当被理解为精确还原工作簿标签顺序和映射。
    3. ``t="inlineStr"`` 的单元格拼接其全部 ``<t>``；其他类型直接采用第一个
       ``<v>`` 的原始文本。日期/时间可能仍是序列号，布尔值可能仍是 ``0/1``，
       数字格式、显示格式、批注、图表、数据透视表、隐藏状态和合并关系均不解析。
    4. 公式 ``<f>`` 从不解释或执行；若文件保存了缓存结果，只会读取相邻
       ``<v>`` 的静态值，缓存缺失时公式文本本身也不会进入节选。这意味着结果
       可能过期，与用户在 Excel/WPS 中重新计算后看到的内容不同。

    ``max_pages`` 对工作簿没有定义，故有意忽略。收集值达到 ``max_chars`` 后停止
    遍历后续单元格/工作表，最后再压缩截断；但每个被访问的 worksheet 仍由
    ``ET.fromstring`` 一次性载入内存，所以这不是流式 XLSX 解析器，超大单张表
    仍可能占用较多内存。

    对 ``.xlsm``，VBA 二进制部件不会被读取或执行；本函数也不启动 Excel/WPS、
    不刷新外部数据连接。所有输出只是 ZIP 内现存 XML 的静态快照。
    """
    del max_pages
    with zipfile.ZipFile(path) as archive:
        title = _core_title(archive)
        names = set(archive.namelist())
        shared: list[str] = []
        if "xl/sharedStrings.xml" in names:
            root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
            for item in root:
                shared.append("".join(node.text or "" for node in item.iter() if node.tag.endswith("}t")))

        sheet_names: list[str] = []
        if "xl/workbook.xml" in names:
            root = ET.fromstring(archive.read("xl/workbook.xml"))
            sheet_names = [node.attrib.get("name", "") for node in root.iter() if node.tag.endswith("}sheet")]

        values: list[str] = []
        for sheet_path in sorted(name for name in names if re.fullmatch(r"xl/worksheets/sheet\d+\.xml", name)):
            root = ET.fromstring(archive.read(sheet_path))
            for cell in (node for node in root.iter() if node.tag.endswith("}c")):
                cell_type = cell.attrib.get("t")
                # 只查看单元格的直接子节点并取第一个 <v>；不读取 <f> 公式文本，
                # 不按 styles.xml 转换日期或显示格式。没有缓存/原始值时为空串。
                raw = next((node.text or "" for node in cell if node.tag.endswith("}v")), "")
                if cell_type == "s" and raw.isdigit() and int(raw) < len(shared):
                    raw = shared[int(raw)]
                elif cell_type == "inlineStr":
                    raw = "".join(node.text or "" for node in cell.iter() if node.tag.endswith("}t"))
                if raw:
                    values.append(raw)
                if sum(len(v) for v in values) >= max_chars:
                    break
            if sum(len(v) for v in values) >= max_chars:
                break
    combined = "工作表：" + "、".join(sheet_names) + "\n" + "\n".join(values)
    return title, compact_text(combined, max_chars), "xlsx-xml"


def _extract_pptx(path: Path, max_chars: int, max_pages: int) -> tuple[str, str, str]:
    """抽取 ``.pptx/.pptm`` 的核心属性标题和有限的幻灯片 XML 文本。

    返回 ``(embedded_title, excerpt, "pptx-xml")``。候选部件只包括名称严格
    匹配 ``ppt/slides/slide数字.xml`` 的文件；先从文件名提取数字作自然排序，
    再取前 ``max_pages`` 个。自然排序避免 ``slide10`` 排在 ``slide2`` 前，但
    幻灯片的实际放映顺序由 ``presentation.xml`` 及关系文件决定，本实现没有
    解析这些关系。因此“前若干页”准确地说是编号最小的若干部件，不能保证与
    用户最后保存的放映顺序完全一致。

    每个选中部件只收集局部名为 ``t`` 的 XML 节点，并在合并后按 ``max_chars``
    压缩截断。备注页、母版/版式、批注、图表工作簿、SmartArt 独立数据、媒体、
    附件和替代文字不会专门读取；视觉位置和格式也不会保留。

    对 ``.pptm``，VBA 项目及动作不会被加载或执行；本函数从不启动
    PowerPoint/WPS，也不访问外部链接。宏格式在此仅复用静态 slide XML 读取。
    """
    with zipfile.ZipFile(path) as archive:
        title = _core_title(archive)
        slide_paths = sorted(
            (name for name in archive.namelist() if re.fullmatch(r"ppt/slides/slide\d+\.xml", name)),
            key=lambda name: int(re.search(r"(\d+)", Path(name).stem).group(1)),
        )[:max_pages]
        text = "\n".join(_xml_text(archive.read(name), "}t") for name in slide_paths)
    return title, compact_text(text, max_chars), "pptx-xml"


def _extract_pdf(path: Path, max_chars: int, max_pages: int) -> tuple[str, str, str]:
    """用 pypdf 抽取 PDF 元数据标题和已有文字层，不进行 OCR 或渲染。

    pypdf 在函数内部延迟导入，使依赖导入失败只影响 PDF，而不会阻止其他格式
    使用本模块；当前项目安装配置虽已声明 pypdf，运行环境损坏或依赖缺失时仍会
    由 ``extract_document`` 捕获异常并写入 ``extraction_error``。

    标题只读 ``reader.metadata`` 中的 ``/Title``，为空时使用空字符串；不会读取
    XMP 等其他元数据源。正文按 ``reader.pages`` 的文档页序处理，最多访问前
    ``max_pages`` 页，并在累计文字达到 ``max_chars`` 后提前停止；最终节选仍由
    ``compact_text`` 统一压缩、截断。pypdf 的 ``extract_text`` 读取 PDF 内容流
    中可解释的文字，不等同于视觉 OCR：图片里的字、手写内容以及没有正确字符
    映射的字体可能完全缺失，复杂多栏版式的阅读顺序也可能与画面不同。

    只要 PDF 至少有一页而所访问页面没有抽到任何文字，本函数就显式抛错，提示
    后续 OCR；由于三元组尚未返回，此时上层也不会保留已读取的 ``/Title``。
    零页 PDF 不在此处抛该特定错误，之后由公共入口依据标题/正文是否都为空判断。

    本代码只调用元数据访问和 ``extract_text``，不会执行 PDF JavaScript、动作、
    表单逻辑、附件或外部链接。加密、损坏、权限限制等读取异常原样交给公共入口
    降级为该文件的错误记录。
    """
    from pypdf import PdfReader

    reader = PdfReader(str(path))
    metadata = reader.metadata or {}
    title = compact_text(str(metadata.get("/Title") or ""), 200)
    chunks: list[str] = []
    for page in reader.pages[:max_pages]:
        chunks.append(page.extract_text() or "")
        # 同样边抽边算长度，达到上限就提前停止，避免解析剩余页面浪费时间。
        if sum(len(chunk) for chunk in chunks) >= max_chars:
            break
    text = compact_text("\n".join(chunks), max_chars)
    if not text and reader.pages:
        raise ValueError("PDF 未提取到文字，可能是扫描件；后续需要接入 OCR")
    return title, text, "pypdf"


def _decode_text(data: bytes) -> str:
    """按固定优先级启发式解码字节流；这不是可靠的字符集检测器。

    文件本身未必声明编码，因此依次尝试 ``utf-8-sig``、``gb18030``、``utf-16``、
    ``big5``，第一个没有抛 ``UnicodeDecodeError`` 的结果立即返回：

    * ``utf-8-sig`` 严格校验 UTF-8，并在存在 UTF-8 BOM 时去掉开头的
      ``U+FEFF``；普通无 BOM UTF-8 也能正常解码。
    * ``gb18030`` 覆盖常见简体中文 Windows 文本及 GBK 字符，但它能接受的字节
      组合很多，某些本来属于 Big5 或无 BOM UTF-16 的数据可能在这里被误判。
    * ``utf-16`` 支持 BOM 指示的字节序；无 BOM 时的默认字节序取决于 Python
      解码器/运行平台。它对许多偶数字节序列都可能“成功”，所以排在中文编码
      之后。这也意味着无 BOM UTF-16 并不保证被正确识别。
    * ``big5`` 是繁体中文环境的最后一个严格候选。

    若四种严格解码都失败，最后以 UTF-8 ``errors="replace"`` 返回，把非法片段
    替换为 ``�``。该降级保证通常还能获得部分命名线索，但乱码不会被另行标记为
    ``extraction_error``；调用方不能把“成功返回字符串”理解为编码判断正确。
    """
    for encoding in ("utf-8-sig", "gb18030", "utf-16", "big5"):
        try:
            return data.decode(encoding)
        except UnicodeDecodeError:
            continue
    return data.decode("utf-8", errors="replace")


def _extract_text(path: Path, max_chars: int, max_pages: int) -> tuple[str, str, str]:
    r"""处理文本类文件：截取字节前缀、猜编码、粗略去标记并压缩文本。

    返回 ``("", excerpt, "plain-text")``；这里没有标题元数据协议，首行也不会
    被提升为 ``embedded_title``。``txt/md/csv/tsv/json/log`` 均作为普通文本，
    不解析 Markdown、表格分隔符或 JSON 结构，因此字段顺序、引号和标点会直接
    成为节选的一部分。

    当前表达式 ``path.read_bytes()[: max_chars * 4]`` 会先把整个文件读入内存，
    再保留前 ``max_chars * 4`` 个字节；它限制后续解码/处理量，却不限制读取时的
    峰值内存。切片也可能落在多字节字符中间，使首选编码严格解码失败并转入后续
    编码，或在最终兜底中产生 ``�``。因此这只是针对常见 UTF-8 文本的成本折中，
    不是精确保证“至少获得 max_chars 个字符”的流式读取策略。

    ``html/htm/xml/rtf`` 额外应用一个宽松正则：``<[^>]+>`` 删除形似标签的片段，
    ``\\[a-z]+\d* ?`` 删除一部分 RTF 控制字，``[{}]`` 删除 RTF 分组花括号。
    它不会解析 DOM、CDATA、HTML 实体、CSS、脚本内容或完整 RTF 转义规则；例如
    ``<script>`` 标签可能被删掉而其文本内容仍会保留。该步骤只为减少命名提示中的
    标记噪声，不是 HTML/RTF 安全清洗器，也不应拿来生成可重新发布的正文。

    文件始终按静态字节读取：不会渲染 HTML/XML，不会执行脚本、加载网络资源或
    计算任何表达式。最后 ``compact_text`` 把所有空白压成单个空格，并把结果
    限制到 ``max_chars`` 个 Python 字符。
    """
    del max_pages
    data = path.read_bytes()[: max_chars * 4]
    text = _decode_text(data)
    if path.suffix.lower() in {".html", ".htm", ".xml", ".rtf"}:
        text = re.sub(r"<[^>]+>|\\[a-z]+\d* ?|[{}]", " ", text, flags=re.IGNORECASE)
    return "", compact_text(text, max_chars), "plain-text"


def extract_document(path: Path, max_chars: int = 6000, max_pages: int = 5) -> DocumentContent:
    """抽取单个文件并把常见读取/解析失败降级为 ``DocumentContent`` 错误。

    参数契约：
        path
            待处理路径，原样保存在结果中；本函数不会把它自动 ``resolve``，也不
            锁定文件。调用期间文件若被其他进程修改，``stat``、哈希和正文可能
            来自略有差异的时刻，调用方应把 SHA-256 当作后续校验依据。
        max_chars
            返回 ``excerpt`` 的字符上限。它不是 token 上限，也不会限制 ZIP/XML
            部件解析本身的内存；默认 6000 只是给命名模型准备节选的经验值。
        max_pages
            PDF 最多访问的文档页数；PPTX/PPTM 最多访问的数字编号 slide 部件数。
            DOCX/DOCM 和 XLSX/XLSM 明确忽略该参数。

    处理顺序与错误降级：
      1. 先 ``stat`` 并读取整个文件计算 SHA-256。任一步抛出普通 ``Exception``，
         立即返回 ``sha256=""``、``size=0``、``modified_at=""`` 且带
         ``extraction_error`` 的对象；即使 ``stat`` 已成功，失败分支也统一使用
         这些哨兵值。空哈希使工作流不会把多个不可读文件误判成重复副本。
      2. 元数据成功后，记录字节大小和 UTC ISO 8601 修改时间。已知不支持的旧版
         Office/图片此时返回对应错误；也就是说它们仍会被完整哈希，但不会解析
         内容或尝试启动外部应用。
      3. OOXML/PDF 扩展名使用专用函数；其余直接调用本入口的路径默认按普通文本
         处理。常规目录扫描只传 ``DISCOVERED_EXTENSIONS`` 中的文件，所以这个
         默认分支主要覆盖已列出的文本类型；它不是未知二进制格式的自动检测器。
      4. 私有抽取器成功返回后，三元组一次性写入 ``embedded_title``、``excerpt``
         和 ``extraction_method``。标题和节选同时为空时，记录“无可用文字”；
         这表示命名所需信息不足，不等同于证明文件损坏。
      5. 私有抽取器抛出的普通 ``Exception`` 被转成“异常类型: 消息”。因为三元组
         赋值尚未完成，三个内容字段保持默认空值，而文件哈希、大小和修改时间仍
         保留；其他文件可以继续处理。

    “降级而不中断”只覆盖上述显式 ``except Exception`` 区域。Python 的
    ``KeyboardInterrupt``、``SystemExit`` 等进程控制异常不会被吞掉；参数类型
    错误或极端时间戳等发生在保护区之外的异常也不在此契约内。因此调用方应把它
    理解为面向正常文件 I/O 与格式解析故障的批处理容错，而非绝对不抛异常。
    """
    try:
        stat = path.stat()
        digest = sha256_file(path)
    except Exception as exc:
        return DocumentContent(
            path=path,
            sha256="",
            size=0,
            modified_at="",
            extraction_error=f"文件读取失败：{type(exc).__name__}: {exc}",
        )
    result = DocumentContent(
        path=path,
        sha256=digest,
        size=stat.st_size,
        # stat.st_mtime 是 POSIX 时间戳；显式用 UTC 转为带时区偏移的 ISO 8601
        # 文本，便于写入方案并跨机器比较。它记录的是哈希前取得的 stat 值，
        # 本函数不会锁文件来保证元数据与随后读取的正文完全原子一致。
        modified_at=datetime.fromtimestamp(stat.st_mtime, tz=timezone.utc).isoformat(),
    )
    if path.suffix.lower() in KNOWN_UNSUPPORTED_EXTENSIONS:
        if path.suffix.lower() in {".doc", ".xls", ".ppt"}:
            result.extraction_error = "旧版二进制 Office 格式暂不支持，请先另存为新版格式"
        else:
            result.extraction_error = "图片或扫描件暂未接入 OCR"
        return result
    extractor = {
        ".docx": _extract_docx,
        ".docm": _extract_docx,
        ".xlsx": _extract_xlsx,
        ".xlsm": _extract_xlsx,
        ".pptx": _extract_pptx,
        ".pptm": _extract_pptx,
        ".pdf": _extract_pdf,
    }.get(path.suffix.lower(), _extract_text)
    try:
        result.embedded_title, result.excerpt, result.extraction_method = extractor(
            path, max_chars, max_pages
        )
        if not result.excerpt and not result.embedded_title:
            result.extraction_error = "未提取到可用于命名的文字"
    except Exception as exc:  # 格式/依赖异常降级到当前文件，批次可继续
        result.extraction_error = f"{type(exc).__name__}: {exc}"
    return result
