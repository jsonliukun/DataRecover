"""跨模块数据结构：规定抽取结果和命名建议在流水线中的字段语义。

``DocumentContent`` 是抽取阶段对某个文件的一次观察结果；``NameSuggestion``
是规则或模型基于该观察给出的候选名称。将它们放在中立模块中，使
``extractors``、``providers`` 和 ``workflow`` 依赖同一份契约，而无需彼此导入。

这里的“契约”主要由类型注解和 docstring 表达，并非运行时验证器。dataclass
不会自动检查传入值是否真是声明类型、置信度是否越界、路径是否存在，也不会自动
净化模型输出；生产者必须遵守约定，消费者仍需在安全边界上校验外部数据。

``from __future__ import annotations`` 让注解延迟保存而不是在类定义时立即求值，
便于前向引用并减少注解造成的运行时耦合。它不会把新版语法“移植”到旧 Python；
本项目仍以 ``pyproject.toml`` 声明的 Python 3.10+ 为准。
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path


@dataclass(slots=True)
class DocumentContent:
    """单个文件的抽取结果，是提取器交给命名器和工作流的输入对象。

    它是“观察快照”而非文件副本：创建后磁盘文件仍可能变化，对象本身也可修改。
    工作流会把哈希、内嵌标题、抽取方式、错误、大小和修改时间等写入方案用于审核；
    ``excerpt`` 当前只用于生成建议，不写入方案，以免方案膨胀并减少正文暴露。

    ``@dataclass`` 自动生成初始化、表示和相等比较方法；``slots=True`` 固定允许的
    属性并省去通常的实例 ``__dict__``，但它不等于 ``frozen=True``，不能提供
    不可变性，也不会执行任何类型或值域检查。

    字段契约：
        path
            本次抽取对应的 ``Path``。生产者可传绝对或相对路径；本数据类不解析
            ``..``、不确认路径存在，也不保证其位于某个根目录下。
        sha256
            正常情况下为读取文件内容计算出的 64 个小写十六进制字符 SHA-256，用于
            方案阶段识别重复文件，并在 apply/undo 前校验内容未变。初始 stat 或
            哈希失败时约定为空字符串；本类自身不校验长度或字符集。
        size
            初始 ``stat`` 得到的字节数。读取元数据/哈希失败时使用 ``0`` 哨兵，
            但合法空文件大小同样为 0，必须结合 ``sha256`` 和错误字段判断。
        modified_at
            初始 ``stat`` 时间转换成的 UTC ISO 8601 字符串，正常值带时区偏移；
            初始读取失败时为空。它用于审计展示，不参与内容一致性判断。
        embedded_title
            文档容器里的标题元数据，例如 OOXML ``docProps/core.xml`` 的 title 或
            PDF ``/Title``。没有该属性、格式不提供或抽取未完成时为空；它是未经
            人工确认的外部内容，不能天然视为可信事实。
        excerpt
            从有限来源取得并经过空白压缩、字符截断的有损正文节选。它不保证完整
            或保持视觉顺序，且可能包含机密内容或提示注入文本；命名器必须把它当
            作不可信数据，而不是程序指令。
        extraction_method
            成功返回三元组的抽取器标识，当前可能为 ``docx-xml``、``xlsx-xml``、
            ``pptx-xml``、``pypdf`` 或 ``plain-text``。抽取器抛异常时通常为空；
            即使标识非空，仍只代表采用过该策略，不代表全文完整或格式完全有效。
        extraction_error
            面向人工排查的诊断文本。空字符串表示抽取阶段没有检测到错误，不表示
            内容完整或无风险；非空可表示文件不可读、格式/OCR 不支持、解析失败或
            没有命名所需文字。后续工作流还可能追加模型调用错误，因此消费者应
            同时查看 ``embedded_title`` 和 ``excerpt``，不能只把它当布尔成功标志。
    """

    path: Path
    sha256: str
    size: int
    modified_at: str
    embedded_title: str = ""
    excerpt: str = ""
    extraction_method: str = ""
    extraction_error: str = ""


@dataclass(slots=True)
class NameSuggestion:
    """规则或模型生成的候选命名信息；它本身不是改名命令。

    该对象刻意不包含源路径、最终目标路径或方案 ``status``。工作流会先把
    ``title`` 净化、追加源文件扩展名并解决重名，再把建议写入方案；是否执行由
    方案审核和 apply 阶段决定。即使正常建议目前默认标为 ``approved``，也不能
    把 ``NameSuggestion`` 的存在理解为已经完成改名或绕过人工审核。

    和 ``DocumentContent`` 一样，dataclass 不在运行时验证字段。内置 API 解析器
    会对部分字段做容错和钳制，但自定义 ``NamingProvider`` 仍必须遵守以下约定：

        title
            非空的建议文件名主干，约定不含目录和扩展名。值来自规则/模型，属于
            不可信输入；工作流必须通过 ``sanitize_stem`` 去除 Windows 非法字符、
            设备保留名和公式样式前缀后，才能与原扩展名组合为最终候选名。
        category
            供人工审核的文档类别标签，例如“周报”“合同”；当前仅写入方案，
            不控制目录、权限、状态或后续路由。
        confidence
            按约定为 0 到 100 的整数，用来表达命名器自己的把握程度，而不是经过
            校准的概率，也不自动决定是否执行。API 返回解析器会转成整数并钳制
            范围；本类和自定义提供方不会自动保证该值合法。
        reason
            面向审核者的简短理由，用于说明标题/分类依据。它同样可能由模型生成，
            只能作为解释线索而非事实证明；写入表格时由工作流执行公式注入转义。
    """

    title: str
    category: str
    confidence: int
    reason: str
