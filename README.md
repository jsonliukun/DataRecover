# 恢复文档智能命名工具

这是一个安全优先的 Python 命令行工具。它读取恢复文档的内部标题和正文节选，生成候选文件名方案；`plan` 本身不会修改任何文档。成功生成建议的行默认是 `status=approved`，因此运行 `apply` 前必须审核方案，并把不想执行的行改成 `pending`。默认使用便于 Excel/WPS 编辑的制表符 `.txt`，也兼容 `.csv`。

如果你不熟悉 Python 或命令行，请直接阅读 **[新手使用指南](新手使用指南.md)**。指南会逐项说明文件放在哪里、需要修改哪个文件，以及从测试到撤销的完整操作。

## 当前支持

- 文档：`.docx/.docm`、`.xlsx/.xlsm`、`.pptx/.pptm`、`.pdf`
- 文本：`.txt`、`.md`、`.csv`、`.tsv`、`.json`、`.xml`、`.html`、`.htm`、`.rtf`、`.log`
- 命名后端：
  - `rules`：不联网，用内部标题或正文首句试跑流程
  - `responses`：OpenAI Responses API
  - `chat`：OpenAI-compatible Chat Completions，可用于很多云端服务以及之后的本地服务
- SHA-256 重复文件识别、改名前哈希校验、路径越界保护、目标冲突保护和撤销日志

旧格式 `.doc/.xls/.ppt`、`.jpg/.jpeg/.png/.bmp/.tif/.tiff` 图片与扫描 PDF 暂不做正文识别。已列出的格式会在方案中标记需要转换或 OCR，不会被错误地自动命名；其他未列出的格式会被忽略。

## 1. 安装

建议使用 Python 3.10 或更高版本。第一次在一台新电脑上测试时，在 PowerShell 中执行：

```powershell
git clone https://github.com/jsonliukun/DataRecover.git
cd DataRecover
py -m venv .venv
.\.venv\Scripts\python.exe -m pip install --upgrade pip
.\.venv\Scripts\python.exe -m pip install .
.\.venv\Scripts\python.exe -m pip check
.\.venv\Scripts\python.exe -c "import httpx, dotenv, pypdf, doc_renamer; print('依赖安装正常')"
if (-not (Test-Path ".env")) { Copy-Item ".env.example" ".env" }
```

也可以在项目根目录运行一键安装与测试脚本：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\setup.ps1
```

虚拟环境只负责隔离 Python 包，不会自动包含第三方依赖，也不应从另一台电脑复制 `.venv`。本项目的直接依赖统一声明在 `pyproject.toml` 中；上面的 `pip install .` 会自动安装 `httpx`、`pypdf` 和 `python-dotenv`。需要修改源码并让改动立即生效的开发者，可以把安装命令改为 `pip install -e .`。

不要把真实恢复文件直接当作唯一副本。先复制一份工作目录，并只对工作副本运行本工具。

## 2. 不联网试跑

先用 20 个文件测试提取效果：

```powershell
doc-renamer plan "D:\Recovered_Working_Copy" --provider rules --limit 20 --output rename_plan_test.txt
```

打开 `rename_plan_test.txt` 查看结果。此命令不会改名。

省略 `--output` 时，默认在当前 PowerShell 目录生成 `rename_plan.txt`；建议像示例一样显式指定路径。

## 3. 配置 API

如果还没有 `.env`，复制 `.env.example` 后再填写；已有 `.env` 时直接编辑，切勿用模板覆盖。密钥只放在 `.env` 的 `DOC_RENAMER_API_KEY`，不要写到命令行、源代码或方案文件。

OpenAI Responses 示例：

```dotenv
DOC_RENAMER_PROVIDER=responses
DOC_RENAMER_API_BASE=https://api.openai.com/v1
DOC_RENAMER_API_KEY=你的密钥
DOC_RENAMER_MODEL=你账户可用的模型ID
```

OpenAI-compatible Chat Completions 示例：

```dotenv
DOC_RENAMER_PROVIDER=chat
DOC_RENAMER_API_BASE=https://服务商地址/v1
DOC_RENAMER_API_KEY=你的密钥
DOC_RENAMER_MODEL=模型ID
```

DeepSeek 示例：

```dotenv
DOC_RENAMER_PROVIDER=chat
DOC_RENAMER_API_BASE=https://api.deepseek.com
DOC_RENAMER_API_KEY=你的密钥
DOC_RENAMER_MODEL=deepseek-v4-flash
```

这里使用的是本项目已验证的 DeepSeek Chat Completions 配置。基础地址不要追加 `/chat/completions`。旧的 `deepseek-chat/deepseek-reasoner` 已于 2026-07-24 停用；需要更强模型时可按账户权限使用 `deepseek-v4-pro`。模型名称变化时以 [DeepSeek 官方更新日志](https://api-docs.deepseek.com/updates/) 为准。

先发送一段虚拟内容测试连通性，不会读取真实文档。`chat` 模式（包括 DeepSeek）正常情况下请求一次；如果模型返回空内容或错误 JSON，程序会自动再请求一次：

```powershell
doc-renamer probe-api
```

再小批量生成方案：

```powershell
doc-renamer plan "D:\Recovered_Working_Copy" --limit 20 --output rename_plan_test.txt
```

工具默认从正文中提取最多 6000 个字符；PDF 最多读取前 5 页，PPTX 最多读取前 5 张幻灯片。API 载荷还会包含当前文件名、扩展名、修改时间和内部标题等元数据。使用云端 API 时，这些内容会离开本机；涉及敏感文件时，应取得授权、确认服务商的数据策略，或继续使用 `rules`，等本地模型部署后再切换。

OpenAI Responses 请求显式设置了 `store=false`，但组织政策、日志与其他数据处理条件仍应以你的服务协议和账户配置为准。

## 4. 人工审核并执行

在方案文件中审核 `proposed_name`。正常建议默认是 `approved`；把不想执行或尚未确认的行改成 `pending`。`duplicate` 和 `error` 行不会自动批准，建议人工处理。

`.txt` 会按 UTF-8 BOM + 制表符生成。读取时会自动识别 UTF-8、带 BOM 的 UTF-16、GB18030、Big5，以及制表符/逗号/分号，因此 Excel/WPS 改变编码后通常也可直接执行；旧版“逗号分隔但扩展名为 `.txt`”的方案也兼容。保存提示出现时选择保留当前文本格式，不要另存为 `.xlsx`。

为避免 Excel/WPS 把恢复出来的文件名或标题当成公式，首个非空白字符为 `= + - @` 的单元格会显示一个前导单引号；原值本来以单引号开头时也会临时加倍，以保证可逆。只编辑 `status` 和 `proposed_name`，不要删除其他列的保护符；程序在 `apply/undo` 时会自动还原真实值。

然后执行：

```powershell
doc-renamer apply "D:\Recovered_Working_Copy" rename_plan_test.txt --log rename_log.txt
```

省略 `--log` 时，默认在当前 PowerShell 目录生成 `rename_log.txt`。

如果原文件在生成方案后发生变化、目标名称已存在或路径异常，该行会失败而不会强行覆盖。

## 5. 撤销

```powershell
doc-renamer undo "D:\Recovered_Working_Copy" rename_log.txt
```

撤销前同样会校验哈希。如果改名后的文件内容已经变化，工具不会恢复该文件，并在结果中计入 `failed`，避免误操作。

## 测试

```powershell
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
```

当前版本应显示 `Ran 16 tests` 和 `OK`。

## 下一步建议

第一轮用 100～200 个有代表性的文件统计准确率。后续可以继续加入 OCR、并发与限流、相似版本聚类、更多 Office 元数据，以及本地 Ollama/vLLM 配置；主工作流不需要重写。
