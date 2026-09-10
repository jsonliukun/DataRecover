"""核心行为的回归测试。

测试全部使用 ``TemporaryDirectory`` 或内存构造的 OOXML 压缩包，不接触用户的
恢复目录。远程 provider 的 ``_post`` 会被 mock，因此测试不会访问真实 API，也
不会读取 ``.env`` 中的密钥。

覆盖重点不是模型命名质量，而是可重复验证的契约：Windows 名称/路径防护、方案
默认 approved、TXT/CSV 编码兼容、执行与撤销、公式转义，以及两种 API 响应结构。
出现过的真实回归（DeepSeek 字段别名、空响应、planned 统计键误改）都应在这里
保留断言，防止今后的重构再次引入。
"""

from __future__ import annotations

import csv
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest.mock import patch

from doc_renamer.providers import (
    ChatCompletionsProvider,
    ProviderSettings,
    ResponsesProvider,
    RulesProvider,
)
from doc_renamer.extractors import extract_document
from doc_renamer.types import DocumentContent
from doc_renamer.utils import safe_resolve_below, sanitize_stem
from doc_renamer.workflow import apply_plan, create_plan, undo_log


class UtilsTests(unittest.TestCase):
    """验证不依赖文件格式或网络的基础安全函数。"""

    def test_sanitize_windows_filename(self):
        """非法字符、设备名和公式前缀应被转换成可落盘主干。"""
        self.assertEqual(sanitize_stem('  项目:计划 / 2025?.  '), "项目_计划 _ 2025")
        self.assertEqual(sanitize_stem("CON"), "_CON")
        self.assertEqual(sanitize_stem("=2+2"), "_=2+2")

    def test_path_escape_is_rejected(self):
        """方案中的 ``..`` 不能把目标解析到指定根目录之外。"""
        with tempfile.TemporaryDirectory() as directory:
            with self.assertRaises(ValueError):
                safe_resolve_below(Path(directory), "../outside.txt")


class WorkflowTests(unittest.TestCase):
    """验证 plan → 人工方案 → apply → undo 的文件系统闭环。"""

    def test_plan_apply_and_undo(self):
        """主路径应保留 planned 统计、默认批准、改名和可撤销性。"""
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory) / "recovered"
            root.mkdir()
            original = root / "20250101_001.txt"
            original.write_text("甲项目验收报告。项目已经通过验收。", encoding="utf-8")
            plan = Path(directory) / "plan.csv"
            counts = create_plan(root, plan, RulesProvider(), progress=lambda _: None)
            self.assertEqual(counts["planned"], 1)

            with plan.open("r", encoding="utf-8-sig", newline="") as handle:
                rows = list(csv.DictReader(handle))
                fieldnames = list(rows[0].keys())
            self.assertEqual(rows[0]["status"], "approved")
            with plan.open("w", encoding="utf-8-sig", newline="") as handle:
                writer = csv.DictWriter(handle, fieldnames=fieldnames)
                writer.writeheader()
                writer.writerows(rows)

            log = Path(directory) / "log.csv"
            applied = apply_plan(root, plan, log)
            self.assertEqual(applied["renamed"], 1)
            renamed = root / rows[0]["proposed_name"]
            self.assertTrue(renamed.exists())
            self.assertFalse(original.exists())

            undone = undo_log(root, log)
            self.assertEqual(undone["restored"], 1)
            self.assertTrue(original.exists())

    def test_approved_empty_name_is_logged_and_does_not_abort_batch(self):
        """已批准但名称为空时应记失败日志，而不是改名或令整批崩溃。"""
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory) / "recovered"
            root.mkdir()
            original = root / "20250101_001.txt"
            original.write_text("测试内容", encoding="utf-8")
            plan = Path(directory) / "plan.csv"
            create_plan(root, plan, RulesProvider(), progress=lambda _: None)
            with plan.open("r", encoding="utf-8-sig", newline="") as handle:
                rows = list(csv.DictReader(handle))
                fieldnames = list(rows[0].keys())
            rows[0]["status"] = "approved"
            rows[0]["proposed_name"] = ""
            with plan.open("w", encoding="utf-8-sig", newline="") as handle:
                writer = csv.DictWriter(handle, fieldnames=fieldnames)
                writer.writeheader()
                writer.writerows(rows)

            log = Path(directory) / "log.csv"
            applied = apply_plan(root, plan, log)
            self.assertEqual(applied["failed"], 1)
            self.assertTrue(original.exists())
            with log.open("r", encoding="utf-8-sig", newline="") as handle:
                log_row = next(csv.DictReader(handle))
            self.assertIn("必须是单独的文件名", log_row["message"])

    def test_existing_plan_is_not_overwritten_by_default(self):
        """未给 ``overwrite`` 时，已有人工方案必须保持原内容。"""
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory) / "recovered"
            root.mkdir()
            (root / "one.txt").write_text("测试内容", encoding="utf-8")
            plan = Path(directory) / "plan.csv"
            plan.write_text("keep me", encoding="utf-8")
            with self.assertRaises(FileExistsError):
                create_plan(root, plan, RulesProvider(), progress=lambda _: None)
            self.assertEqual(plan.read_text(encoding="utf-8"), "keep me")

    def test_old_office_file_is_reported_not_parsed_as_text(self):
        """旧二进制 Office 文件应明确标记不支持，不能按纯文本误读。"""
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "recovered.doc"
            path.write_bytes(b"binary-data")
            result = extract_document(path)
            self.assertIn("暂不支持", result.extraction_error)
            self.assertEqual(result.excerpt, "")

    def test_txt_plan_is_tsv_and_excel_utf16_can_be_applied(self):
        """默认 TXT 应写 TSV，并能读取 Excel 常见的 UTF-16 保存结果。"""
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory) / "recovered"
            root.mkdir()
            original = root / "20250101_001.txt"
            original.write_text("甲项目验收报告。", encoding="utf-8")
            plan = Path(directory) / "plan.txt"
            create_plan(root, plan, RulesProvider(), progress=lambda _: None)

            generated = plan.read_text(encoding="utf-8-sig")
            self.assertIn("\t", generated.splitlines()[0])
            rows = list(csv.DictReader(generated.splitlines(), delimiter="\t"))
            fieldnames = list(rows[0].keys())
            rows[0]["status"] = "approved"

            # Excel 的“Unicode 文本”通常保存成带 BOM 的 UTF-16 制表符文件。
            with plan.open("w", encoding="utf-16", newline="") as handle:
                writer = csv.DictWriter(handle, fieldnames=fieldnames, delimiter="\t")
                writer.writeheader()
                writer.writerows(rows)

            log = Path(directory) / "log.txt"
            applied = apply_plan(root, plan, log)
            self.assertEqual(applied["renamed"], 1)
            self.assertIn("\t", log.read_text(encoding="utf-8-sig").splitlines()[0])
            self.assertEqual(undo_log(root, log)["restored"], 1)
            self.assertTrue(original.exists())

    def test_gb18030_csv_can_be_applied(self):
        """中文 Windows/WPS 保存的无 BOM GB18030 方案仍应可执行。"""
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory) / "recovered"
            root.mkdir()
            original = root / "20250101_001.txt"
            original.write_text("中文项目总结。", encoding="utf-8")
            plan = Path(directory) / "plan.csv"
            create_plan(root, plan, RulesProvider(), progress=lambda _: None)
            with plan.open("r", encoding="utf-8-sig", newline="") as handle:
                rows = list(csv.DictReader(handle))
                fieldnames = list(rows[0].keys())
            rows[0]["status"] = "approved"
            with plan.open("w", encoding="gb18030", newline="") as handle:
                writer = csv.DictWriter(handle, fieldnames=fieldnames)
                writer.writeheader()
                writer.writerows(rows)

            applied = apply_plan(root, plan, Path(directory) / "log.csv")
            self.assertEqual(applied["renamed"], 1)
            self.assertFalse(original.exists())

    def test_invalid_plan_header_is_rejected_before_renaming(self):
        """缺少安全关键列时，应在任何文件改名前整体拒绝方案。"""
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory) / "recovered"
            root.mkdir()
            original = root / "keep.txt"
            original.write_text("不要改名", encoding="utf-8")
            plan = Path(directory) / "broken.txt"
            plan.write_text(
                "status\toriginal_relpath\tproposed_name\n"
                "approved\tkeep.txt\trenamed.txt\n",
                encoding="utf-8-sig",
            )
            with self.assertRaisesRegex(ValueError, "sha256"):
                apply_plan(root, plan, Path(directory) / "log.txt")
            self.assertTrue(original.exists())
            self.assertFalse((root / "renamed.txt").exists())

    def test_formula_like_cells_are_escaped_and_paths_are_restored(self):
        """表格公式防护必须可逆，不能因此找不到真实源文件。"""
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory) / "recovered"
            root.mkdir()
            original = root / "=2+2.txt"
            original.write_text("安全测试标题。", encoding="utf-8")
            plan = Path(directory) / "plan.txt"
            create_plan(root, plan, RulesProvider(), progress=lambda _: None)

            with plan.open("r", encoding="utf-8-sig", newline="") as handle:
                rows = list(csv.DictReader(handle, delimiter="\t"))
                fieldnames = list(rows[0].keys())
            self.assertEqual(rows[0]["original_relpath"], "'=2+2.txt")
            self.assertEqual(rows[0]["original_name"], "'=2+2.txt")
            rows[0]["status"] = "approved"
            with plan.open("w", encoding="utf-8-sig", newline="") as handle:
                writer = csv.DictWriter(handle, fieldnames=fieldnames, delimiter="\t")
                writer.writeheader()
                writer.writerows(rows)

            log = Path(directory) / "log.txt"
            applied = apply_plan(root, plan, log)
            self.assertEqual(applied["renamed"], 1)
            self.assertFalse(original.exists())
            self.assertEqual(undo_log(root, log)["restored"], 1)
            self.assertTrue(original.exists())

    def test_legacy_comma_delimited_txt_plan_is_supported(self):
        """扩展名虽为 TXT 的旧逗号方案应按内容探测并保持兼容。"""
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory) / "recovered"
            root.mkdir()
            original = root / "legacy.txt"
            original.write_text("旧方案兼容测试。", encoding="utf-8")
            csv_plan = Path(directory) / "plan.csv"
            create_plan(root, csv_plan, RulesProvider(), progress=lambda _: None)
            legacy_plan = Path(directory) / "plan.txt"
            legacy_plan.write_bytes(csv_plan.read_bytes())
            with legacy_plan.open("r", encoding="utf-8-sig", newline="") as handle:
                rows = list(csv.DictReader(handle))
                fieldnames = list(rows[0].keys())
            rows[0]["status"] = "approved"
            with legacy_plan.open("w", encoding="utf-8-sig", newline="") as handle:
                writer = csv.DictWriter(handle, fieldnames=fieldnames)
                writer.writeheader()
                writer.writerows(rows)

            applied = apply_plan(root, legacy_plan, Path(directory) / "log.txt")
            self.assertEqual(applied["renamed"], 1)
            self.assertFalse(original.exists())


class ExtractorTests(unittest.TestCase):
    """用最小 OOXML 样本验证元数据与正文提取，不依赖 Office。"""

    CORE_XML = """<?xml version="1.0" encoding="UTF-8"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
 xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>内部项目标题</dc:title></cp:coreProperties>"""

    def test_docx_xml_extraction(self):
        """DOCX core title 与 document.xml 正文应同时进入抽取结果。"""
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "recovered.docx"
            with zipfile.ZipFile(path, "w") as archive:
                archive.writestr("docProps/core.xml", self.CORE_XML)
                archive.writestr(
                    "word/document.xml",
                    '<w:document xmlns:w="urn:w"><w:body><w:p><w:r><w:t>验收报告正文</w:t></w:r></w:p></w:body></w:document>',
                )
            result = extract_document(path)
            self.assertEqual(result.embedded_title, "内部项目标题")
            self.assertIn("验收报告正文", result.excerpt)
            self.assertFalse(result.extraction_error)


class ProviderTests(unittest.TestCase):
    """以 mock 响应验证 provider 请求契约和容错解析，不发网络请求。"""

    def setUp(self):
        """构造各测试共享的假配置和不含真实资料的 DocumentContent。"""
        self.settings = ProviderSettings(
            provider="responses",
            api_base="https://example.invalid/v1",
            api_key="test-key",
            model="test-model",
            retries=1,
        )
        self.document = DocumentContent(
            path=Path("20250101_001.docx"),
            sha256="0" * 64,
            size=100,
            modified_at="2025-01-01T00:00:00+00:00",
            excerpt="甲项目验收报告正文",
        )

    def test_responses_output_array_is_parsed(self):
        """Responses 的嵌套 output_text 与严格 JSON Schema 配置应被识别。"""
        provider = ResponsesProvider(self.settings)
        response = {
            "output": [
                {
                    "content": [
                        {
                            "type": "output_text",
                            "text": '{"title":"甲项目验收报告","category":"验收报告","confidence":93,"reason":"正文明确"}',
                        }
                    ]
                }
            ]
        }
        with patch.object(provider, "_post", return_value=response) as post:
            suggestion = provider.suggest(self.document)
        self.assertEqual(suggestion.title, "甲项目验收报告")
        payload = post.call_args.args[1]
        self.assertFalse(payload["store"])
        self.assertEqual(payload["text"]["format"]["type"], "json_schema")

    def test_chat_completion_is_parsed(self):
        """Chat fenced JSON 应可解析，请求必须启用 json_object 模式。"""
        settings = ProviderSettings(
            provider="chat",
            api_base="https://example.invalid/v1",
            api_key="test-key",
            model="test-model",
            retries=1,
        )
        provider = ChatCompletionsProvider(settings)
        response = {
            "choices": [
                {
                    "message": {
                        "content": '```json\n{"title":"采购合同","category":"合同","confidence":88,"reason":"包含合同条款"}\n```'
                    }
                }
            ]
        }
        with patch.object(provider, "_post", return_value=response) as post:
            suggestion = provider.suggest(self.document)
        self.assertEqual(suggestion.category, "合同")
        self.assertEqual(suggestion.confidence, 88)
        payload = post.call_args.args[1]
        self.assertEqual(payload["response_format"], {"type": "json_object"})

    def test_chat_completion_accepts_deepseek_filename_alias(self):
        """DeepSeek 偶发的 suggested_filename 别名不能再触发空标题错误。"""
        settings = ProviderSettings(
            provider="chat",
            api_base="https://example.invalid",
            api_key="test-key",
            model="deepseek-v4-flash",
            retries=1,
        )
        provider = ChatCompletionsProvider(settings)
        response = {
            "choices": [
                {
                    "finish_reason": "stop",
                    "message": {
                        "content": '{"suggested_filename":"测试项目周报","confidence":85}'
                    },
                }
            ]
        }
        with patch.object(provider, "_post", return_value=response):
            suggestion = provider.suggest(self.document)
        self.assertEqual(suggestion.title, "测试项目周报")
        self.assertEqual(suggestion.category, "待分类")

    def test_chat_completion_retries_empty_content_once(self):
        """Chat 首次 content 为空时应仅做一次内容级重试并采用第二次结果。"""
        settings = ProviderSettings(
            provider="chat",
            api_base="https://example.invalid",
            api_key="test-key",
            model="deepseek-v4-flash",
            retries=1,
        )
        provider = ChatCompletionsProvider(settings)
        empty = {"choices": [{"finish_reason": "stop", "message": {"content": ""}}]}
        valid = {
            "choices": [
                {
                    "finish_reason": "stop",
                    "message": {
                        "content": '{"title":"甲项目验收报告","category":"验收报告","confidence":91,"reason":"正文明确"}'
                    },
                }
            ]
        }
        with patch.object(provider, "_post", side_effect=[empty, valid]) as post:
            suggestion = provider.suggest(self.document)
        self.assertEqual(suggestion.title, "甲项目验收报告")
        self.assertEqual(post.call_count, 2)


if __name__ == "__main__":
    unittest.main()
