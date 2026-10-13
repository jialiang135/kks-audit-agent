# -*- coding: utf-8 -*-
"""明细表不允许出现空格子（用户明确要求）。

背景：审核报告与问题清单是给业务人员看的，任何一格留空都会被读成"这里没查出东西"。
最容易空的两类是：
1. 「AI建议编码」——模型认为无需改码的行，或该行属"改码解决不了"的问题（名称卫生、
   编号断号、机组文本一致性…）；
2. 源表本身没填的列（如「专业」列整列为空）。
两者都必须写明文字（"无需改码 + 原因" / "—（源表已空）"），不能留白。
"""
from __future__ import annotations

import pathlib
import re
import sys
import unittest

PROJECT = pathlib.Path(__file__).resolve().parents[1]
if str(PROJECT) not in sys.path:
    sys.path.insert(0, str(PROJECT))

import report_business as RB  # noqa: E402

TMP = PROJECT / "_tmp_extract"


def _result() -> dict:
    """一份覆盖"改码型 / 非改码型 / 源表为空"三类行的审核结果。"""
    issues = [
        {"status": "open", "priority": "P0", "rule_id": "KKS-04b", "excel_row": 5,
         "kks_code": "00LBA10GP0O1", "name": "轴流风机O冒充0", "parent_code": "00LBA10",
         "message": "设备编号AN片段 '0O1' 含错误字符。", "suggestion": "建议候选：00LBA10GP001", "evidence": "e",
         "final_decision": "confirmed_issue", "final_summary": "确认为问题", "profession": ""},
        {"status": "open", "priority": "P2", "rule_id": "KKS-30", "excel_row": 19,
         "kks_code": "00XAB10GQ001", "name": "轴流风机（壁式）", "parent_code": "00XAB10",
         "message": "名称含全角符号。", "suggestion": "半角化", "evidence": "e",
         "final_decision": "confirmed_issue", "final_summary": "确认为问题", "profession": ""},
        {"status": "open", "priority": "P2", "rule_id": "KKS-F", "excel_row": 70,
         "kks_code": "00XAC10FM001", "name": "安全阀", "parent_code": "00XAC10",
         "message": "同父级编号断号：047。", "suggestion": "确认是否漏编", "evidence": "e",
         "final_decision": "confirmed_issue", "final_summary": "确认为问题", "profession": ""},
        {"status": "open", "priority": "P1", "rule_id": "KKS-28", "excel_row": 1040,
         "kks_code": "01XAM20GQ001", "name": "轴流风机1", "parent_code": "01XAM20", "unit": "2号机组",
         "message": "机组文本与编码前缀不一致。", "suggestion": "以编码前缀为准", "evidence": "e",
         "final_decision": "confirmed_issue", "final_summary": "确认为问题", "profession": ""},
        # 这一行用来验证"源表没填的列"（专业列为空）也有占位文字
        {"status": "open", "priority": "P2", "rule_id": "KKS-30", "excel_row": 300,
         "kks_code": "00XAD10GQ005", "name": "加压送风口（常闭）", "parent_code": "00XAD10",
         "message": "名称含全角符号。", "suggestion": "半角化", "evidence": "e",
         "final_decision": "confirmed_issue", "final_summary": "确认为问题", "profession": ""},
    ]
    return {
        "source_file": "blank.xlsx", "sheet": "暖通专业", "header_row": 1, "data_rows": 4,
        "columns": ["KKS编码", "父级KKS", "设备名称", "机组"],
        "sheet_info": [{"sheet": "暖通专业", "max_row": 5, "max_col": 4, "nonempty_rows": 4,
                        "first_nonempty_row": 1, "last_nonempty_row": 5,
                        "header_preview": ["KKS编码", "父级KKS", "设备名称", "机组"]}],
        "metrics": {"unique_codes": 4, "duplicate_code_groups": 0, "orphan_rows": 0,
                    "prefix_mismatch_rows": 0, "long_code_rows": 0},
        "priority_counts": {"P0": 1, "P1": 1, "P2": 2}, "issue_count": 4, "resolved_count": 0,
        "conclusion": "测试结论", "structural_notes": [], "issues": issues, "resolved_reviews": [],
        "import_governance": {"target": "DM8 LOCATIONS", "direct_import_ready": False,
                              "long_code_requires_extension_model": False,
                              "historical_kks_rewrite_forbidden": True, "runtime_dm8_validation": False},
        "scope_comparison": None,
        "_ai_context": {"records": [{"excel_row": i["excel_row"], "kks_code": i["kks_code"],
                                     "name": i["name"], "parent_code": i["parent_code"]} for i in issues]},
        "ai_review": {"status": "completed", "candidate_count": 4, "reviewed_count": 4},
    }


def _detail_tables(html: str) -> list[tuple[str, list[str], list[list[str]]]]:
    """抽出所有带「建议说明」列的明细表（含所属章节标题）。"""
    tables: list[tuple[str, list[str], list[list[str]]]] = []
    section = "（开头）"
    for token in re.split(r"(<h2>.*?</h2>|<h3[^>]*>.*?</h3>)", html, flags=re.DOTALL):
        head = re.match(r"<h[23][^>]*>(.*?)</h[23]>", token, flags=re.DOTALL)
        if head:
            section = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", "", head.group(1))).strip()
            continue
        for table in re.findall(r"<table[^>]*>.*?</table>", token, flags=re.DOTALL):
            headers = [re.sub(r"\s+", " ", re.sub(r"<[^>]+>", "", x)).strip()
                       for x in re.findall(r"<th[^>]*>(.*?)</th>", table, flags=re.DOTALL)]
            body = []
            for tr in re.findall(r"<tr>(.*?)</tr>", table, flags=re.DOTALL):
                cells = [re.sub(r"\s+", " ", re.sub(r"<[^>]+>", "", c)).strip()
                         for c in re.findall(r"<td[^>]*>(.*?)</td>", tr, flags=re.DOTALL)]
                if cells:
                    body.append(cells)
            tables.append((section, headers, body))
    return tables


class NoBlankCellsTest(unittest.TestCase):
    def setUp(self):
        self.result = _result()
        self.ov = RB.build_overview(self.result)
        self.html = RB.render_html(self.result, self.ov)

    def test_every_detail_cell_has_text(self):
        checked = 0
        for section, headers, body in _detail_tables(self.html):
            if "审查依据" not in headers:
                continue
            for row in body:
                for idx, value in enumerate(row):
                    col = headers[idx] if idx < len(headers) else f"col{idx}"
                    self.assertTrue(
                        value.strip(),
                        f"{section} 的「{col}」列出现空格子：{row}",
                    )
                    checked += 1
        self.assertGreater(checked, 0, "没有找到带建议说明的明细表，测试没起作用")

    def test_no_replacement_code_columns(self):
        for section, headers, body in _detail_tables(self.html):
            self.assertNotIn("AI建议编码", headers)
            self.assertNotIn("建议说明", headers)

    def test_blank_source_column_is_labeled(self):
        """源表为空的列要写明「源表未填」，不能编内容也不能留白。

        「专业」这类列在明细表里会出现，源表没填时必须显示占位说明。
        """
        self.assertEqual(RB._BLANK_DASH, "—（源表未填）")
        self.assertEqual(RB.cell_text(""), RB._BLANK_DASH)
        self.assertEqual(RB.cell_text(None), RB._BLANK_DASH)
        self.assertEqual(RB.cell_text("   "), RB._BLANK_DASH)
        self.assertEqual(RB.cell_text("无"), RB._BLANK_DASH)
        self.assertEqual(RB.cell_text("—"), RB._BLANK_DASH)
        self.assertEqual(RB.cell_text("暖通"), "暖通")
        # 数字要保持原类型（Excel 里要能排序/筛选）
        self.assertEqual(RB._blank_cell(12), 12)
        self.assertEqual(RB._blank_cell(""), RB._BLANK_DASH)

    def test_workbook_rows_have_no_blank_cells(self):
        import openpyxl

        TMP.mkdir(parents=True, exist_ok=True)
        out = TMP / "test_no_blank_cells.xlsx"
        RB.render_xlsx(out, self.result, self.ov)
        wb = openpyxl.load_workbook(out, read_only=False)
        try:
            for name in wb.sheetnames:
                if name in ("概览", "AI建议编码", "结构说明与范围", "AI复核建议"):
                    continue
                ws = wb[name]
                rows = [r for r in ws.iter_rows(values_only=True) if r and any(v not in (None, "") for v in r)]
                if len(rows) < 4:
                    continue
                headers = rows[2]
                for row in rows[3:]:
                    for idx, value in enumerate(row[: len(headers)]):
                        col = headers[idx] if idx < len(headers) else f"col{idx}"
                        self.assertTrue(
                            str(value or "").strip(),
                            f"{name} 的「{col}」列出现空格子：{row}",
                        )
        finally:
            wb.close()
            out.unlink(missing_ok=True)


if __name__ == "__main__":
    unittest.main()
