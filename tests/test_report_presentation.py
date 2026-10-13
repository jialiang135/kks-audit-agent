import unittest
from pathlib import Path

from openpyxl import load_workbook

import run_audit


class ReportPresentationTests(unittest.TestCase):
    def make_result(self):
        issue = lambda priority, rule_id, row, code, message: {
            "status": "open",
            "priority": priority,
            "rule_id": rule_id,
            "excel_row": row,
            "kks_code": code,
            "message": message,
            "suggestion": "请核对后整改。",
            "evidence": "测试证据",
            "final_decision": "confirmed_issue",
        }
        return {
            "source_file": "sample.xlsx",
            "sheet": "Sheet1",
            "header_row": 1,
            "data_rows": 2,
            "columns": ["KKS", "父级", "名称"],
            "sheet_info": [{"sheet": "Sheet1", "max_row": 3, "max_col": 3, "nonempty_rows": 2, "first_nonempty_row": 1, "last_nonempty_row": 3, "header_preview": ["KKS", "父级", "名称"]}],
            "metrics": {"unique_codes": 2, "duplicate_code_groups": 0, "orphan_rows": 0, "prefix_mismatch_rows": 0, "long_code_rows": 1},
            "priority_counts": {"P0": 1, "P1": 1, "P2": 1},
            "issue_count": 3,
            "resolved_count": 0,
            "conclusion": "测试结论",
            "structural_notes": [],
            "issues": [
                issue("P0", "KKS-04b", 2, "01AAA10QM0O1", "字符错误"),
                issue("P1", "KKS-17/22", 3, "01AAA10QM001A", "扩展编码"),
                issue("P2", "KKS-18", 4, "01AAA10QM002", "历史迁移"),
            ],
            "resolved_reviews": [],
            "import_governance": {"direct_import_ready": False},
            "scope_comparison": None,
            "_ai_context": {"records": [{"kks_code": "01AAA10QM001"}, {"kks_code": "01AAA10QM001A"}]},
        }

    def test_business_metrics_use_real_priority_counts(self):
        result = self.make_result()
        self.assertEqual(
            run_audit.quality_metrics(result),
            {
                "effective_kks": 2,
                "device_level_codes": 1,
                "duplicate_codes": 0,
                "parent_errors": 0,
                "p0": 1,
                "p1": 1,
                "p2": 1,
                "rule_p1": 1,
                "rule_p2": 1,
                "focused_p1": 1,
                "focused_p2": 1,
            },
        )

    def test_html_and_excel_share_the_api_summary_numbers(self):
        from report_business import build_overview
        result = self.make_result()
        api = run_audit.quality_metrics(result)
        overview = build_overview(result)
        kpis = {k["key"]: int(k["value"].replace(",", "")) for k in overview["kpis"] if "key" in k}
        for key in ("effective_kks", "device_level_codes", "duplicate_codes", "parent_errors", "p0", "p1", "p2"):
            self.assertEqual(api[key], kpis[key])
        import tempfile
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "metrics.xlsx"
            run_audit.write_issue_workbook_xlsx(path, result)
            workbook = load_workbook(path, read_only=True, data_only=True)
            actual = {row[1]: row[2] for row in workbook["概览"].iter_rows(values_only=True) if row[0] == "统一指标"}
            for key, label in (("effective_kks", "有效KKS数量"), ("device_level_codes", "设备级编码"),
                               ("duplicate_codes", "重复编码"), ("parent_errors", "父级问题"),
                               ("p0", "高风险问题P0"), ("p1", "待整改或待核问题P1"), ("p2", "提示问题P2")):
                self.assertEqual(api[key], actual[label])
            workbook.close()

    def test_problem_workbook_uses_business_group_sheets(self):
        path = Path.cwd() / "outputs" / "test_report_presentation.xlsx"
        path.parent.mkdir(parents=True, exist_ok=True)
        try:
            run_audit.write_issue_workbook_xlsx(path, self.make_result())
            workbook = load_workbook(path, read_only=False, data_only=True)
            # 广元风格 sheet 结构：概览 + 按业务分组的明细 sheet + 结构说明
            self.assertIn("概览", workbook.sheetnames)
            self.assertIn("P0_阻断", workbook.sheetnames)
            self.assertIn("P2_其他提示", workbook.sheetnames)
            self.assertIn("结构说明与范围", workbook.sheetnames)
            self.assertTrue(workbook["概览"]["A1"].value.startswith("sample 审核总览"))
            # P0_阻断 sheet：第 3 行为表头、数据从第 4 行开始（与广元样式一致）
            # 「AI建议编码 / 建议说明」两列固定插在「编码」之后：用户要求明细表任何一格
            # 都要有文字解释，不能留空格子（没有建议码时该列显示"无需改码"并说明原因）。
            headers = [cell.value for cell in workbook["P0_阻断"][3]]
            self.assertEqual(headers, ["工作表", "行", "编码", "规则", "问题", "整改建议", "审查结论", "审查依据"])
            first = [cell.value for cell in workbook["P0_阻断"][4]]
            self.assertEqual(first[1:3], [2, "01AAA10QM0O1"])
            self.assertEqual(first[3:6], ["KKS-04b", "字符错误", "请核对后整改。"])
            for value in first:
                self.assertTrue(str(value or "").strip(), "明细行不允许出现空格子")
        finally:
            path.unlink(missing_ok=True)


if __name__ == "__main__":
    unittest.main()
