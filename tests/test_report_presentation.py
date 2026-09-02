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

    def test_problem_workbook_keeps_formal_and_hidden_ai_views(self):
        path = Path.cwd() / "outputs" / "test_report_presentation.xlsx"
        path.parent.mkdir(parents=True, exist_ok=True)
        try:
            run_audit.write_issue_workbook_xlsx(path, self.make_result())
            workbook = load_workbook(path, read_only=False, data_only=True)
            self.assertEqual(list(workbook["问题清单"].iter_rows(min_row=1, max_row=1, values_only=True))[0], ("等级", "规则", "Excel 行号", "KKS", "问题定义", "规则依据", "问题", "整改建议"))
            # 首条 issue 行（KKS-04b）：规则依据列按默认火电体系给出条款出处
            first_row = list(workbook["问题清单"].iter_rows(min_row=2, max_row=2, values_only=True))[0]
            self.assertEqual(first_row[5], run_audit.rule_source("KKS-04b", run_audit.STANDARD_DEFAULT))
            self.assertEqual(workbook["AI复核（技术）"].sheet_state, "hidden")
            self.assertEqual(workbook["概览"]["A1"].value, "KKS 编码质量审核报告")
        finally:
            path.unlink(missing_ok=True)


if __name__ == "__main__":
    unittest.main()
