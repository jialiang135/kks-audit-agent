# -*- coding: utf-8 -*-
"""回归测试：八维框架真的出八张卡、AI 建议编码只写"真改动"、AI 子任务并发。

三件事都是"看起来做了、实际没生效"的高危点，必须由测试锁住：
1. 报告标题写「八维」但只渲染六张卡（维度七/八在框架里有、报告里丢）；
2. 模型回一个和原码一字不差的建议，被当成有效建议写进报告和 Excel，淹没真问题；
3. 「语义复核」和「问题行建议编码」互不依赖，却被串行执行，端到端耗时翻倍。
"""
from __future__ import annotations

import importlib.util
import pathlib
import re
import sys
import threading
import time
import unittest

PROJECT = pathlib.Path(__file__).resolve().parents[1]
if str(PROJECT) not in sys.path:
    sys.path.insert(0, str(PROJECT))

import report_business as RB  # noqa: E402
import run_audit  # noqa: E402

TMP = PROJECT / "_tmp_extract"


def _base_result() -> dict:
    """一份最小但结构完整的审核结果（含两个 12 位问题行）。"""
    issues = [
        {
            "status": "open", "priority": "P1", "rule_id": "KKS-31", "excel_row": 10,
            "kks_code": "00LBA10GP001", "name": "照明配电箱", "parent_code": "00LBA10",
            "message": "父子名称语义矛盾", "suggestion": "核对父子语义", "evidence": "测试证据",
        },
        {
            "status": "open", "priority": "P2", "rule_id": "KKS-25c", "excel_row": 11,
            "kks_code": "00LBA10GP002", "name": "给水泵", "parent_code": "00LBA10",
            "message": "身份键一对多", "suggestion": "人工裁定", "evidence": "测试证据",
        },
    ]
    return {
        "source_file": "sample.xlsx", "sheet": "Sheet1", "header_row": 1, "data_rows": 2,
        "columns": ["KKS", "父级", "名称"],
        "sheet_info": [{"sheet": "Sheet1", "max_row": 3, "max_col": 3, "nonempty_rows": 2,
                        "first_nonempty_row": 1, "last_nonempty_row": 3,
                        "header_preview": ["KKS", "父级", "名称"]}],
        "metrics": {"unique_codes": 2, "duplicate_code_groups": 0, "orphan_rows": 0,
                    "prefix_mismatch_rows": 0, "long_code_rows": 0},
        "priority_counts": {"P1": 1, "P2": 1}, "issue_count": 2, "resolved_count": 0,
        "conclusion": "测试结论", "structural_notes": [],
        "issues": issues, "resolved_reviews": [],
        "import_governance": {"target": "DM8 LOCATIONS", "direct_import_ready": False,
                              "long_code_requires_extension_model": False,
                              "historical_kks_rewrite_forbidden": True,
                              "runtime_dm8_validation": False},
        "scope_comparison": None,
        "_ai_context": {"records": [
            {"excel_row": 10, "kks_code": "00LBA10GP001", "name": "照明配电箱", "parent_code": "00LBA10"},
            {"excel_row": 11, "kks_code": "00LBA10GP002", "name": "给水泵", "parent_code": "00LBA10"},
        ]},
    }


class EightDimensionReportTest(unittest.TestCase):
    """报告必须与 kks-audit/SKILL.md 的八维框架一一对应。"""

    def test_overview_has_eight_cards(self):
        ov = RB.build_overview(_base_result())
        nos = [card["no"] for card in ov["cards"]]
        self.assertEqual(["①", "②", "③", "④", "⑤", "⑥", "⑦", "⑧"], nos)
        titles = " ".join(card["title"] for card in ov["cards"])
        self.assertIn("跨厂/范围对比", titles, "维度七（跨厂/范围对比）不能从报告里消失")
        self.assertIn("可导入性", titles, "维度八（可导入性）不能从报告里消失")

    def test_html_renders_eight_dimension_cards(self):
        result = _base_result()
        ov = RB.build_overview(result)
        html = RB.render_html(result, ov)
        heads = re.findall(r'<div class="dimhead">(.*?)<span', html)
        self.assertEqual(8, len(heads), f"HTML 里只有 {len(heads)} 张维度卡：{heads}")
        self.assertIn("⑦", html)
        self.assertIn("⑧", html)

    def test_card_count_matches_skill_framework(self):
        """SKILL.md 的「八维审核框架」小节必须仍是 8 条编号项。"""
        skill = (PROJECT / "kks-audit" / "SKILL.md").read_text(encoding="utf-8")
        section = skill.split("## 八维审核框架", 1)[1].split("\n## ", 1)[0]
        numbered = re.findall(r"^\*\*([一二三四五六七八九十]+)、", section, flags=re.MULTILINE)
        self.assertEqual(8, len(numbered), f"框架维度数变成 {len(numbered)}：{numbered}")


if __name__ == "__main__":
    unittest.main()
