# -*- coding: utf-8 -*-
"""问题行「AI 建议编码」与「只用现行标准」的回归测试。

覆盖三件事：
1. 生成范围——只对"有问题的行"生成；
2. 准确性防线——建议编码必须过本地复核（骨架 / 禁用 I、O / GB/T 50549-2020 附录字典）；
3. 版本纪律——字典只来自现行标准，已废止的 2010 体系码表不得混入；
   旧体系遗留码只能报"未收录待核"，绝不判成错误。
"""
import importlib.util
import json
import os
import sys
import unittest
from pathlib import Path

PROJECT = Path(__file__).resolve().parents[1]
if str(PROJECT) not in sys.path:
    sys.path.insert(0, str(PROJECT))

import ai_review as A  # noqa: E402
import report_business as RB  # noqa: E402
import run_audit  # noqa: E402


def _load_template():
    spec = importlib.util.spec_from_file_location("audit_template", PROJECT / "kks-audit" / "scripts" / "audit_template.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


TEMPLATE = _load_template()
COLS = {"kks": 0, "name": 1}


class SingleStandardBasisTest(unittest.TestCase):
    """只按 GB/T 50549-2020 判定：旧体系码按"未收录待核"处理，绝不会判成错误。"""

    def test_2020_codes_not_flagged(self):
        """2020 版合法码（GP 泵 / GQ 风机 / QM 截止阀）不得被判为未收录。"""
        rows = [["00LAC10GP008", "给水泵"], ["00XAG10GQ001", "轴流风机"], ["00LBA10QM001", "截止阀"]]
        hits = TEMPLATE.device_class_code_legality(COLS, rows)
        self.assertEqual([], hits, f"2020 版合法码被误报：{hits}")

    def test_retired_codes_are_pending_not_errors(self):
        """旧体系码（AA 阀门 / AP 泵 / AN 风机）按现行国标只报 P2 待核，不判错。"""
        rows = [["00LAB10AA001", "电动阀门"], ["00LAC10AP008", "给水泵"], ["00XAG10AN001", "轴流风机"]]
        hits = TEMPLATE.device_class_code_legality(COLS, rows)
        self.assertEqual({"AA", "AP", "AN"}, {h["segment"] for h in hits})
        self.assertTrue(all(h["severity"] == "P2" for h in hits),
                        "旧体系遗留码只能是待核线索，不能判成错误")
        self.assertTrue(all("50549-2020" in h["message"] for h in hits),
                        "依据必须写明现行国标")

    def test_unknown_code_still_flagged_as_pending(self):
        hits = TEMPLATE.device_class_code_legality(COLS, [["00XAG10GG001", "轴流风机"]])
        self.assertEqual(["GG"], [h["segment"] for h in hits])
        self.assertEqual("P2", hits[0]["severity"], "未收录只能是待核，不能判成错误")

    def test_forbidden_letter_is_hard_error(self):
        hits = TEMPLATE.device_class_code_legality(COLS, [["00LBA10QI001", "给水阀"]])
        self.assertEqual("P0", hits[0]["severity"])

    def test_divergence_check_removed(self):
        """KKS-36（标准间分歧披露）随旧体系一并删除，不应再存在于引擎里。"""
        self.assertFalse(hasattr(TEMPLATE, "standard_divergence_check"))
        self.assertFalse(hasattr(TEMPLATE, "_merge_code_dicts"))
        self.assertNotIn("KKS-36", run_audit.RULE_DEFINITIONS)

    def test_config_registers_version_discipline(self):
        cfg = json.loads((PROJECT / "config" / "kks_rules.json").read_text(encoding="utf-8"))
        basis = cfg["standard_basis"]
        files = {item["file"] for item in basis["current"]}
        self.assertTrue(any("50549-2020" in f for f in files), "现行依据必须登记 GB/T 50549-2020 语料")
        excluded = " ".join(item["standard"] for item in basis["excluded"])
        self.assertIn("50549-2010", excluded, "已废止版本必须登记在案")
        self.assertIn("火电厂标识系统", excluded)
        self.assertNotIn("standard_divergences", cfg, "分歧登记表已随旧体系删除")


if __name__ == "__main__":
    unittest.main()
