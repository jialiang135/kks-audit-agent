# -*- coding: utf-8 -*-
"""国标分类码语义规则回归测试（KKS-32 / KKS-33 / KKS-35）。

规则依据：
  GB/T 50549-2020 4.4.2（系统分类码按附录A）、4.4.3（设备分类码按附录B）、
  4.2.1 注（禁用 I、O）、表5.1.4（阀门设备分类码）、表5.1.5（QM/WP 编号分段）。
分类码字典由 tools/build_code_index.py 从两份标准附录生成到 config/kks_code_index.json。

测试锚定三点：
  1. 合法码不得误报（拿真实数据校准过：舟山暖通 1190 行系统码 0 告警）；
  2. 含禁用字母是硬错误，必须逐行报且为 P0；
  3. 其余"未收录"一律 P2 待核，且**按分类码聚合**（同一码只出一条，避免刷屏）。
"""
import unittest

import run_audit
from run_audit import rule_source

RULES = run_audit.SKILL_RULES


def _cols():
    return {"kks": 0, "name": 1, "parent": None, "old": None, "dtype": None, "unit": None}


def _rows(*pairs):
    return [[code, name] for code, name in pairs]


class CodeIndexTests(unittest.TestCase):
    def test_index_is_loaded(self):
        index = RULES._load_code_index()
        self.assertTrue(index, "config/kks_code_index.json 应可加载")
        self.assertGreater(len(index.get("system_codes", {})), 1000)
        self.assertGreater(len(index.get("device_codes", {})), 100)
        self.assertIn("A", index.get("system_s1_groups", {}))

    def test_t1_groups_cover_a_and_b(self):
        """附录B 分页导致 A/B 主组缺失，已按两份文档互证补全——补全后不应再误报 A/B 主组。"""
        groups = RULES._load_code_index().get("device_t1_groups", {})
        self.assertIn("A", groups)
        self.assertIn("B", groups)


class SystemClassCodeTests(unittest.TestCase):
    def test_valid_code_is_not_reported(self):
        # LBA（主蒸汽管道系统）是 GB/T 附录A 收录的系统分类码
        hits = RULES.system_class_code_legality(_cols(), _rows(("10LBA10GP008", "给水泵")))
        self.assertEqual(hits, [])

    def test_unknown_code_is_reported_as_pending(self):
        hits = RULES.system_class_code_legality(_cols(), _rows(("10QQQ10GP001", "给水泵")))
        self.assertEqual(len(hits), 1)
        self.assertEqual(hits[0]["severity"], "P2", "未收录只能是待核线索，不能判为错误")
        self.assertIn("QQQ", hits[0]["message"])

    def test_aggregates_by_code_instead_of_per_row(self):
        rows = _rows(*[(f"10QQQ10GP{i:03d}", "给水泵") for i in range(1, 21)])
        hits = RULES.system_class_code_legality(_cols(), rows)
        self.assertEqual(len(hits), 1, "20 行同一个码应只出一条聚合告警")
        self.assertEqual(hits[0]["count"], 20)

    def test_short_codes_are_skipped(self):
        hits = RULES.system_class_code_legality(_cols(), _rows(("10LBA", "系统级")))
        self.assertEqual(hits, [])


class DeviceClassCodeTests(unittest.TestCase):
    def test_valid_code_is_not_reported(self):
        hits = RULES.device_class_code_legality(_cols(), _rows(("10LAC10GP008", "给水泵")))
        self.assertEqual(hits, [])

    def test_forbidden_letter_i_is_hard_error(self):
        hits = RULES.device_class_code_legality(_cols(), _rows(("10LBA10QI001", "给水阀")))
        self.assertEqual(len(hits), 1)
        self.assertEqual(hits[0]["severity"], "P0")
        self.assertIn("禁用字母", hits[0]["message"])

    def test_forbidden_letter_o_is_hard_error(self):
        hits = RULES.device_class_code_legality(_cols(), _rows(("10LBA10QO001", "给水阀")))
        self.assertEqual(hits[0]["severity"], "P0")

    def test_unknown_device_code_is_pending_and_aggregated(self):
        rows = _rows(*[(f"10LAB10ZZ{i:03d}", "设备") for i in range(1, 6)])
        hits = RULES.device_class_code_legality(_cols(), rows)
        self.assertEqual(len(hits), 1)
        self.assertEqual(hits[0]["severity"], "P2")
        self.assertEqual(hits[0]["count"], 5)

    def test_short_codes_are_skipped(self):
        hits = RULES.device_class_code_legality(_cols(), _rows(("10LBA10", "系统级")))
        self.assertEqual(hits, [])


class ValveSegmentTests(unittest.TestCase):
    """GB/T 表5.1.5：QM（截止阀）、WP（管道）设备编号按功能分段。"""

    def test_general_segment_accepts_any_name(self):
        hits = RULES.valve_pipe_number_segment_check(_cols(), _rows(("10LAB20QM110", "20系统疏水阀")))
        self.assertEqual(hits, [], "001～299 为通用段，名称含疏水不应报")

    def test_matching_segment_is_not_reported(self):
        rows = _rows(("10LAB20QM450", "20系统疏水阀"), ("10LAB20WP620", "20系统取样管"))
        self.assertEqual(RULES.valve_pipe_number_segment_check(_cols(), rows), [])

    def test_contradiction_is_reported_as_pending(self):
        hits = RULES.valve_pipe_number_segment_check(_cols(), _rows(("10LAB20QM350", "20系统疏水阀")))
        self.assertEqual(len(hits), 1)
        self.assertEqual(hits[0][4], "P2")
        self.assertIn("疏水", hits[0][2])

    def test_non_valve_codes_are_ignored(self):
        hits = RULES.valve_pipe_number_segment_check(_cols(), _rows(("10LAB10GP350", "20系统疏水阀")))
        self.assertEqual(hits, [], "只对 QM/WP 生效")


class RuleSourceTests(unittest.TestCase):
    def test_new_rules_cite_national_standard(self):
        for rule_id in ("KKS-32", "KKS-33", "KKS-35"):
            for standard in ("huodian", "wind", "solar"):
                src = rule_source(rule_id, standard)
                self.assertIn("50549-2020", src, f"{rule_id}/{standard} 应引用国标")

    def test_valve_rule_also_cites_group_rule(self):
        self.assertIn("8.1.5", rule_source("KKS-35", "huodian"))

    def test_definitions_exist(self):
        for rule_id in ("KKS-32", "KKS-33", "KKS-34", "KKS-35"):
            self.assertTrue(run_audit.RULE_DEFINITIONS.get(rule_id))


if __name__ == "__main__":
    unittest.main()
