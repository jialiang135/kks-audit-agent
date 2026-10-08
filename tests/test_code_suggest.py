# -*- coding: utf-8 -*-
"""问题行「AI 建议编码」与跨标准分类码冲突的回归测试。

覆盖三件事：
1. 生成范围——只对"有问题的行"生成；
2. 准确性防线——建议编码必须过本地复核（骨架/禁用字母/两套标准字典）；
3. 冲突处置——同码在两套标准里含义不同时不判错，改为披露；两套字典取并集判定。
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


def _load_template():
    spec = importlib.util.spec_from_file_location("audit_template", PROJECT / "kks-audit" / "scripts" / "audit_template.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


TEMPLATE = _load_template()
COLS = {"kks": 0, "name": 1}


class SuggestDictTest(unittest.TestCase):
    """候选码与本地复核。"""

    @classmethod
    def setUpClass(cls):
        cls.sys_codes, cls.dev_codes = A._suggest_dicts()

    def test_dicts_cover_both_standards(self):
        index = json.loads((PROJECT / "config" / "kks_code_index.json").read_text(encoding="utf-8"))
        self.assertGreater(len(self.sys_codes), len(index["system_codes"]))
        self.assertGreater(len(self.dev_codes), len(index["device_codes"]))
        # 同码异义的两个释义都要在候选里出现，否则模型只知道一套标准
        self.assertIn("国标", self.dev_codes["GP"])
        self.assertIn("企标", self.dev_codes["GP"])

    def test_enterprise_code_is_a_candidate(self):
        """企标语义的码必须能被检索出来（GP=照明检修配电箱 是企标定义）。"""
        got = [c["code"] for c in A._candidate_codes("照明检修配电箱", self.dev_codes, 5)]
        self.assertIn("GP", got)

    def test_pump_name_matches_pump_codes(self):
        got = [c["code"] for c in A._candidate_codes("给水泵", self.dev_codes, 5)]
        self.assertTrue({"GP", "AP"} & set(got), f"给水泵 未匹配到泵类码：{got}")

    def test_validation_rejects_violations(self):
        bad_cases = {
            "00LBA10II001": "禁用字母",
            "00LBA10AP00": "长度",
            "00LBA10AP0011": "长度",
            "00LB110AP001": "系统分类码",   # 第 5 位是数字 → 系统分类码段不是纯字母
            "00LBA10AP0O1": "禁用字母",
        }
        for code, expect in bad_cases.items():
            ok, reason = A._validate_suggested(code, self.sys_codes, self.dev_codes)
            self.assertFalse(ok, f"{code} 应被拒绝")
            self.assertIn(expect, reason, f"{code} 的拒绝原因应包含「{expect}」，实际：{reason}")

    def test_validation_accepts_legal_codes(self):
        for code in ("00LBA10AP008", "Y0XAG10AN001", "00LBA10AA001"):
            ok, reason = A._validate_suggested(code, self.sys_codes, self.dev_codes)
            self.assertTrue(ok, f"{code} 应通过校验，实际：{reason}")

    def test_validation_rejects_code_missing_from_dicts(self):
        ok, reason = A._validate_suggested("00LBA10ZZ001", self.sys_codes, self.dev_codes)
        self.assertFalse(ok)
        self.assertIn("设备分类码", reason)


class SuggestScopeTest(unittest.TestCase):
    """只对有问题的行生成；开关关闭时不调用模型。"""

    def _result(self):
        return {
            "issues": [
                {"excel_row": 10, "kks_code": "00LBA10GG001", "name": "轴流风机", "priority": "P2",
                 "rule_id": "KKS-33", "category": "设备分类码未收录", "message": "m", "suggestion": "s"},
                {"excel_row": 10, "kks_code": "00LBA10GG001", "name": "轴流风机", "priority": "P2",
                 "rule_id": "KKS-36", "category": "标准间分类码分歧", "message": "m2", "suggestion": "s2"},
                {"excel_row": 11, "kks_code": "", "priority": "P0", "rule_id": "KKS-00",
                 "category": "缺码", "message": "无码", "suggestion": "补码"},
                {"excel_row": 12, "kks_code": "00LBA10GP001", "priority": "P1", "rule_id": "KKS-31",
                 "category": "父子名称语义矛盾", "message": "m3", "suggestion": "s3"},
            ],
            "records": [
                {"excel_row": 10, "kks_code": "00LBA10GG001", "name": "轴流风机", "parent_code": "00LBA10"},
                {"excel_row": 12, "kks_code": "00LBA10GP001", "name": "照明配电箱", "parent_code": "00LBA10"},
            ],
        }

    def test_only_issue_rows_in_scope(self):
        rows = A._suggest_rows(self._result())
        row_ids = sorted(r["row"] for r in rows)
        self.assertEqual(["10", "12"], row_ids, "无码行（KKS-00）不应进入建议编码范围")

    def test_same_row_issues_are_merged(self):
        rows = {r["row"]: r for r in A._suggest_rows(self._result())}
        self.assertEqual(2, len(rows["10"]["rules"]), "同一行的多条问题应合并到一条待生成记录")

    def test_payload_carries_clause_and_candidates(self):
        row = A._suggest_rows(self._result())[0]
        payload = A._suggest_payload_row(row, self.sys_codes if hasattr(self, "sys_codes") else A._suggest_dicts()[0],
                                         A._suggest_dicts()[1])
        self.assertIn("命中规则", payload)
        self.assertIn("设备分类码候选", payload)
        self.assertTrue(payload["当前分类码含义"], "应给出当前分类码在两套标准里的含义")

    def test_disabled_by_env_switch(self):
        old = os.environ.get("AI_CODE_SUGGEST_ENABLED")
        os.environ["AI_CODE_SUGGEST_ENABLED"] = "0"
        try:
            result = A.suggest_correct_codes(self._result())
        finally:
            if old is None:
                os.environ.pop("AI_CODE_SUGGEST_ENABLED", None)
            else:
                os.environ["AI_CODE_SUGGEST_ENABLED"] = old
        self.assertEqual("disabled", result["status"])
        self.assertEqual({}, result["suggestions"])

    def test_disabled_when_ai_disabled(self):
        keys = {k: os.environ.get(k) for k in ("AI_ENABLED", "AI_API_KEY")}
        os.environ["AI_ENABLED"] = "false"
        try:
            result = A.suggest_correct_codes(self._result())
        finally:
            for k, v in keys.items():
                if v is None:
                    os.environ.pop(k, None)
                else:
                    os.environ[k] = v
        self.assertEqual("disabled", result["status"])


class SuggestReportTest(unittest.TestCase):
    """报告层：列插入位置、HTML/Excel 呈现。"""

    def _result_with_suggestion(self):
        result = {
            "source_file": "demo.xlsx",
            "data_rows": 2,
            "issues": [
                {"excel_row": 10, "kks_code": "00LBA10GP001", "name": "照明配电箱", "priority": "P1",
                 "rule_id": "KKS-31", "category": "父子名称语义矛盾", "message": "m", "suggestion": "s",
                 "status": "open"},
            ],
            "records": [{"excel_row": 10, "kks_code": "00LBA10GP001", "name": "照明配电箱"}],
            "_ai_context": {"records": [{"excel_row": 10, "kks_code": "00LBA10GP001", "name": "照明配电箱"}]},
            "code_suggestions": {
                "status": "completed", "model": "demo", "requested": 1, "validated": 1, "rejected": 0,
                "suggestions": {"10": {"row": "10", "suggested_code": "00LBA10AA001", "changed": "A1A2",
                                       "basis": ["GB/T 50549-2020 表5.1.4"],
                                       "reason": "阀门应为 AA", "confidence": 0.8,
                                       "valid": True, "invalid_reason": ""}},
            },
        }
        return result

    def test_column_inserted_right_after_code(self):
        ov = RB.build_overview(self._result_with_suggestion())
        buckets = [b for b in ov["buckets"] if b["items"]]
        self.assertTrue(buckets, "应至少有一个含明细的业务桶")
        bucket = buckets[0]
        kinds = [c["k"] for c in bucket["columns"]]
        self.assertIn("ai_suggested_code", kinds)
        self.assertEqual(kinds.index("kks_code") + 1, kinds.index("ai_suggested_code"),
                         "AI建议编码 应紧跟在「编码」列之后")
        self.assertEqual("00LBA10AA001", bucket["items"][0]["ai_suggested_code"])

    def test_no_column_without_suggestions(self):
        result = self._result_with_suggestion()
        result["code_suggestions"] = {"status": "completed", "suggestions": {}}
        ov = RB.build_overview(result)
        for bucket in ov["buckets"]:
            self.assertNotIn("ai_suggested_code", [c["k"] for c in bucket["columns"]])

    def test_html_and_xlsx_contain_section(self):
        result = self._result_with_suggestion()
        ov = RB.build_overview(result)
        html = RB.render_html(result, ov)
        self.assertIn("十、AI 建议编码", html)
        self.assertIn("AI建议编码", html)
        # 写到 .gitignore 掉的 _tmp_extract/ 下，避免测试过程触发沙箱的安全删除拦截
        out_dir = PROJECT / "_tmp_extract"
        out_dir.mkdir(parents=True, exist_ok=True)
        out = out_dir / "test_code_suggest.xlsx"
        RB.render_xlsx(out, result, ov)
        import openpyxl
        wb = openpyxl.load_workbook(out, read_only=True)
        self.assertIn("AI建议编码", wb.sheetnames)


class CrossStandardConflictTest(unittest.TestCase):
    """两套标准的分类码冲突：并集判定 + 分歧披露，不判错。"""

    def test_enterprise_codes_not_flagged(self):
        """企标合法码（AA 阀门 / AP 泵 / AN 风机）不得被判为未收录。"""
        rows = [["00LAB10AA001", "电动阀门"], ["00LAC10AP008", "给水泵"], ["00XAG10AN001", "轴流风机"]]
        hits = TEMPLATE.device_class_code_legality(COLS, rows)
        self.assertEqual([], hits, f"企标码被误报：{hits}")

    def test_unknown_code_still_flagged_as_pending(self):
        hits = TEMPLATE.device_class_code_legality(COLS, [["00XAG10GG001", "轴流风机"]])
        self.assertEqual(["GG"], [h["segment"] for h in hits])
        self.assertEqual("P2", hits[0]["severity"], "未收录只能是待核，不能判成错误")

    def test_forbidden_letter_is_hard_error(self):
        hits = TEMPLATE.device_class_code_legality(COLS, [["00LBA10QI001", "给水阀"]])
        self.assertEqual("P0", hits[0]["severity"])

    def test_divergence_is_disclosed_not_judged(self):
        hits = TEMPLATE.standard_divergence_check(COLS, [["00LBA10GP001", "照明检修配电箱"]])
        self.assertTrue(hits, "GP 属于两套标准同码异义，应被披露")
        self.assertTrue(all(h["severity"] == "P2" for h in hits))
        self.assertIn("这不是错误", hits[0]["suggestion"])
        self.assertIn("国标", hits[0]["message"])
        self.assertIn("企标", hits[0]["message"])

    def test_config_registers_divergences(self):
        cfg = json.loads((PROJECT / "config" / "kks_rules.json").read_text(encoding="utf-8"))
        entries = cfg["standard_divergences"]["code_level_divergences"]
        scopes = {e["scope"] for e in entries}
        self.assertIn("设备分类码主组 T1", scopes)
        self.assertIn("阀门类", scopes)


if __name__ == "__main__":
    unittest.main()
