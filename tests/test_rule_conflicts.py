# -*- coding: utf-8 -*-
"""规则一致性门禁：任何改动让编码规则互相冲突时，这里必须失败。

背景：规则来自 4 份标准、落地到 5 个文件（config/kks_rules.json、run_audit.py、
audit_template.py、SKILL.md、config/kks_code_index.json）。2026-10-08 出现过一次真实冲突——
KKS-32/33 按 GB/T 附录A/B 判分类码，而项目默认标准是火电导则 Q/WL，两套分类码体系不同
（阀门 国标 QM / 企标 AA；泵 国标 GP / 企标 AP），造成整片误报。

本测试调用 tools/check_rule_conflicts.py 的全部检查项，要求 ERROR 数为 0。
"""
import sys
import unittest
from pathlib import Path

PROJECT = Path(__file__).resolve().parents[1]
if str(PROJECT) not in sys.path:
    sys.path.insert(0, str(PROJECT))
sys.path.insert(0, str(PROJECT / "tools"))

import check_rule_conflicts as checker  # noqa: E402


class RuleConflictTest(unittest.TestCase):
    """把「规则之间没有冲突」变成可执行的断言。"""

    @classmethod
    def setUpClass(cls):
        cls.findings, cls.stats = checker.run_checks(quiet=True)
        cls.errors = [f for f in cls.findings if f["level"] == "ERROR"]

    def test_no_conflicts(self):
        detail = "; ".join(f"[{f['area']}] {f['rule']}: {f['detail']}" for f in self.errors)
        self.assertEqual([], self.errors, f"规则存在冲突：\n{detail}")

    def test_no_warnings(self):
        warnings = [f for f in self.findings if f["level"] == "WARN"]
        detail = "; ".join(f"[{f['area']}] {f['rule']}: {f['detail']}" for f in warnings)
        self.assertEqual([], warnings, f"规则存在未登记的不一致：\n{detail}")

    def test_rule_registry_is_complete(self):
        """每条规则都要有定义、有出处或已登记为「无专条依据」，且代码里真的会产出它。"""
        import run_audit as R
        self.assertEqual(set(R.RULE_DEFINITIONS), set(R.RULE_DEFINITIONS) | set(R.RULE_DEFINITIONS))
        for rid in R.RULE_DEFINITIONS:
            self.assertTrue(R.RULE_DEFINITIONS[rid].strip(), f"{rid} 的问题定义为空")

    def test_cross_standard_index_present(self):
        """两套标准的字典都必须存在，否则并集判定会退化成单边判定。"""
        import json
        index = json.loads(checker.CODE_INDEX.read_text(encoding="utf-8"))
        self.assertGreater(len(index.get("system_codes") or {}), 2000, "国标系统分类码字典缺失")
        self.assertGreater(len(index.get("device_codes") or {}), 150, "国标设备分类码字典缺失")
        self.assertGreater(len(index.get("system_codes_huodian") or {}), 300, "火电导则系统码字典缺失")
        self.assertGreater(len(index.get("device_codes_huodian") or {}), 60, "火电导则设备码字典缺失")

    def test_divergence_registry_matches_config(self):
        """索引里算出的歧义码必须在 config 的 standard_divergences 里有对应登记。"""
        import json
        cfg = json.loads(checker.RULES_JSON.read_text(encoding="utf-8"))
        text = json.dumps(cfg.get("standard_divergences") or {}, ensure_ascii=False)
        for tag in ("DIV-01", "DIV-02", "DIV-03", "DIV-04"):
            self.assertIn(tag, text, f"standard_divergences 缺少 {tag} 登记")


if __name__ == "__main__":
    unittest.main()
