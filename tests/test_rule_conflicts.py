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

    def test_index_holds_single_current_standard(self):
        """索引必须只含现行标准的字典；残留旧体系数据会让判定与建议编码指向错设备。"""
        import json
        index = json.loads(checker.CODE_INDEX.read_text(encoding="utf-8"))
        self.assertGreater(len(index.get("system_codes") or {}), 2000, "系统分类码字典缺失")
        self.assertGreater(len(index.get("device_codes") or {}), 150, "设备分类码字典缺失")
        self.assertIn("50549-2020", str(index.get("standard") or ""), "索引未声明现行标准")
        for key in index:
            self.assertFalse(key.endswith("_huodian"), f"索引里残留旧体系字典：{key}")
            self.assertNotIn("divergence", key, f"索引里残留分歧清单：{key}")

    def test_version_discipline_registered(self):
        """现行依据与已排除的旧版本都必须在 config 里登记，否则无法机器判定依据是否过期。"""
        import json
        cfg = json.loads(checker.RULES_JSON.read_text(encoding="utf-8"))
        basis = cfg.get("standard_basis") or {}
        self.assertTrue(basis.get("current"), "standard_basis.current 为空")
        self.assertTrue(basis.get("excluded"), "standard_basis.excluded 为空")
        current = " ".join(json.dumps(x, ensure_ascii=False) for x in basis["current"])
        excluded = " ".join(json.dumps(x, ensure_ascii=False) for x in basis["excluded"])
        self.assertIn("50549-2020", current)
        for token in ("50549-2010", "火电厂标识系统"):
            self.assertIn(token, excluded, f"已废止来源未登记：{token}")
        self.assertNotIn("standard_divergences", cfg, "分歧登记表已随旧体系删除")


if __name__ == "__main__":
    unittest.main()
