# -*- coding: utf-8 -*-
"""KKS 设备树生成规则回归测试。

这些规则全部由竣工版实测反推得到，错了会静默产生大量错编码（实测曾出现
上级码准确率从 100% 掉到 6.8% 的事故），所以必须用测试锁死。
"""
from __future__ import annotations

import pathlib
import sys
import unittest

PROJECT = pathlib.Path(__file__).resolve().parents[1]
if str(PROJECT) not in sys.path:
    sys.path.insert(0, str(PROJECT))
if str(PROJECT / "kks_gen") not in sys.path:
    sys.path.insert(0, str(PROJECT / "kks_gen"))

import component_rules as CR  # noqa: E402
import name_rules as NR  # noqa: E402
import tree_rules as TR  # noqa: E402


class ParentMappingTest(unittest.TestCase):
    """父级映射：设备分类码不单独成节点，后缀分三类挂法。"""

    def test_device_class_code_is_not_a_level(self):
        """07AAB02UH001 的上级是 7 位子系统，不是 9 位设备组。"""
        self.assertEqual("07AAB02", TR.direct_parent("07AAB02UH001"))

    def test_dot_suffix_hangs_under_its_stem(self):
        """点后缀挂在主干设备下（竣工版行 559）。"""
        self.assertEqual("07BFA01UC001", TR.direct_parent("07BFA01UC001.A1"))

    def test_letter_suffix_hangs_under_subsystem(self):
        """字母尾缀是分相，直接挂 7 位子系统（竣工版行 60）。"""
        self.assertEqual("07ABG01", TR.direct_parent("07ABG01BA001A"))

    def test_unit_code_has_no_parent(self):
        """2 位机组码到顶，绝不能拼出 '0'。"""
        self.assertEqual("", TR.direct_parent("07"))
        self.assertEqual("", TR.direct_parent("08"))

    def test_gb_device_code_plus_component_suffix(self):
        """国标设备码 + 部件码（现场汇总表省略了连字符）→ 上级是那台设备。

        实测（某汇总表 16,675 行里 16,504 行属这种）：
        `01HAL11GP001XB21` = 设备 01HAL11GP001 + 部件 XB21
        修复前这批一律推不出上级（44.7% 无上级），补这条后降到 5.9%。
        """
        self.assertEqual("01HAL11GP001", TR.direct_parent("01HAL11GP001XB21"))
        self.assertEqual("02HAL11GP001", TR.direct_parent("02HAL11GP001XB41"))
        self.assertEqual("01HAL11GP001", TR.direct_parent("01HAL11GP001XB48"))

    def test_gb_device_code_plus_underscore_seq(self):
        """设备码 + 下划线序号：`01HDU10BP101_1` → 上级 `01HDU10BP101`。"""
        self.assertEqual("01HDU10BP101", TR.direct_parent("01HDU10BP101_1"))
        self.assertEqual("01HDU10BP101", TR.direct_parent("01HDU10BP101_2"))

    def test_long_code_without_valid_head_is_not_forced(self):
        """前 12 位不是合法设备码时，不要硬造上级（宁可留空交人工）。"""
        self.assertEqual("", TR.direct_parent("01BTM11JYZZGZ"))     # 13 位、头不合法
        self.assertEqual("", TR.direct_parent("01HFA01CX01F11"))    # 14 位、头不是设备码

    def test_system_child_with_dash(self):
        """5 位系统码下的设备（连字符前只有 5 位）：上级是系统码本身。

        实测依据：竣工版 `07CGQ-N001`「7号炉1号炉管泄露探头」，上级应为 `07CGQ`。
        这与"12 位设备下的部件"是两种不同形态，不能都当部件处理。
        """
        self.assertEqual("07CGQ", TR.direct_parent("07CGQ-N001"))
        self.assertEqual("07CGQ", TR.direct_parent("07CGQ-D001"))
        self.assertEqual("07AAB02UH001", TR.direct_parent("07AAB02UH001-KF01"))

    def test_parent_len_chain(self):
        self.assertEqual("07AAB", TR.direct_parent("07AAB02"))
        self.assertEqual("07A", TR.direct_parent("07AAB"))
        self.assertEqual("07", TR.direct_parent("07A"))

    def test_ancestor_chain_order(self):
        """上级链由近到远，且不含 9 位设备组。"""
        self.assertEqual(["07AAB02", "07AAB", "07A", "07"], TR.ancestor_chain("07AAB02UH001"))


class UnitNumberTest(unittest.TestCase):
    """机组号口径：07=7号机组；设备编号不能被误改。"""

    def test_fix_wrong_unit_reference(self):
        fixed, changed = TR.fix_unit_reference("1号机NCS测控屏", "07AAB02UH001")
        self.assertEqual("7号机NCS测控屏", fixed)
        self.assertTrue(changed)

    def test_fix_short_form(self):
        fixed, _ = TR.fix_unit_reference("2号机远动采集控制屏", "08AAC01UH001")
        self.assertEqual("8号机远动采集控制屏", fixed)

    def test_equipment_number_must_not_change(self):
        """'1号主变' 里的 1号 是设备编号，改了就是错。"""
        for name in ("1号主变 高压侧PT", "1号主变 高压侧避雷器A相", "500kV 1号主变进线间隔系统"):
            fixed, changed = TR.fix_unit_reference(name, "07ABG01BA001")
            self.assertEqual(name, fixed, f"{name} 不该被改")
            self.assertFalse(changed)

    def test_already_correct_untouched(self):
        fixed, changed = TR.fix_unit_reference("7号机NCS测控屏", "07AAB02UH001")
        self.assertEqual("7号机NCS测控屏", fixed)
        self.assertFalse(changed)


class NameRuleTest(unittest.TestCase):
    """名称生成：不做过度归一化。"""

    def test_no_unit_abbreviation(self):
        """'号机组' 不能缩写成 '号机'：竣工版两种写法并存。"""
        self.assertEqual("7号机NCS测控屏系统", TR.normalize_name("7号机NCS测控屏系统"))
        self.assertEqual("7号机组NCS测控屏", TR.normalize_name("7号机组NCS测控屏"))

    def test_fullwidth_letters_normalized(self):
        self.assertEqual("7号机DEH(A)控制柜", TR.normalize_name("7号机DEH（A）控制柜").replace("（", "(").replace("）", ")"))

    def test_component_word_lookup(self):
        self.assertEqual("CT", NR.component_type_word("07BBA00UC001-BC01"))
        self.assertEqual("过电压保护器", NR.component_type_word("07BBA00UC001-FA01"))
        self.assertEqual("接地刀闸", NR.component_type_word("07BBA00UC001-QC01"))

    def test_phase_suffix_no_space(self):
        """分相后缀不加空格：'避雷器A相' 而不是 '避雷器 A相'。"""
        self.assertEqual("A相", NR.phase_suffix("07ABG01FA001A"))
        self.assertEqual("", NR.phase_suffix("07ABG01FA001"))

    def test_sibling_name_inherits_base(self):
        name, ok = NR.sibling_name("07ABG01FA001A", ["1号主变 高压侧避雷器", "1号主变 高压侧避雷器C相"])
        self.assertTrue(ok)
        self.assertEqual("1号主变 高压侧避雷器A相", name)

    def test_generate_name_prefers_original(self):
        name, rule, conf, human = NR.generate_name(
            code="07AAB02UH001", original_name="7号机NCS测控屏",
            parent_name="1号机NCS测控屏系统", parent_code="07AAB02",
        )
        self.assertEqual("7号机NCS测控屏", name)
        self.assertIn("R1", rule)
        self.assertFalse(human)


class ComponentRuleTest(unittest.TestCase):
    """部件配对与编号。"""

    def test_expected_components_by_device_class(self):
        self.assertIn("KF", CR.expected_components("07AAB02UH001"))
        self.assertIn("MA", CR.expected_components("07BBA01QM001"))
        self.assertEqual([], CR.expected_components("07AAB"))     # 层级不够

    def test_component_kks_format(self):
        self.assertEqual("07AAB02UH001-KF01", CR.component_kks("07AAB02UH001", "KF", 1))
        self.assertEqual("07AAB02UH001-QA12", CR.component_kks("07AAB02UH001", "QA", 12))

    def test_component_name_uses_parent(self):
        self.assertEqual("7号机NCS测控屏开关", CR.component_name("7号机NCS测控屏", "QA"))
        self.assertEqual("", CR.component_name("7号机NCS测控屏", "ZZ"))


class NoSpeculativeComponentsTest(unittest.TestCase):
    """🔴 锁死负面结论：不许自动"猜"部件。

    实测（鄂州四期 85,754 个设备）两条路都是 0% 精确率：
      - 兄弟签名参照：全量一个都补不出（有部件的设备里 5,088 个只有 1 个部件，兄弟也没有）
      - 设备类型/名称关键词推断：补出 3,830 个，精确率 0%
    所以 plan_components 默认必须返回空；要探索性生成必须显式传 allow_speculative=True。
    """

    @classmethod
    def setUpClass(cls):
        import pathlib as _p
        import sys as _s
        _s.path.insert(0, str(_p.Path(__file__).resolve().parents[1] / "kks_gen"))

    def test_default_returns_nothing(self):
        import component_plan as CP

        devices = [
            {"kks": "07HLB10GP005", "name": "7A一次风机润滑油1号冷却器", "cls": "GP",
             "parent_kks": "07HLB10", "parent_name": "7A一次风机系统"},
            {"kks": "07HLB10GP009", "name": "7A电动给水泵", "cls": "GP",
             "parent_kks": "07HLB10", "parent_name": "7A给水系统"},
        ]
        self.assertEqual([], CP.plan_components(devices, {}))

    def test_speculative_requires_explicit_flag_and_is_flagged(self):
        import component_plan as CP

        devices = [{"kks": "07HLB10GP009", "name": "7A电动给水泵", "cls": "GP",
                    "parent_kks": "07HLB10", "parent_name": "7A给水系统"}]
        plans = CP.plan_components(devices, {}, allow_speculative=True)
        self.assertTrue(plans, "显式开启时应能生成探索性结果")
        for p in plans:
            self.assertTrue(p.need_human, "探索性生成必须标需人核")
            self.assertIn("探索性", p.source)
            self.assertLessEqual(p.confidence, 0.5, "探索性结果置信度不能高")

    def test_missing_component_list_still_produced(self):
        """不补部件，但要给出"应编未编"清单交人工。"""
        import component_plan as CP

        devices = [{"kks": "07ABG01QB001", "name": "店50116刀闸", "cls": "QB",
                    "parent_kks": "07ABG01", "parent_name": "500kV 1号主变进线间隔系统"}]
        missing = CP.list_missing_components(devices, {})
        self.assertEqual(1, len(missing))
        self.assertIn("带部件", missing[0]["reason"])



class BuildTreeTest(unittest.TestCase):
    """整树构建：父级链完整、每个种子行都有输出。"""

    def _seeds(self):
        return [
            {"excel_row": 6, "design_code": "07AAB02UH001", "original_name": "7号机组NCS测控屏",
             "profession": "电气二次", "drawing": "F1284S-D0702-10"},
            {"excel_row": 19, "design_code": "07AAC01UH001", "original_name": "7号机远动采集控制屏",
             "profession": "电气二次", "drawing": ""},
        ]

    def test_every_seed_gets_full_columns(self):
        result = TR.build_tree(self._seeds())
        nodes = result["nodes"]
        for seed in self._seeds():
            code = seed["design_code"]
            self.assertIn(code, nodes)
            node = nodes[code]
            self.assertTrue(node.kks, "KKS编码不能为空")
            self.assertTrue(node.name, "设备名称不能为空")
            self.assertTrue(node.parent_kks, "上级KKS编码不能为空")

    def test_parent_nodes_created_without_design_code(self):
        nodes = TR.build_tree(self._seeds())["nodes"]
        for expected in ("07AAB02", "07AAB", "07A", "07"):
            self.assertIn(expected, nodes)
            self.assertFalse(nodes[expected].design_code, f"{expected} 应是父节点（无设计编码）")

    def test_parent_name_derived_from_children(self):
        nodes = TR.build_tree(self._seeds())["nodes"]
        self.assertTrue(nodes["07AAB02"].name, "父节点名称应能由子名称推导")

    def test_hierarchy_is_closed(self):
        """不存在断链：每个非顶节点的 parent_kks 都在表内。"""
        nodes = TR.build_tree(self._seeds())["nodes"]
        for kks, node in nodes.items():
            if kks == "07":
                self.assertEqual("", node.parent_kks)
                continue
            self.assertIn(node.parent_kks, nodes, f"{kks} 的上级 {node.parent_kks} 不在树内")


if __name__ == "__main__":
    unittest.main()
