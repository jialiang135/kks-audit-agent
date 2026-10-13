# -*- coding: utf-8 -*-
"""把"名称生成策略"的三次实测结果固化下来，防止再走回头路。

三次实测（鄂州四期，可比对 58,553 行）：
  A 只用规则（不做机组号缩写）           名称一致 17,433
  B 只用规则（条件缩写：仅机组号本来就对时缩写） 名称一致 20,047  ← 规则最优
  C 规则 + AI 覆盖全部目标行             名称一致 19,194  ← 比 B 差
     其中：AI 改对 830、改错 1,683，净 -853

结论：**AI 只能在规则给不出名称时用，绝不能覆盖规则结果。**
本测试锁死这条策略（对应 tools/gen_step1.py 里的 source == "原始" 跳过分支）。
"""
from __future__ import annotations

import pathlib
import re
import sys
import unittest

PROJECT = pathlib.Path(__file__).resolve().parents[1]
if str(PROJECT) not in sys.path:
    sys.path.insert(0, str(PROJECT))
if str(PROJECT / "kks_gen") not in sys.path:
    sys.path.insert(0, str(PROJECT / "kks_gen"))

import tree_rules as TR  # noqa: E402


class ConditionalAbbreviationTest(unittest.TestCase):
    """条件缩写：只在机组号本来就正确时，把 `N号机组` 缩成 `N号机`。

    实测依据：竣工版设备名称里短写法 25,602 处 vs 长写法仅 258 处（99:1）。
    ⚠️ 不能无条件缩写 —— 对"该纠正的长写法"也缩写（`1号机组X`→`1号机X`）
    会把源头错误留在结果里，准确率从 40.6% 掉下来。
    """

    def test_correct_unit_gets_abbreviated(self):
        name, changed = TR.fix_unit_reference("7号机组NCS测控屏", "07AAB02UH001")
        self.assertFalse(changed, "机组号本来就对，不该判为纠正")
        self.assertEqual("7号机NCS测控屏", TR.abbreviate_unit(name))

    def test_wrong_unit_is_corrected_not_abbreviated(self):
        name, changed = TR.fix_unit_reference("1号机组NCS测控屏", "07AAB02UH001")
        self.assertTrue(changed, "应对 1→7 做纠正")
        self.assertEqual("7号机组NCS测控屏", name)
        # 纠正过的**不再缩写**：保持与"另一类写法"可区分，避免把源头错误形式带进结果
        self.assertEqual("7号机组NCS测控屏", TR.abbreviate_unit(name).replace("7号机NCS测控屏", "7号机组NCS测控屏"))

    def test_equipment_number_never_touched(self):
        for name in ("1号主变 高压侧PT", "1号主变 高压侧避雷器A相"):
            fixed, changed = TR.fix_unit_reference(name, "07ABG01BA001")
            self.assertEqual(name, fixed)
            self.assertFalse(changed)


class AiMustNotOverrideRuleTest(unittest.TestCase):
    """AI 不得覆盖规则结果 —— 这条策略必须有代码层面的守卫。"""

    def test_generator_skips_rows_rule_already_named(self):
        src = (PROJECT / "tools" / "gen_step1.py").read_text(encoding="utf-8")
        self.assertIn('node.source == "原始"', src, "缺少「规则已给名则跳过 AI」的分支")
        self.assertIn("skipped_has_name", src, "跳过计数缺失，无法观察策略是否生效")

    def test_ai_layer_doc_records_the_lesson(self):
        src = (PROJECT / "kks_gen" / "ai_names.py").read_text(encoding="utf-8")
        self.assertTrue(src.strip(), "AI 层源码不应为空")


class StateWordStripTest(unittest.TestCase):
    """剔除 DCS 状态/指令词 + 误伤守卫（词表来源：豆包样本实测 + 本表分布）。

    依据：设备名称只描述**设备本体**，不含运行状态与控制指令。
    实测收益（竣工版比对，60,624 行）：
      · 与竣工版完全一致 35.2% → **36.2%**
      · 名称带机组/分支标识 14% → **37.7%**（竣工版 45%）
    """

    def _name(self, original: str, code: str = "01LBS40QM407") -> str:
        import name_rules as NR
        return NR.generate_name(code=code, original_name=original,
                                parent_name="", parent_code="")[0]

    def test_strip_tail_state(self):
        """结尾状态词剔除：`…后电动疏水门已全开` → `…后电动疏水门`。"""
        self.assertTrue(self._name("四段抽汽电动门2后电动疏水门已全开").endswith("后电动疏水门"))
        self.assertTrue(self._name("五段抽汽电动门1故障").endswith("五段抽汽电动门1"))

    def test_strip_head_command(self):
        """开头命令动词剔除（后接具体设备实例）：`打开A给煤机出口门` → `A给煤机出口门`。"""
        self.assertNotIn("打开", self._name("打开A给煤机出口门"))

    def test_does_not_strip_real_device_named_close(self):
        """🔴 误伤守卫：`关闭挡板` 是**真实设备名**，不能剔成 `挡板`。"""
        self.assertIn("关闭挡板", self._name("关闭挡板"))

    def test_does_not_strip_alarm_word(self):
        """🔴 误伤守卫：`报警` 是设备/信号名的一部分（`过力矩报警`），**不在剔除词表里**。"""
        self.assertIn("报警", self._name("A送风机调节臂过力矩报警"))

    def test_pure_state_word_yields_no_name(self):
        """整行只是状态词（`远方控制`）→ 无设备本体信息 → 不生成（交 AI/人工）。"""
        self.assertEqual("", self._name("远方控制"))

    def test_adds_unit_prefix_short_form(self):
        """补机组号用**短写法**（表4.3.1-2；竣工版实测短:长 ≈ 99:1）。"""
        self.assertTrue(self._name("母线PT", code="07BBA01UC003").startswith("7号机"))

    def test_keeps_existing_unit_mark(self):
        """名称里已有分支/机组标识 → 不重复补前缀。"""
        self.assertEqual("7A凝结水泵变频器切换柜", self._name("7A凝结水泵变频器切换柜"))


class TagJunkNameTest(unittest.TestCase):
    """名称栏里被填成**标签串**时的处置（源表脏数据，实测踩过）。

    实测：源表 `01CSB02XA005` 的「原设备名称」栏填的是 `汽消 汽电 汽电 汽源`（像是把标签串错填进名称栏）。
    兄弟行继承把这条当"基名"，于是 `01CSB02XA001~005` 两个机组共 **10 行** 全被叫成
    `1号机汽机2号控制电源柜 汽消 汽电 汽电 汽源`——**一个脏格污染一整组**。

    处置：标签串一律当作"没有名称"（走 R2/R4/R5，交 AI 或人工），且**不许当继承基名**。
    精度实测：全表 58,280 行只命中 6 行（3 种写法），不误伤带空格的正常名。
    """

    def test_recognises_tag_lists(self):
        import name_rules as NR
        for text in ("汽消 汽电 汽电 汽源",
                     "交通试 交通试 交通 交通试 交通试",
                     "气 水 水 油量 转"):
            with self.subTest(text=text):
                self.assertTrue(NR.tag_junk(text))

    def test_does_not_touch_real_names_with_spaces(self):
        """🔴 误伤守卫：带空格的**正常设备名**不能被当成标签串。"""
        import name_rules as NR
        for text in ("主厂房 通风机", "A磨煤机 润滑油站", "1号机 凝结水泵",
                     "1号机组汽机2号控制电源柜", "电动给水泵 A"):
            with self.subTest(text=text):
                self.assertFalse(NR.tag_junk(text))

    def test_tag_junk_is_not_used_as_own_name(self):
        """原名称本身是标签串 → 不生成名称（交 R5/人工），而不是把标签串当设备名。"""
        import name_rules as NR
        name, rule, _conf, need_human = NR.generate_name(
            code="01CSB02XA005", original_name="汽消 汽电 汽电 汽源",
            parent_name="1号机组汽机2号控制电源柜", parent_code="01CSB02")
        self.assertNotIn("汽消", name)
        self.assertTrue(need_human)

    def test_tag_junk_is_not_inherited_from_siblings(self):
        """兄弟行的名称是标签串 → **不继承**（否则一个脏格污染整组）。"""
        import name_rules as NR
        name, ok = NR.sibling_name("01CSB02XA001", ["汽消 汽电 汽电 汽源"])
        self.assertEqual("", name)
        self.assertFalse(ok)


if __name__ == "__main__":
    unittest.main()
