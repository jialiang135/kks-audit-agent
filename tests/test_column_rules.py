# -*- coding: utf-8 -*-
"""列规则表（`kks_gen/column_rules.py`）的规则级测试：**一条规则一个测试**。

为什么单独测：回填原来只硬编码 4 列，改成"列规则表"后**新增列 = 加一条规则**，
规则本身就成了可回归的单元。这里锁死三类能力：

1. 系统列：码表释义 → **可读短名**（不是照搬整句）
2. 设备类列：设备分类码 → **上位类目**（人看的是实例，不是国标的功能描述）
3. 取不到就**留空 + 说明原因**（不猜；预留段尤其不能当系统名写出去）
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

import column_rules as CR  # noqa: E402


class ReadableNameTest(unittest.TestCase):
    """规则：系统/设备**释义压成可读短名**（4.4.2 附录A 的原文是整句，不能照搬）。"""

    def test_strips_range_note_and_takes_core_noun(self):
        """`用于…系统的抽汽管道系统 / 从(不包括):…` → `抽汽管道系统`。"""
        raw = ("用于主凝结水预加热系统的抽汽管道系统 / 从(不包括):汽轮机出口或其他系统分支"
               "至(不包括):主冷凝热水")
        self.assertEqual("抽汽管道系统", CR.readable_name(raw))

    def test_strips_parenthesised_supplement(self):
        """`给水管道系统(不包括给水泵和给水加热)` → `给水管道系统`。"""
        raw = "给水管道系统(不包括给水泵和给水加热) / 从(不包括):给水箱出口 至(不包括):锅炉入口"
        self.assertEqual("给水管道系统", CR.readable_name(raw))

    def test_keeps_text_outside_brackets(self):
        """括号**在中间**时，括号外正文要保留：`凝结水泵(电动)系统` → `凝结水泵系统`。"""
        self.assertEqual("凝结水泵系统", CR.readable_name("凝结水泵(电动)系统"))

    def test_nested_and_unclosed_brackets(self):
        """OCR 嵌套/不闭合括号：`输电 [标准电压 150(154)kV]` → `输电`（不能留 `kV]`）。"""
        self.assertEqual("输电", CR.readable_name("输电 [标准电压 150(154)kV]"))
        self.assertEqual("输电", CR.readable_name("输电 [标准电压 150(154)kV"))

    def test_voltage_band_keeps_substantive_name(self):
        """电压式后面还有实质名称 → 去电压式留名称：`Un ≤ 400V 交流配电系统` → `交流配电系统`。"""
        self.assertEqual("UPS交流系统", CR.readable_name("Un ≤ 400V UPS 交流系统"))
        self.assertEqual("交流配电系统", CR.readable_name("400V < Un ≤ 1000V 交流配电系统"))
        self.assertEqual("安全应用交流电源系统", CR.readable_name("400V < Un ≤ 1000V 安全应用交流电源系统"))

    def test_voltage_band_without_name_uses_voltage(self):
        """电压式后面只剩用途词 → 电压就是名字：`525kV ≤ Un < 765kV 输电` → `525kV输电系统`。"""
        self.assertEqual("525kV输电系统", CR.readable_name("525kV ≤ Un < 765kV 输电"))
        self.assertEqual("420kV以上系统", CR.readable_name("Un > 420kV 系统"))

    def test_no_mid_token_truncation(self):
        """限长截断不能留半截记号（实测出过 `765kV≤Un<1000k`）。"""
        out = CR.readable_name("765kV ≤ Un < 1000kV 输电")
        self.assertEqual("765kV输电系统", out)
        self.assertFalse(out.endswith(("k", "V", "<", "≤", ">")))

    def test_comma_merged_cell_takes_main_name(self):
        """OCR 合并单元格残留：`中压主供电系统1,电压等级1` → `中压主供电系统1`。"""
        self.assertEqual("中压主供电系统1", CR.readable_name("中压主供电系统1,电压等级1"))

    def test_empty_is_empty(self):
        self.assertEqual("", CR.readable_name(""))
        self.assertEqual("", CR.readable_name(None))


class DeviceClassTest(unittest.TestCase):
    """规则：设备分类码 T₁T₂ → **上位类目**（按附录B 释义里的「应用实例」概括）。"""

    def test_valve(self):
        self.assertEqual("阀门", CR.device_class("QM"))

    def test_pump(self):
        self.assertEqual("泵", CR.device_class("GP"))

    def test_cabinet(self):
        self.assertEqual("柜／屏／箱", CR.device_class("UC"))

    def test_sensor(self):
        self.assertEqual("传感器／变送器", CR.device_class("BP"))
        self.assertEqual("传感器／变送器", CR.device_class("BT"))

    def test_transformer(self):
        self.assertEqual("变压器", CR.device_class("TA"))

    def test_majority_not_first_match(self):
        """守坑：不能"首次命中即取"——`GP` 的释义里也有「阀」字，会被误判成阀门。"""
        self.assertEqual("泵", CR.device_class("GP"))
        self.assertNotEqual("阀门", CR.device_class("GP"))

    def test_unknown_code_is_blank(self):
        self.assertEqual("", CR.device_class("ZZ"))
        self.assertEqual("", CR.device_class(""))


class DeriveTest(unittest.TestCase):
    """规则：`derive(列规则, 编码)` → 可直接写进那一列的值。"""

    def test_system_from_code(self):
        self.assertEqual("765kV输电系统", CR.derive("system", "01ABC01QA021"))

    def test_device_class_from_code(self):
        self.assertEqual("开关／断路器", CR.derive("device_class", "01ABC01QA021"))
        self.assertEqual("阀门", CR.derive("device_class", "01ABC01QM031"))

    def test_system_also_works_on_system_level_node(self):
        """系统级码（5 位）也能查出系统名——父级行同样要填这一列。"""
        self.assertEqual("765kV输电系统", CR.derive("system", "01ABC"))

    def test_device_class_needs_device_level_code(self):
        """设备分类码在 T₁T₂（第 8~9 位）：码长不够时**留空**，不猜。"""
        self.assertEqual("", CR.derive("device_class", "01ABC"))
        self.assertEqual("", CR.derive("device_class", "01"))

    def test_non_identifier_is_blank(self):
        """设计院自编标签（`PQ1`）不是标识 → 留空。"""
        self.assertEqual("", CR.derive("system", "PQ1"))
        self.assertEqual("", CR.derive("device_class", "PQ1"))
        self.assertEqual("", CR.derive("system", ""))

    def test_unknown_rule_is_blank(self):
        self.assertEqual("", CR.derive("no_such_rule", "01ABC01QA021"))

    def test_value_is_a_short_name_not_a_sentence(self):
        """值必须是**短名**，不能是码表整句（本列存在的意义就是易读）。"""
        for code in ("01ABC01QA021", "01RER01QM031", "01AC01GP011"):
            val = CR.derive("system", code)
            self.assertNotIn(" / ", val)
            self.assertNotIn("从(不包括)", val)
            self.assertLessEqual(len(val), 14)


class ReservedSystemTest(unittest.TestCase):
    """规则：附录A「后续标准预留」段 = 国标**未定义该码** → 不是系统名，留空 + 说明原因。"""

    def test_recognises_reserved_range(self):
        self.assertTrue(CR.is_reserved_system("RER"))
        self.assertFalse(CR.is_reserved_system("ABC"))
        self.assertFalse(CR.is_reserved_system(""))

    def test_derive_returns_blank_for_reserved(self):
        self.assertEqual("", CR.derive("system", "01RER01QM031"))

    def test_note_explains_why_blank(self):
        note = CR.derive_note("system", "01RER01QM031")
        self.assertIn("预留", note)
        self.assertIn("RER", note)
        self.assertIn("需人工确认", note)

    def test_no_note_when_value_available(self):
        self.assertEqual("", CR.derive_note("system", "01ABC01QA021"))


class ColumnRuleRegistryTest(unittest.TestCase):
    """规则：**新增列 = 加一条规则**（注册表决定回填时给哪些列推导）。"""

    def test_registered_columns(self):
        self.assertEqual("system", CR.COLUMN_RULES.get("系统"))
        self.assertEqual("device_class", CR.COLUMN_RULES.get("设备类"))
        self.assertEqual("device_class", CR.COLUMN_RULES.get("设备类型"))

    def test_every_registered_rule_is_implemented(self):
        """守坑：注册了规则名但 `derive` 不认 → 那一列会永远填不出来（静默失效）。"""
        for header, rule in CR.COLUMN_RULES.items():
            with self.subTest(header=header):
                self.assertIsInstance(rule, str)
                self.assertTrue(CR.derive(rule, "01ABC01QA021"),
                                f"列「{header}」注册的规则 {rule} 推不出值")


if __name__ == "__main__":
    unittest.main()
