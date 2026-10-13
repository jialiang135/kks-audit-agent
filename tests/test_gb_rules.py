# -*- coding: utf-8 -*-
"""GB/T 50549-2020 合规规则测试。

全部依据国标原文，不靠反推：
  4.4.2 系统分类码 = 3 个大写字母
  4.4.3 设备分类码 = 2 个大写字母
  4.4.4 系统码 = 系统分类码 + 2 位编号
  4.4.5 设备码 = 设备分类码 + 3 位编号
  注 1  禁用字母 I、O
  表4.3.1-2 全厂码：01~60 机组 / 61~90 两机共用 / 91~99 三机以上 / 00 公用
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

import gb_rules as GB  # noqa: E402


class PlantCodeTest(unittest.TestCase):
    """机组号 → 全厂码（国标表4.3.1-2）。"""

    def test_single_unit(self):
        self.assertEqual("01", GB.unit_label_to_plant_code("1号机组"))
        self.assertEqual("02", GB.unit_label_to_plant_code("2号机组"))
        self.assertEqual("07", GB.unit_label_to_plant_code("7号机组"))
        self.assertEqual("60", GB.unit_label_to_plant_code("60号机组"))

    def test_two_unit_shared(self):
        """61~90 = 两机共用；61 对应 1、2 号机组。

        🔴 这里踩过坑：用 `findall(r"(\\d+)号")` 只会拿到紧邻"号"的那个数字，
        「1、2号机组共用」只解析出 [2]，于是永远匹配不上 61。
        """
        self.assertEqual("61", GB.unit_label_to_plant_code("1、2号机组共用"))
        self.assertEqual("61", GB.unit_label_to_plant_code("1、2号机组公用"))
        self.assertEqual("62", GB.unit_label_to_plant_code("3、4号机组共用"))
        self.assertEqual("90", GB.unit_label_to_plant_code("59、60号机组共用"))

    def test_three_or_more_shared(self):
        self.assertEqual("", GB.unit_label_to_plant_code("1、2、3号机组共用"))
        self.assertEqual("", GB.unit_label_to_plant_code("1、2、3、4号机组共用"))

    def test_plant_wide(self):
        """00 = 全厂公用；**必须写明"全厂"**才判 00。"""
        self.assertEqual("00", GB.unit_label_to_plant_code("全厂公用"))
        self.assertEqual("00", GB.unit_label_to_plant_code("全厂"))

    def test_bare_gongyong_is_ambiguous_never_00(self):
        """🔴 裸「公用」**不能判成 00**（2026-10-10 实测纠正，原实现是"含公用即 00"）。

        实测（热控点表）：`61LAC10`（电动给水泵）**同一系统**下并存两种写法——
        一部分标「公用」、一部分标「1、2号机组共用」；全厂只有 3 行真·全厂公用
        （500kV GIS）用的才是 `00`。可见裸「公用」是本项目里"共用"的简称。
        据它把 `61`（1、2号机共用）改成 `00`（全厂公用）会**改变含义**，
        并把同一系统的设备劈成两个分支（实测 127 行）。
        """
        self.assertEqual("", GB.unit_label_to_plant_code("公用"))
        self.assertEqual("", GB.unit_label_to_plant_code("共用"))
        self.assertTrue(GB.ambiguous_public_label("公用"))
        self.assertTrue(GB.ambiguous_public_label("共用"))
        # 判得出明确码的说法不算歧义
        self.assertFalse(GB.ambiguous_public_label("全厂公用"))
        self.assertFalse(GB.ambiguous_public_label("1、2号机组共用"))
        self.assertFalse(GB.ambiguous_public_label(""))

    def test_unknown_returns_empty_never_guess(self):
        """无法判定的**绝不猜**（宁可留空交人工）。"""
        for label in ("", "待复核", "未注明", "  ", "未知", "1、3号机组共用"):
            self.assertEqual("", GB.unit_label_to_plant_code(label), f"{label!r} 不应给出全厂码")


class ComplianceTest(unittest.TestCase):
    def test_valid_full_device(self):
        res = GB.check_compliance("01HAC13BP101")
        self.assertTrue(res["ok"])
        self.assertEqual("device", res["level"])

    def test_valid_system_code(self):
        res = GB.check_compliance("01HAC13")
        self.assertTrue(res["ok"])
        self.assertEqual("system", res["level"])

    def test_extended_function_plus_product(self):
        """国标 4.5：产品面标识与功能面标识组合使用 → 前 12 位是功能面标识。"""
        res = GB.check_compliance("=01HAL11GP001-MA01")
        self.assertTrue(res["ok"])
        self.assertEqual("combination", res["level"])

    def test_banned_letter_o(self):
        """国标注 1：禁用字母 I、O。"""
        res = GB.check_compliance("01HOC13BP101")
        self.assertFalse(res["ok"])
        self.assertIn("禁用字母", res["reason"])

    def test_banned_letter_i(self):
        res = GB.check_compliance("01HIC13BP101")
        self.assertFalse(res["ok"])
        self.assertIn("禁用字母", res["reason"])

    def test_wrong_system_class_length(self):
        """系统分类码必须是 3 个字母；`01HA13BP101`（只有 2 个）不合规。"""
        self.assertFalse(GB.check_compliance("01HA13BP101")["ok"])

    def test_wrong_device_class_letters(self):
        """设备分类码必须是 2 个字母；数字位置放错就不合规。"""
        self.assertFalse(GB.check_compliance("01HAC13B1101")["ok"])

    def test_too_short_fragment_is_flagged(self):
        self.assertFalse(GB.check_compliance("01BBC01BC")["ok"])   # 9 位
        self.assertFalse(GB.check_compliance("01HFA01CX01F11")["ok"])  # 前 12 位不是设备码


if __name__ == "__main__":
    unittest.main()
