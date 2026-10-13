# -*- coding: utf-8 -*-
"""AI 名称层回归测试（不打真实接口，只测校验逻辑与 payload 组装）。"""
from __future__ import annotations

import pathlib
import sys
import unittest

PROJECT = pathlib.Path(__file__).resolve().parents[1]
if str(PROJECT) not in sys.path:
    sys.path.insert(0, str(PROJECT))
if str(PROJECT / "kks_gen") not in sys.path:
    sys.path.insert(0, str(PROJECT / "kks_gen"))

import ai_names as AN  # noqa: E402


def _row(**kw) -> dict:
    base = {"row": "1", "kks": "07AAB02UH001", "original_name": "", "parent_name": "",
            "parent_kks": "", "profession": "", "drawing": ""}
    base.update(kw)
    return base


class NameValidationTest(unittest.TestCase):
    """本地复核：生成的名称必须自洽，否则退回人工。"""

    def test_accepts_normal_rewrite(self):
        row = _row(original_name="7号机组NCS测控屏")
        ok, _ = AN._validate_name({"name": "7号机NCS测控屏"}, row)
        self.assertTrue(ok, "机组号缩写是合法改写")

    def test_accepts_short_name_with_parent_prefix(self):
        """短名称补父级前缀会变长，这是设计意图，不能拦。"""
        row = _row(original_name="母线PT", parent_name="10kV工作7A段")
        ok, _ = AN._validate_name({"name": "10kV工作7A段 母线PT"}, row)
        self.assertTrue(ok)

    def test_rejects_wrong_unit_number(self):
        """07 前缀的行不许出现「1号机」。"""
        row = _row(kks="07AAB02UH001", original_name="7号机组NCS测控屏")
        ok, reason = AN._validate_name({"name": "1号机NCS测控屏"}, row)
        self.assertFalse(ok)
        self.assertIn("机组号", reason)

    def test_allows_equipment_number_one(self):
        """'1号主变' 里的 1号 是设备编号，必须允许。"""
        row = _row(kks="07ABG01BA001", original_name="1号主变 高压侧PT")
        ok, reason = AN._validate_name({"name": "1号主变 高压侧PT"}, row)
        self.assertTrue(ok, reason)

    def test_rejects_empty_and_short(self):
        row = _row(original_name="轴流风机")
        self.assertFalse(AN._validate_name({"name": ""}, row)[0])
        self.assertFalse(AN._validate_name({"name": "风机"}, row)[0])

    def test_rejects_control_chars(self):
        row = _row(original_name="轴流风机")
        ok, reason = AN._validate_name({"name": "轴流\n风机"}, row)
        self.assertFalse(ok)

    def test_rejects_totally_unrelated_name(self):
        row = _row(original_name="凝结水泵变频器切换柜")
        ok, reason = AN._validate_name({"name": "汽机房照明配电箱"}, row)
        self.assertFalse(ok)
        self.assertIn("偏离", reason)

    def test_rejects_much_shorter_loss_of_info(self):
        row = _row(original_name="10kV工作7A段工作电源进线开关")
        ok, reason = AN._validate_name({"name": "开关"}, row)
        self.assertFalse(ok)


class PayloadTest(unittest.TestCase):
    """请求体：必须带上模型判断所需的全部证据。"""

    def test_payload_carries_evidence(self):
        rows = [{
            "row": "10", "kks": "07ABG01BA001A", "original_name": "",
            "parent_name": "500kV 1号主变进线间隔系统", "parent_kks": "07ABG01",
            "profession": "电气一次", "drawing": "F1284S-D0104-02",
            "device_meaning": "电压互感器", "phase_suffix": "A相",
            "same_class_names": ["1号主变 高压侧PT"],
            "sibling_names": ["1号主变 高压侧PT B相"],
        }]
        payload = AN._build_payload(rows)
        item = payload["待改写的行"][0]
        for key in ("原名称", "上级设备名称", "KKS编码", "设备分类码含义", "分相后缀",
                    "同类设备已有名称", "同父级兄弟名称"):
            self.assertIn(key, item, f"证据 {key} 缺失会让模型判断失据")
        self.assertEqual("A相", item["分相后缀"])

    def test_fewshot_is_from_real_evidence(self):
        """few-shot 必须是真实成果，不是编的——否则模型学不到本厂写法。"""
        self.assertGreaterEqual(len(AN.FEWSHOT), 6)
        for old, parent, code, name in AN.FEWSHOT:
            self.assertTrue(code.startswith(("07", "08")))
            self.assertTrue(name)


class DisabledWithoutConfigTest(unittest.TestCase):
    def test_disabled_when_ai_off(self):
        import os
        old = os.environ.get("AI_ENABLED")
        os.environ["AI_ENABLED"] = "false"
        try:
            result = AN.resolve_names([_row(original_name="轴流风机")])
        finally:
            if old is None:
                os.environ.pop("AI_ENABLED", None)
            else:
                os.environ["AI_ENABLED"] = old
        self.assertEqual("disabled", result["status"])
        self.assertIn("AI 未启用", result["reason"])
        self.assertEqual({}, result["items"])


if __name__ == "__main__":
    unittest.main()
