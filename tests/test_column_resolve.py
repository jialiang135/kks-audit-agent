# -*- coding: utf-8 -*-
"""列定位的鲁棒性回归测试。

背景（实测事故）：原来用**固定列名**（"原名称"/"设计编码"）读输入，
换个表头叫法就整个流程失效 —— 把「设计编码」写成「设计码」、「原名称」写成「图纸名称」时，
生成器把这两列当空，跑出来的东西全是空的。

各厂/各设计院表头写法不统一，所以必须按语义模糊匹配。本测试锁死这个能力。
"""
from __future__ import annotations

import pathlib
import sys
import unittest

PROJECT = pathlib.Path(__file__).resolve().parents[1]
if str(PROJECT) not in sys.path:
    sys.path.insert(0, str(PROJECT))
if str(PROJECT / "tools") not in sys.path:
    sys.path.insert(0, str(PROJECT / "tools"))

from gen_step1 import resolve_columns  # noqa: E402


class ColumnResolveTest(unittest.TestCase):
    def test_standard_header(self):
        header = ["序号", "原名称", "设计编码", "设备名称", "KKS编码",
                  "上级设备名称", "上级KKS编码", "专业", "图纸名称或图号"]
        col = resolve_columns(header)
        self.assertEqual(1, col["original_name"])
        self.assertEqual(2, col["design_code"])
        self.assertEqual(4, col["kks"])
        self.assertEqual(3, col["device_name"])
        self.assertEqual(6, col["parent_kks"])
        self.assertEqual(5, col["parent_name"])
        self.assertEqual(7, col["profession"])
        self.assertEqual(8, col["drawing"])

    def test_renamed_columns_still_resolve(self):
        """列名换成「设计码 / 图纸名称 / 设备名 / KKS码(最终) / 父级编码」也要认得。"""
        header = ["序号", "图纸名称", "设计码", "设备名", "KKS码(最终)",
                  "父级名称", "父级编码", "专业", "图纸名称或图号"]
        col = resolve_columns(header)
        self.assertEqual(1, col["original_name"], "图纸名称 应识别为原名称")
        self.assertEqual(2, col["design_code"], "设计码 应识别为设计编码")
        self.assertEqual(3, col["device_name"])
        self.assertEqual(4, col["kks"])
        self.assertEqual(5, col["parent_name"])
        self.assertEqual(6, col["parent_kks"])

    def test_shuffled_order_still_resolves(self):
        header = ["序号", "原名称", "设计编码", "上级设备名称", "上级KKS编码",
                  "设备名称", "KKS编码", "专业"]
        col = resolve_columns(header)
        self.assertEqual(5, col["device_name"])
        self.assertEqual(6, col["kks"])
        self.assertEqual(3, col["parent_name"])
        self.assertEqual(4, col["parent_kks"])

    def test_design_code_not_confused_with_kks(self):
        """「设计编码」与「KKS编码」都在时不能串。"""
        header = ["原名称", "设计编码", "KKS编码", "设备名称"]
        col = resolve_columns(header)
        self.assertEqual(1, col["design_code"])
        self.assertEqual(2, col["kks"])

    def test_parent_kks_not_confused_with_kks(self):
        header = ["KKS编码", "上级KKS编码", "设计编码", "原名称", "设备名称"]
        col = resolve_columns(header)
        self.assertEqual(0, col["kks"])
        self.assertEqual(1, col["parent_kks"])

    def test_original_name_not_confused_with_device_name(self):
        header = ["设备名称", "原名称", "设计编码", "KKS编码"]
        col = resolve_columns(header)
        self.assertEqual(0, col["device_name"])
        self.assertEqual(1, col["original_name"])

    def test_missing_key_column_returns_none(self):
        """关键列真的没有时返回 None（调用方据此报错，而不是静默生成空数据）。"""
        col = resolve_columns(["序号", "名称", "编码"])
        self.assertIsNone(col["design_code"])
        self.assertIsNone(col["original_name"])

    def test_client_template_with_original_and_revised_pairs(self):
        """业主模板：原始列与修编列成对出现（舟山暖通两版真实表头）。

        实测坑：`设备名称（原设计图内）` 里的 `设计图内` 同时也是
        `原KKS编码（设计图内）` 的组成部分——用 `原设计图` 当命中词才能只落在名称列。
        """
        header = ["序号", "机组号", "系统", "设备类型", "设备名称（原设计图内）*",
                  "设备名称（修改后）", "原KKS编码（设计图内）", "KKS编码（修编后）*",
                  "父级KKS编码*", "修改意见（审核人填写）", "图纸号", "备注"]
        col = resolve_columns(header)
        self.assertEqual(4, col["original_name"], "应取「原设计图内」的名称列")
        self.assertEqual(6, col["design_code"], "应取「设计图内」的原始编码列")
        self.assertEqual(5, col["device_name"], "应取「修改后」的名称列")
        self.assertEqual(7, col["kks"], "应取「修编后」的编码列")
        self.assertEqual(8, col["parent_kks"])
        self.assertEqual(3, col["device_type"])
        self.assertEqual(2, col["system_name"])

    def test_header_row_detected_for_client_template(self):
        """真实文件：第 1 行是"表格说明"长文本，表头在第 2 行——必须能自动探测到。"""
        from gen_step1 import detect_header_row
        rows = [
            ["一、表格说明：该表是依据提资内容，按《中国中煤能源集团电厂标识系统编码标准》编制"],
            ["序号", "机组号", "系统", "设备类型", "设备名称（原设计图内）*", "设备名称（修改后）",
             "原KKS编码（设计图内）", "KKS编码（修编后）*", "父级KKS编码*", "备注"],
            ["1", "全厂公用", None, None, "全厂公用", None, "00", None, "-1", None],
        ]
        self.assertEqual(1, detect_header_row(rows))


if __name__ == "__main__":
    unittest.main()
