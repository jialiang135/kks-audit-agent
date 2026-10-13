# -*- coding: utf-8 -*-
"""真实文件结构的读取回归测试。

来源：用户给的真实汇总表 —— 第 1 行是"表格说明"长文本（合并单元格），
表头在第 2 行，列名叫「原KKS编码」「原设备名称」而不是「设计编码」「原名称」。
修复前：生成器死在第 1 行当表头 → "表头识别失败，缺少关键列"。
"""
from __future__ import annotations

import pathlib
import sys
import unittest

import openpyxl

PROJECT = pathlib.Path(__file__).resolve().parents[1]
if str(PROJECT) not in sys.path:
    sys.path.insert(0, str(PROJECT))
if str(PROJECT / "tools") not in sys.path:
    sys.path.insert(0, str(PROJECT / "tools"))

import gen_step1  # noqa: E402
import tempfile  # noqa: E402

DESC_ROW = ("一、表格说明：该表是依据提资内容，对照清册、清单、布置图、安装图、系统图"
            "收集到的专业设备汇总表，表中红色字体为清册、清单内未列出但对照系统图新增加的设备")
REAL_HEADER = ["序号", "机组号", "系统", "设备类型", "原设备名称",
               "设备名称（按标准修改加机组号）*", "原KKS编码", "KKS编码（修编后）*",
               "上级KKS编码*", "专业", "图纸名称或图号"]
SUB_HEADER = [None, None, None, None, None,
              "按照《中国中煤能源集团电厂标识系统编码应用细则》修改", "专业人员提供资料内设备KKS编码",
              "按照《中国中煤能源集团电厂标识系统编码应用细则》修改", "系统和设备上级归属的KKS编码"]


def _write(path: pathlib.Path) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "sheet1"
    ws.append([DESC_ROW])
    ws.append(REAL_HEADER)
    ws.append(SUB_HEADER)
    ws.append([1, "1号机组", "锅炉", "压力变送器", "省煤器出口压力1",
               "", "01HAC13BP101", "", "", "热控", "K0104-01"])
    ws.append([2, "2号机组", "锅炉", "压力变送器", "省煤器出口压力1",
               "", "02HAC13BP101", "", "", "热控", "K0104-01"])
    ws.append([3, "1号机组", "锅炉", "压力变送器", "省煤器出口压力2",
               "", "01HAC13BP102", "", "", "热控", "K0104-01"])
    wb.save(path)


class HeaderDetectionTest(unittest.TestCase):
    """表头不在第 1 行时也要能找到。"""

    @classmethod
    def setUpClass(cls):
        cls._tmp = tempfile.TemporaryDirectory()
        cls.xlsx = pathlib.Path(cls._tmp.name) / "real.xlsx"
        _write(cls.xlsx)

    @classmethod
    def tearDownClass(cls):
        cls._tmp.cleanup()

    def test_detects_header_below_description_row(self):
        rows = [
            [DESC_ROW],
            REAL_HEADER,
            SUB_HEADER,
        ]
        self.assertEqual(1, gen_step1.detect_header_row(rows), "表头应定位到第 2 行")

    def test_detects_standard_header_at_first_row(self):
        rows = [["序号", "原名称", "设计编码", "设备名称", "KKS编码"], ["1", "x", "07AAB02UH001", "y", "07AAB02UH001"]]
        self.assertEqual(0, gen_step1.detect_header_row(rows))

    def test_returns_none_when_no_header(self):
        rows = [[DESC_ROW], ["随便", "写点", "东西"]]
        self.assertIsNone(gen_step1.detect_header_row(rows))

    def test_loads_real_layout(self):
        seeds, _layout = gen_step1.load(self.xlsx)
        self.assertEqual(3, len(seeds), "应读出 3 条数据（说明行不算）")
        codes = sorted(s["design_code"] for s in seeds)
        self.assertEqual(["01HAC13BP101", "01HAC13BP102", "02HAC13BP101"], codes)
        self.assertEqual("省煤器出口压力1", seeds[0]["original_name"])

    def test_original_kks_column_used_as_design_code(self):
        """列名是「原KKS编码」时也要认得（不能用「原」字把它排除掉）。"""
        col = gen_step1.resolve_columns(REAL_HEADER)
        self.assertEqual(REAL_HEADER.index("原KKS编码"), col["design_code"])
        self.assertEqual(REAL_HEADER.index("原设备名称"), col["original_name"])
        self.assertEqual(REAL_HEADER.index("KKS编码（修编后）*"), col["kks"])
        self.assertEqual(REAL_HEADER.index("上级KKS编码*"), col["parent_kks"])


class DesignCodeCleaningTest(unittest.TestCase):
    """设计编码清洗。"""

    def test_strips_inner_spaces(self):
        """真实表格里编码被空格断开：`07GQJ40 BT301` → `07GQJ40BT301`。"""
        self.assertEqual("07GQJ40BT301", gen_step1.normalize_design_code("07GQJ40 BT301"))

    def test_rejects_chinese_text(self):
        """编码列里塞了名称/说明的，判为无效而不是硬当编码。"""
        self.assertEqual("", gen_step1.normalize_design_code("7号炉辅汽至锅炉71-54m区域产品伴热用汽管"))
        self.assertEqual("", gen_step1.normalize_design_code("专业人员提供资料内设备KKS编码"))

    def test_fullwidth_normalized(self):
        self.assertEqual("07AAB02UH001", gen_step1.normalize_design_code("０７ＡＡＢ０２ＵＨ００１"))

    def test_keeps_dot_suffix(self):
        self.assertEqual("07BFA01UC001.A1", gen_step1.normalize_design_code("07BFA01UC001.A1"))

    def test_empty(self):
        self.assertEqual("", gen_step1.normalize_design_code(""))
        self.assertEqual("", gen_step1.normalize_design_code("   "))


if __name__ == "__main__":
    unittest.main()
