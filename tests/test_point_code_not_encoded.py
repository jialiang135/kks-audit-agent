# -*- coding: utf-8 -*-
"""指令/测点行不编码 —— **端到端**回归（业务口径 2026-10-12）。

## 口径（用户原话）

> 这些都是指令和测点，发出这个指令和信号的才是设备。只有设备才有 kks 编码。

即：`原KKS编码 = 设备码 + 信号码`（`01HLB10GQ001` + `XB21`）的行**不是设备**，
「KKS编码（修编后）」列**留空**，只在「修编说明」里写明它所服务的设备本体码。

## 为什么要写成端到端（跑 `gen_step1.main`）而不是只测 `revise_code`

旧做法（"测点归一"）在 `revise.py` 里把这类行改成设备本体码，本身**有单测风格的自洽性**；
真正出事的是 `revise → 建树 → 回填` 这条**接线**：12 个不同测点全落成同一个 `01HLB10GQ001`，
一张本来无重复的表（输入 58,276 码 / 唯一 58,274）被做成 **3,563 类、18,088 行重复**。
所以这里断言的是**交付文件里的那一列**，而不是引擎的中间返回値。
"""
from __future__ import annotations

import contextlib
import io
import pathlib
import sys
import tempfile
import unittest
from collections import Counter

import openpyxl

PROJECT = pathlib.Path(__file__).resolve().parents[1]
for p in (str(PROJECT), str(PROJECT / "tools"), str(PROJECT / "kks_gen")):
    if p not in sys.path:
        sys.path.insert(0, p)

import gen_step1  # noqa: E402

HEADER = ["序号", "机组号", "系统", "设备类型", "原设备名称",
          "设备名称（按标准修改加机组号）*", "原KKS编码", "KKS编码（修编后）*",
          "父级KKS编码*", "修改意见（审核人填写）"]

# 一张小表，复刻真实提资里"一个设备 + 它的测点/指令"的形态
ROWS = [
    # 原设备名称,              原KKS编码
    ("A一次风机",             "01HLB10GQ001"),      # 设备本体行 → 有码
    ("A一次风机已运行",        "01HLB10GQ001XB21"),  # 测点（信号源 X）
    ("A一次风机已停止",        "01HLB10GQ001XB22"),
    ("启动A一次风机",          "01HLB10GQ001YB21"),  # 指令（信号应用 Y）
    ("停止A一次风机",          "01HLB10GQ001YB22"),
    ("A一次风机DCS控制位",     "01HLB10GQ001XB41"),
]
BODY = "01HLB10GQ001"
POINT_CODES = [c for _n, c in ROWS[1:]]


class PointCodeNotEncodedEndToEndTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls._tmp = tempfile.TemporaryDirectory()
        root = pathlib.Path(cls._tmp.name)
        cls.src = root / "提资.xlsx"
        cls.out_dir = root / "out"
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "sheet1"
        ws.append(["一、表格说明：该表是依据提资内容整理的专业设备汇总表"])
        ws.append(HEADER)
        for i, (name, code) in enumerate(ROWS, start=1):
            ws.append([i, "1号机组", "", "", name, "", code, "", "", ""])
        wb.save(cls.src)
        wb.close()

        argv = sys.argv
        sys.argv = ["gen_step1.py", str(cls.src), str(cls.out_dir), "--no-ai"]
        try:
            with contextlib.redirect_stdout(io.StringIO()):
                gen_step1.main()
        finally:
            sys.argv = argv
        cls.filled = cls.out_dir / f"{cls.src.stem}_已填写.xlsx"
        cls.rows = cls._read(cls.filled)

    @classmethod
    def tearDownClass(cls):
        cls._tmp.cleanup()

    @staticmethod
    def _read(path: pathlib.Path) -> list[dict]:
        """按**输出表头**建列名映射（别写死下标：左端多了「原文件序号」列）。"""
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb[wb.sheetnames[0]]
        head = next(r for r in ws.iter_rows(min_row=1, max_row=6, values_only=True)
                    if r and "原KKS编码" in [str(c or "").strip() for c in r])
        idx = {str(h).strip(): i for i, h in enumerate(head) if str(h or "").strip()}
        out = [{h: (r[i] if i < len(r) else None) for h, i in idx.items()}
               for r in ws.iter_rows(min_row=3, values_only=True)]
        wb.close()
        return [r for r in out if r.get("原KKS编码")]

    def _by_code(self) -> dict[str, dict]:
        return {str(r["原KKS编码"]).strip(): r for r in self.rows}

    def test_output_was_produced(self):
        self.assertTrue(self.filled.exists(), "端到端跑通、交付文件已生成")
        self.assertEqual(len(self.rows), len(ROWS), "原表每一行都要在")

    def test_device_row_keeps_its_code(self):
        """设备本体行**必须**有码 —— 不能把口径执行成"整列都不填"。"""
        r = self._by_code()[BODY]
        self.assertEqual(BODY, str(r["KKS编码（修编后）*"] or "").strip())

    def test_point_and_command_rows_have_no_code(self):
        """指令/测点行的「KKS编码（修编后）」一律**留空**。"""
        by = self._by_code()
        for code in POINT_CODES:
            with self.subTest(code=code):
                got = by[code]["KKS编码（修编后）*"]
                self.assertIn(got, (None, ""), f"{code} 是指令/测点，不该拿到编码")

    def test_no_duplicate_codes_in_deliverable(self):
        """🔴 回归本体：交付列里**不得出现重复值**。

        旧口径下这 6 行会给出 6 个 `01HLB10GQ001`（5 行重复）。
        """
        codes = [str(r["KKS编码（修编后）*"]).strip() for r in self.rows
                 if str(r["KKS编码（修编后）*"] or "").strip()]
        dup = {c: n for c, n in Counter(codes).items() if n > 1}
        self.assertEqual({}, dup, f"交付编码列出现重复值：{dup}")

    def test_note_names_the_owning_device(self):
        """留空不能留成"没查出东西"——说明里必须写清归属的设备本体码。"""
        by = self._by_code()
        for code in POINT_CODES:
            with self.subTest(code=code):
                note = str(by[code].get("修编说明") or "")
                self.assertIn("指令/测点", note)
                self.assertIn(BODY, note, "说明里要给出所服务的设备本体码")

    def test_parent_column_still_filled(self):
        """留空**只针对编码列**：这些行的「父级KKS编码」照旧要有值。"""
        by = self._by_code()
        for code in POINT_CODES:
            with self.subTest(code=code):
                self.assertTrue(str(by[code]["父级KKS编码*"] or "").strip(),
                                "指令/测点行仍要挂在设备/系统下，父级列不能空")


if __name__ == "__main__":
    unittest.main()
