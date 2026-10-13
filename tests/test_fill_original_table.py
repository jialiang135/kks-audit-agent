# -*- coding: utf-8 -*-
"""回填原表功能的回归测试（**契约已于 2026-10-10 按业务确认变更**）。

用户诉求：**直接在他原表的列里填**，且**不改变原文件行序**——
原表所有行（**保持原文件行序**）+ 生成的父级行（插在「它的第一个子行」上方），
第一列新增「原文件序号」= 该行在原文件的 Excel 实际行号（新增行留空），
于是拿输出能回原文件逐行人工核对。

新版约束：
  · 不动源文件；保留原表全部列（整体右移一列给新的序号列让位）
  · **只填原本为空的格**（业务铁律 2026-10-10）：原表有数据的格绝不回填
  · **引擎没有结论的行不覆盖**（待人工/未判定）——不能把人家已有的值清掉
  · 生成的父级行就近插入；**不生成部件行**（实测精确率 0%）
  · 父级名按**国标码表语义 + 机组号**（2/3/5 位），7/12 位用子节点公共前缀
"""
from __future__ import annotations

import pathlib
import sys
import tempfile
import unittest

import openpyxl

PROJECT = pathlib.Path(__file__).resolve().parents[1]
for p in (str(PROJECT), str(PROJECT / "tools"), str(PROJECT / "kks_gen")):
    if p not in sys.path:
        sys.path.insert(0, p)

import gen_step1  # noqa: E402
from tree_rules import build_tree  # noqa: E402

HEADER = ["序号", "机组号", "原设备名称", "设备名称（按标准修改加机组号）*",
          "原KKS编码", "KKS编码（修编后）*", "父级KKS编码*", "修改意见（审核人填写）"]
IDX = {h: i for i, h in enumerate(HEADER)}


class FillOriginalTableTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls._tmp = tempfile.TemporaryDirectory()
        root = pathlib.Path(cls._tmp.name)
        cls.src = root / "user.xlsx"
        cls.out = root / "user_已填写.xlsx"
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "sheet1"
        ws.append(["一、表格说明：该表是依据提资内容整理的专业设备汇总表，红色字体为新增设备"])
        ws.append(HEADER)
        ws.append([1, "1号机组", "省煤器出口压力变送器1", "", "01HAC13BP101", "", "", ""])
        ws.append([2, "2号机组", "省煤器出口压力变送器1", "", "02HAC13BP101", "", "", ""])
        # 第 3 行：**设备名称列预填了人工内容** —— 用来验证"原表有数据绝不回填"；
        # 修改意见列也预填，验证其它列原样保留
        ws.append([3, "1号机组", "省煤器出口压力变送器2", "人工填过的名称", "01HAC13BP102", "", "", "已核对"])
        wb.save(cls.src)
        cls.original_bytes = cls.src.read_bytes()

        cls.seeds, cls.layout = gen_step1.load(cls.src)
        built = build_tree(cls.seeds)
        cls.nodes = built["nodes"]
        cls.stat = gen_step1.fill_original_table(
            cls.src, cls.out, cls.nodes, cls.layout,
            ordered=gen_step1._tree_order(cls.nodes), seeds=cls.seeds)

    @classmethod
    def tearDownClass(cls):
        cls._tmp.cleanup()

    def _rows(self):
        """把输出读成 [dict(列名→值)]。

        ⚠️ 按**输出的表头行**建列名映射，不用固定下标：2026-10-10 新增「原文件序号」列后
           整表右移一列，写死下标会让全部断言假失败（踩过）。
        """
        wb = openpyxl.load_workbook(self.out, data_only=True)
        ws = wb[wb.sheetnames[0]]
        head = next(r for r in ws.iter_rows(min_row=1, max_row=5, values_only=True)
                    if r and "序号" in [str(c or "").strip() for c in r])
        idx = {str(h).strip(): i for i, h in enumerate(head) if str(h or "").strip()}
        out = []
        for r in ws.iter_rows(min_row=3, values_only=True):
            out.append({h: (r[i] if i < len(r) else None) for h, i in idx.items()})
        wb.close()
        return out

    # ---------- 基本约束 ----------
    def test_source_file_untouched(self):
        self.assertEqual(self.original_bytes, self.src.read_bytes(), "源文件不能被改动")

    def test_preserves_other_columns(self):
        """用户的其它列（机组号 / 修改意见）必须原样保留。"""
        by_code = {r["原KKS编码"]: r for r in self._rows() if r.get("原KKS编码")}
        self.assertEqual("1号机组", by_code["01HAC13BP102"]["机组号"])
        self.assertEqual("已核对", by_code["01HAC13BP102"]["修改意见（审核人填写）"])

    def test_description_row_not_touched(self):
        """第 1 行的表格说明不能被当数据回填。"""
        wb = openpyxl.load_workbook(self.out, data_only=True)
        ws = wb[wb.sheetnames[0]]
        self.assertIsNone(ws.cell(row=1, column=IDX["KKS编码（修编后）*"] + 1).value, "说明行不该被回填")
        wb.close()

    # ---------- 目标列写入 / 覆盖 ----------
    def test_writes_target_columns(self):
        by_code = {r["原KKS编码"]: r for r in self._rows() if r.get("原KKS编码")}
        r = by_code["01HAC13BP101"]
        self.assertEqual("01HAC13BP101", r["KKS编码（修编后）*"])
        self.assertEqual("01HAC13", r["父级KKS编码*"])
        self.assertTrue(r["设备名称（按标准修改加机组号）*"])

    def test_does_not_overwrite_existing_cells(self):
        """🔴 业务铁律：**原表一开始有数据的格，绝对不回填**（引擎只填空的）。

        （2026-10-10 确认。此前一度按"覆盖"实现，在竣工版上覆盖了 24.7 万格 —— 已撤销。）
        """
        by_code = {r["原KKS编码"]: r for r in self._rows() if r.get("原KKS编码")}
        self.assertEqual("人工填过的名称", by_code["01HAC13BP102"]["设备名称（按标准修改加机组号）*"],
                         "原表已填的设备名称必须原样保留，不能被引擎结果覆盖")
        self.assertGreaterEqual(self.stat.get("cells_kept", 0), 1, "应统计到保留原有内容的格数")

    # ---------- 生成的父级行 + 树序 ----------
    def test_original_row_order_kept_and_parents_inserted_before_first_child(self):
        """🔴 行序口径（2026-10-10 用户定调）：**原表行序不变**，新增父级行插在它第一个子行上方。

        夹具里三行的原表顺序是 `01HAC13BP101` → `02HAC13BP101` → `01HAC13BP102`，
        所以输出里 `01HAC13BP102` 必须**在最后**（按树序排会把它夹到 02 组的后面）。
        """
        codes = [r.get("KKS编码（修编后）*") for r in self._rows()]
        expect = ["01", "01H", "01HAC", "01HAC13", "01HAC13BP101",
                  "02", "02H", "02HAC", "02HAC13", "02HAC13BP101",
                  "01HAC13BP102"]
        self.assertEqual(expect, codes)

    def test_original_rows_keep_excel_row_numbers_in_first_column(self):
        """新增第一列「原文件序号」= 该行在原文件的 **Excel 实际行号**；新增父级行留空。

        口径："有号 = 原文件来的行，空号 = 新增行"，于是输出任何一行都能回原文件定位。
        """
        rows = self._rows()
        seq = {r.get("KKS编码（修编后）*"): r.get("原文件序号") for r in rows}
        self.assertEqual(3, seq.get("01HAC13BP101"), "夹具第 1 个数据行在 Excel 第 3 行")
        self.assertEqual(4, seq.get("02HAC13BP101"))
        self.assertEqual(5, seq.get("01HAC13BP102"))
        for code in ("01", "01H", "01HAC", "01HAC13", "02", "02HAC13"):   # 生成行
            self.assertIsNone(seq.get(code), f"{code} 是新增父级行，序号格应留空")

    def test_header_gets_the_new_column_name(self):
        """表头行第一列写列名「原文件序号」（不是行号）。"""
        wb = openpyxl.load_workbook(self.out, data_only=True)
        ws = wb[wb.sheetnames[0]]
        self.assertEqual("原文件序号", ws.cell(row=2, column=1).value)
        self.assertEqual("序号", ws.cell(row=2, column=2).value, "原「序号」列右移一列")
        wb.close()

    def test_all_source_columns_shifted_right_by_one(self):
        """原表 8 列整体右移一列，内容与样式（含红色字体）都跟着走。"""
        wb = openpyxl.load_workbook(self.out, data_only=True)
        ws = wb[wb.sheetnames[0]]
        head = [str(ws.cell(row=2, column=c).value or "") for c in range(1, 11)]
        self.assertEqual("原文件序号", head[0])
        self.assertEqual(list(HEADER), head[1:1 + len(HEADER)], "原表列顺序不变，只是整体右移")
        wb.close()

    def test_generated_rows_have_no_source_columns(self):
        """生成的父级行没有原名称/原设计编码（它们不是摘录来的）。"""
        r = {x.get("KKS编码（修编后）*"): x for x in self._rows()}["01HAC13"]
        self.assertIsNone(r["原KKS编码"])
        self.assertIsNone(r["原设备名称"])
        self.assertEqual("01HAC", r["父级KKS编码*"])

    def test_parent_names_follow_code_semantics_and_unit(self):
        """父级名 = 机组号 + 附录A 码表名（2/3/5 位）。"""
        r = {x.get("KKS编码（修编后）*"): x for x in self._rows()}
        nm = lambda c: r[c]["设备名称（按标准修改加机组号）*"]
        self.assertEqual("1号机组", nm("01"))                # 表4.3.1-2
        self.assertEqual("2号机组", nm("02"))
        self.assertEqual("1号机热产生", nm("01H"))            # 附录A「H」→ 热产生（取第一分句）
        self.assertEqual("1号机省煤器系统", nm("01HAC"))      # 附录A「HAC」→ 省煤器系统
        self.assertEqual("2号机省煤器系统", nm("02HAC"))
        # 生成行数统计应与"没有原KKS编码的行"数一致
        self.assertEqual(self.stat.get("generated_rows", 0),
                         len([x for x in self._rows() if not x.get("原KKS编码")]))

    def test_no_component_rows(self):
        """不生成部件行（KKS 里不应出现 `-`）。"""
        for x in self._rows():
            code = str(x.get("KKS编码（修编后）*") or "")
            self.assertNotIn("-", code, f"不应生成部件行：{code}")


class FinalCodeFallbackTest(unittest.TestCase):
    """**只有最终编码、没有设计编码**的行：用已有编码当种子参与建树（只审查，绝不改写）。

    实测背景（鄂州四期竣工版）：113,413 行里 47,984 行属于这一类——只有 KKS 编码和设备名称。
    以前这些行整行"未判定"，上级编码/系统/设备类全填不出来；
    但它们的信息是现成的，不利用就是白白浪费一整张表。
    """

    @classmethod
    def setUpClass(cls):
        cls._tmp = tempfile.TemporaryDirectory()
        root = pathlib.Path(cls._tmp.name)
        cls.src = root / "final_only.xlsx"
        cls.out = root / "final_only_已填写.xlsx"
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["一、表格说明：该表是依据提资内容整理的专业设备汇总表"])
        ws.append(HEADER)
        # 第 1 行：设计编码空、KKS 编码有值 → 走兜底种子；设备名称预填（铁律：不动）
        ws.append([1, "1号机组", "", "人工压力变送器名称1", "", "01HAC13BP101", "", ""])
        # 第 2 行：同上，父级编码列也预填 → 验证也不回填
        ws.append([2, "1号机组", "", "人工压力变送器名称2", "", "01HAC13BP102", "预填父级", ""])
        # 第 3 行：设计编码填的是名称串（无效），但最终编码有效 → 仍要兜底用最终编码
        ws.append([3, "1号机组", "", "人工压力变送器名称3", "见附图", "01HAC13BP103", "", ""])
        # 第 4 行：编码列全空 → 没有任何可用信息，跳过（不生成）
        ws.append([4, "1号机组", "", "", "", "", "", ""])
        # 第 5 行：设计编码是名称串、最终编码也没有 → 跳过
        ws.append([5, "1号机组", "", "", "见附图一", "", "", ""])
        wb.save(cls.src)

        cls.seeds, cls.layout = gen_step1.load(cls.src)
        cls.stat = gen_step1.fill_original_table(
            cls.src, cls.out, build_tree(cls.seeds)["nodes"], cls.layout,
            ordered=gen_step1._tree_order(build_tree(cls.seeds)["nodes"]), seeds=cls.seeds)

    @classmethod
    def tearDownClass(cls):
        cls._tmp.cleanup()

    def _rows(self):
        """按输出表头行建列名映射（新增「原文件序号」列后整表右移，写死下标会假失败）。"""
        wb = openpyxl.load_workbook(self.out, data_only=True)
        ws = wb[wb.sheetnames[0]]
        head = next(r for r in ws.iter_rows(min_row=1, max_row=5, values_only=True)
                    if r and "序号" in [str(c or "").strip() for c in r])
        idx = {str(h).strip(): i for i, h in enumerate(head) if str(h or "").strip()}
        out = [{h: (r[i] if i < len(r) else None) for h, i in idx.items()}
               for r in ws.iter_rows(min_row=3, values_only=True)]
        wb.close()
        return out

    def test_final_code_column_used_as_seed(self):
        by_code = {r["KKS编码（修编后）*"]: r for r in self._rows()}
        self.assertIn("01HAC13BP101", by_code, "已有最终编码的行要参与建树，不能整行未判定")
        self.assertEqual("01HAC13", by_code["01HAC13BP101"]["父级KKS编码*"],
                         "空着的上级编码要能推出来")

    def test_rows_without_any_code_are_skipped(self):
        """没有任何可用编码的行 → 不生成、不判定（保持原样）。"""
        self.assertEqual(3, len(self.seeds), "应只认下 3 个有编码来源的行")

    def test_invalid_design_code_still_falls_back_to_final_code(self):
        """设计编码列填的是名称串（无效）时，仍要用最终编码兜底，别整行丢掉。"""
        row = next(s for s in self.seeds if s["design_code"] == "01HAC13BP103")
        self.assertTrue(row["code_is_final"])

    def test_final_code_never_rewritten(self):
        """🔴 原表已有的最终编码**一字不改**（用户铁律）。"""
        by_code = {r["KKS编码（修编后）*"]: r for r in self._rows()}
        self.assertEqual("01HAC13BP102", by_code["01HAC13BP102"]["KKS编码（修编后）*"])

    def test_preexisting_parent_code_kept(self):
        """原表已填的上级编码也不回填。"""
        by_code = {r["KKS编码（修编后）*"]: r for r in self._rows()}
        self.assertEqual("预填父级", by_code["01HAC13BP102"]["父级KKS编码*"])

    def test_name_source_falls_back_to_device_name_column(self):
        """「原设备名称」为空时，用已有的「设备名称」当名称来源（别再猜一遍）。"""
        row = next(s for s in self.seeds if s["design_code"] == "01HAC13BP101")
        self.assertEqual("人工压力变送器名称1", row["original_name"])
        self.assertTrue(row["code_is_final"])

    def test_no_status_echo_notes(self):
        """「已符合国标」这类状态回显不该进「修编说明」——没要改的事就不写说明。"""
        for r in self._rows():
            note = str(r.get("修改意见（审核人填写）") or "")
            self.assertNotIn("已符合国标", note)
        # 说明列本身：不能出现纯状态回显
        wb = openpyxl.load_workbook(self.out, data_only=True)
        ws = wb[wb.sheetnames[0]]
        note_col = ws.max_column
        for row in ws.iter_rows(min_row=3, values_only=True):
            v = str(row[note_col - 1] or "")
            self.assertNotIn("已符合国标", v)
        wb.close()

    def test_source_file_untouched(self):
        wb = openpyxl.load_workbook(self.src, data_only=True)
        ws = wb[wb.sheetnames[0]]
        self.assertEqual("01HAC13BP101", ws.cell(row=3, column=IDX["KKS编码（修编后）*"] + 1).value)
        wb.close()



class LayoutRemapTest(unittest.TestCase):
    """左插一列后，**版式**必须一起右移（列宽 / 合并区 / 自动筛选 / 冻结窗格）。

    为什么单独立一个夹具：主夹具没有合并单元格、没有列宽、没有筛选，测不到这段。
    而这三样错位（列宽跑到隔壁、合并区压住新列、筛选少一列）都是"打开文件才看得出来"的坑。
    """

    @classmethod
    def setUpClass(cls):
        cls._tmp = tempfile.TemporaryDirectory()
        root = pathlib.Path(cls._tmp.name)
        cls.src = root / "layout.xlsx"
        cls.out = root / "layout_已填写.xlsx"
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["表格说明：本表用于验证版式右移"])          # 第 1 行
        ws.merge_cells("A1:H1")                              # 合并区（原表就有）
        ws.append(HEADER)                                    # 第 2 行 = 表头
        ws.append([1, "1号机组", "省煤器出口压力变送器1", "", "01HAC13BP101", "", "", ""])
        ws.column_dimensions["A"].width = 6.5
        ws.column_dimensions["B"].width = 11.0
        ws.column_dimensions["C"].width = 30.0
        ws.auto_filter.ref = "A2:H3"
        ws.freeze_panes = "B2"                               # 冻结 1 列 + 1 行
        wb.save(cls.src)

        cls.seeds, cls.layout = gen_step1.load(cls.src)
        built = build_tree(cls.seeds)
        gen_step1.fill_original_table(cls.src, cls.out, built["nodes"], cls.layout,
                                      ordered=gen_step1._tree_order(built["nodes"]),
                                      seeds=cls.seeds)
        cls.ws = openpyxl.load_workbook(cls.out)
        cls.sheet = cls.ws[cls.ws.sheetnames[0]]

    @classmethod
    def tearDownClass(cls):
        cls.ws.close()
        cls._tmp.cleanup()

    def test_merged_range_shifted(self):
        """合并区必须跟着右移，否则原来的 A1:H1 会盖住新的「原文件序号」列。"""
        ranges = {str(r) for r in self.sheet.merged_cells.ranges}
        self.assertIn("B1:I1", ranges)
        self.assertNotIn("A1:H1", ranges)

    def test_column_widths_shifted(self):
        """列宽跟着右移：原 A 的 6.5 应落到 B，A 变成新序号列（默认宽度）。"""
        self.assertAlmostEqual(6.5, self.sheet.column_dimensions["B"].width, places=1)
        self.assertAlmostEqual(11.0, self.sheet.column_dimensions["C"].width, places=1)
        self.assertAlmostEqual(30.0, self.sheet.column_dimensions["D"].width, places=1)

    def test_autofilter_covers_new_column(self):
        """筛选范围右移，且**左边界拉到 A** —— 这样能直接按「原文件序号」筛行。"""
        self.assertTrue(str(self.sheet.auto_filter.ref).startswith("A2:"))
        self.assertEqual("A2:I3", str(self.sheet.auto_filter.ref))

    def test_freeze_panes_shifted(self):
        """冻结了 1 列的（B2）右移到 C2，保证冻结的还是原来那一列。"""
        self.assertEqual("C2", self.sheet.freeze_panes)


class WritableColumnsTest(unittest.TestCase):
    """校验脚本与引擎的**可写列口径**必须一致（防漂移）。

    `tools/verify_fill_back.py` 靠"可写列"集合区分"引擎会写的列"与"必须原样保留的身份列"。
    引擎哪天新增一列（如 2026-10-10 的「原文件序号」）而校验脚本没跟上，
    校验会把新列当身份列 → 源表当然没有它 → **所有行都被误判成丢失**（假失败）。
    这里把两边口径钉住，任何一边改了都会红。
    """

    def test_writable_set_covers_engine_columns(self):
        import verify_fill_back as V
        for head in ("原文件序号", "设备名称（按标准修改加机组号）*", "KKS编码（修编后）*",
                     "父级KKS编码*", "上级设备名称", "上级KKS编码", "系统", "设备类型",
                     "设备类", "修编说明"):
            with self.subTest(head=head):
                self.assertTrue(V.is_writable(head), f"「{head}」是引擎会写的列，应算可写")

    def test_identity_columns_are_not_writable(self):
        import verify_fill_back as V
        for head in ("序号", "机组号", "原设备名称", "原KKS编码", "修改意见（审核人填写）",
                     "图纸名称或图号", "备注", "位置标识码"):
            with self.subTest(head=head):
                self.assertFalse(V.is_writable(head), f"「{head}」必须原样保留，不该算可写")


if __name__ == "__main__":
    unittest.main()
