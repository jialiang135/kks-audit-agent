# -*- coding: utf-8 -*-
"""修编引擎的规则级回归测试（`kks_gen/revise.py`）。

**一条规则一个测试**，测试名标出所依据的总表规则号 / 硬伤编号；
覆盖总表 §12「引擎硬伤表」里**与引擎行为**相关的那几条（校验层规则见
`test_identifier_rules.py`）。

判据出处见 `kks_gen/项目编码规则总表.md` 与 `kks_gen/revise.py` 模块文档。
"""
from __future__ import annotations

import pathlib
import sys
import unittest

PROJECT = pathlib.Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT))
sys.path.insert(0, str(PROJECT / "kks_gen"))

import revise as R  # noqa: E402
import identifiers as I  # noqa: E402


def unknown_system_code() -> str:
    """找一个**真正未收录**的 3 字母系统码：它本身、以及它的前 2 位，都不在码表内。

    为什么要动态找：写死 `XZZ` 会踩坑——`XZ` 恰好是 2 字母命中，按总表 §12.3
    属"仅前两位命中 → **待核但不改动**"，不是"未收录"。（实测踩过这个假设。）
    """
    letters = "ABCDEFGHJKLMNPQRSTUVWXYZ"
    for a in letters:
        for b in letters:
            for c in letters:
                code = a + b + c
                if code not in I.SYS_DICT and code[:2] not in I.SYS_DICT:
                    return code
    raise AssertionError("找不到未收录码（字典被污染？）")


class PlantCodeTest(unittest.TestCase):
    """总表 §3（表4.3.1-2）：全厂码与机组号。"""

    def test_4_3_1_2_plant_matches_unit_label(self):
        """§3｜全厂码与「机组号」列冲突 → 待人工裁定归属（本项目 07=7号机组、08=8号机组）。"""
        r = R.revise_code("08AAB02UH001", unit_label="7号机组")
        self.assertEqual("unresolved", r.status)
        self.assertEqual("", r.revised)
        self.assertEqual("08AAB02UH001", r.original)

    def test_no_unit_label_means_no_guess(self):
        """§12 纪律「绝不猜」｜机组号列给不出结论时，**不得**改全厂码。"""
        r = R.revise_code("08AAB02UH001", unit_label="待复核")
        self.assertEqual("ok", r.status)
        self.assertEqual("08AAB02UH001", r.revised)

    def test_bare_gongyong_does_not_flip_61_to_00(self):
        """🔴 裸「公用」**不得**把 61（两机共用）改成 00（全厂公用）——2026-10-10 实测纠正。

        实测（热控点表）：`61LAC10`（电动给水泵）同一系统下并存「公用」与「1、2号机组共用」
        两种写法；全厂只有 3 行真·全厂公用（500kV GIS）用的是 `00`。
        原实现见「公用」就判 00 → 改错了 127 行，还把同一系统劈成两个分支。
        """
        r = R.revise_code("61LAC10BP111", unit_label="公用")
        self.assertEqual("ok", r.status)
        self.assertEqual("61LAC10BP111", r.revised, "编码不得被改动")
        self.assertIn("需人工确认", r.note, "要在说明里提请人工判断")

    def test_explicit_plant_wide_still_flips_to_00(self):
        """写明全厂公用与原61冲突时，不自动改变设备归属。"""
        r = R.revise_code("61LAC10BP111", unit_label="全厂公用")
        self.assertEqual("unresolved", r.status)
        self.assertEqual("", r.revised)
        self.assertIn("冲突", r.note)

    def test_missing_plant_code_with_ambiguous_label_explains_why(self):
        """缺全厂码 + 机组号只写「公用」→ 待人工，且说明要**说清是歧义**（不是"机组号为空"）。"""
        r = R.revise_code("CAE01", unit_label="公用")
        self.assertEqual("unresolved", r.status)
        self.assertEqual("", r.revised)
        self.assertIn("公用", r.note)
        self.assertIn("无法区分", r.note)


class SystemCodeEvidenceTest(unittest.TestCase):
    """总表 §12.16（实测）：系统分类码**只认「系统」列的证据**。"""

    def test_12_16_never_guess_system_from_device_name(self):
        """§12.16｜**禁止**靠设备名称关键词猜系统分类码——实测 656 行证据强度全为零。

        系统码未收录、描述里写着"给水"（看似指向给水系统），但**没有「系统」列** → 必须不动。
        """
        bad = unknown_system_code()
        code = f"01{bad}01BP101"
        r = R.revise_code(code, unit_label="1号机组",
                          description="给水系统出口压力", system_name="")
        self.assertEqual("ok", r.status)
        self.assertEqual(code, r.revised)
        self.assertEqual(bad, r.revised[2:5])          # 系统分类码未被改动
        self.assertTrue(r.needs_review)                # 但标为待核（§10.13）

    def test_system_changed_only_with_system_column(self):
        """有「系统」列证据时才改系统分类码。"""
        bad = unknown_system_code()
        r = R.revise_code(f"01{bad}01BP101", unit_label="1号机组",
                          description="", system_name="给水管道系统")
        self.assertEqual("revised", r.status)
        self.assertNotEqual(bad, r.revised[2:5], r.revised)
        self.assertIn("系统分类码", r.note)

    def test_12_3_prefix_hit_is_pending_but_not_changed(self):
        """§12.3｜系统码只命中前 2 位（中煤按 2 位定义）→ **待核**，不得改动、不得判错。"""
        import identifiers as I2
        prefix = next(c for c in ("XZ", "CA", "HN", "LB") if c in I2.SYS_DICT)
        code = f"01{prefix}X01BP101"
        r = R.revise_code(code, unit_label="1号机组")
        self.assertEqual("ok", r.status)
        self.assertEqual(code, r.revised)
        self.assertEqual("pending", r.level)


class UnresolvedVsPendingTest(unittest.TestCase):
    """总表 §10.13 / §12.8：**「码未收录」只作待核线索，不判为错误**。"""

    def test_10_13_unknown_code_kept_not_cleared(self):
        """§10.13｜未收录码：**保留原码**、status=ok、needs_review=True、level=pending。"""
        r = R.revise_code("01XZZ01ZZ101", unit_label="1号机组")
        self.assertEqual("ok", r.status)
        self.assertEqual("01XZZ01ZZ101", r.revised)
        self.assertTrue(r.needs_review)
        self.assertEqual("pending", r.level)

    def test_10_13_note_says_pending_not_error(self):
        """§10.13｜说明里必须写明"待核线索"，不能写成"非法码"。"""
        r = R.revise_code("01XZZ01ZZ101", unit_label="1号机组")
        self.assertIn("未收录", r.note)
        self.assertIn("待核", r.note)

    def test_12_11_excluded_nuclear_code_unresolved(self):
        """§12.11（中煤细则 T25）｜J/K 前缀核电码在中煤体系下非法 → 留空交人工。"""
        r = R.revise_code("07JAA01AA001", unit_label="7号机组")
        self.assertEqual("unresolved", r.status)
        self.assertEqual("", r.revised)
        self.assertEqual("error", r.level)
        self.assertIn("核电码", r.note)


class DotRuleTest(unittest.TestCase):
    """总表 §1.11 / §12.2：`.` 只用于 5 类标识，功能面内部不得有点。"""

    def test_12_2_internal_dot_removed(self):
        """§12.2｜功能面内部插点 → 去掉（有明确国标依据，可直接修）。"""
        r = R.revise_code("01HAC13.BP101", unit_label="1号机组")
        self.assertEqual("revised", r.status)
        self.assertEqual("01HAC13BP101", r.revised)

    def test_suffix_after_dot_preserved(self):
        """竣工版实测｜`07BFA01UC001.A1` 的非标准附加写法须留原码证据，不直接写成合规新码。"""
        r = R.revise_code("07BFA01UC001.A1", unit_label="7号机组")
        self.assertEqual("", r.revised)
        self.assertEqual("07BFA01UC001.A1", r.original)

    def test_phase_suffix_preserved(self):
        """`kks_gen/README.md` §2｜字母尾缀须保留原码证据并待核工程约定。"""
        r = R.revise_code("07ABG01BA001A", unit_label="7号机组")
        self.assertEqual("", r.revised)
        self.assertEqual("unresolved", r.status)
        self.assertEqual("07ABG01BA001A", r.original)


class MeasurePointTest(unittest.TestCase):
    """总表 §8 / §12.7：测点专项（本项目核心）。"""

    def test_12_7_band_701_899_unresolved(self):
        """§12.7（编制说明 5.2.7 表5）｜测量回路编号 701~899 是**锁定不可用** → 留空交人工。"""
        r = R.revise_code("01BMT11BC701", unit_label="1号机组")
        self.assertEqual("unresolved", r.status)
        self.assertEqual("", r.revised)
        self.assertEqual("error", r.level)
        self.assertIn("701", r.note)

    def test_8_3_band_101_199_ok(self):
        """§8.3｜模拟量远传测点落 101~199 → 合规。"""
        r = R.revise_code("01BMT11BC101", unit_label="1号机组")
        self.assertEqual("ok", r.status)
        self.assertEqual("01BMT11BC101", r.revised)

    def test_8_4_device_code_by_description(self):
        """§8.4（附录B）｜设备分类码可按**业务描述**匹配（描述讲的就是被测物理量）。"""
        r = R.revise_code("01BMT11ZZ001", unit_label="1号机组",
                          description="凝结水泵出口压力")
        self.assertEqual("BP", r.revised[7:9])      # 全厂码2+系统码5 = 7 位之后是 T₁T₂
        self.assertEqual("revised", r.status)


class NumberPaddingTest(unittest.TestCase):
    """总表 `4-05`/`4-08`/`4-10`（4.4.4-2、4.4.5-2）：每位上的 0 必须写出。"""

    def test_under_padded_numbers_are_padded(self):
        """编号少写前导零 → **补足**（`2`→`02`、`1`→`001`）。

        实测这是真实数据的主要形态之一（鄂州四期 `07PAB02QM01`、`07HAG70QM31` 等共 200+ 行）。
        解读唯一（剩余字符是"字母段+数字段"且只有一种切法）→ 有明确国标依据，可直接修。
        """
        r = R.revise_code("07AAB2UH1", unit_label="7号机组")
        self.assertEqual("revised", r.status)
        self.assertEqual("07AAB02UH001", r.revised)
        self.assertEqual("07AAB2UH1", r.original)

    def test_pad_measure_point_number(self):
        """`07PAB02QM01` → `07PAB02QM001`（设备编号补足 3 位）。"""
        r = R.revise_code("07PAB02QM01", unit_label="7号机组")
        self.assertEqual("revised", r.status)
        self.assertEqual("07PAB02QM001", r.revised)

    def test_ambiguous_code_left_to_human(self):
        """歧义码（数字段里混了字母，无法唯一切段）→ **留空交人工**，不猜。"""
        r = R.revise_code("07LAB13BBP101", unit_label="7号机组")
        self.assertEqual("unresolved", r.status)
        self.assertEqual("", r.revised)

    def test_valid_length_numbers_untouched(self):
        """已是标准位数的编号 → 原样保留。"""
        r = R.revise_code("07AAB02UH001", unit_label="7号机组")
        self.assertEqual("ok", r.status)
        self.assertEqual("07AAB02UH001", r.revised)


class ExtendedIdentifierTest(unittest.TestCase):
    """竣工版实测的扩展形态（`kks_gen/README.md` §2/§5）：必须识别且**原样保留**。"""

    def test_seven_digit_system_code_is_compliant(self):
        """7 位系统码（`07BFA01`）**本身合规**，不得要求它补设备分类码。

        踩过：早期实现把它判成"缺设备分类码、待人工"，一次误伤 132 行合规系统码。
        """
        r = R.revise_code("07BFA01", unit_label="7号机组")
        self.assertEqual("ok", r.status)
        self.assertEqual("07BFA01", r.revised)
        self.assertFalse(r.needs_review)

    def test_five_digit_system_code_missing_sysno(self):
        """5 位（`07RED`）= 全厂码+系统分类码，**缺系统编号** → 保留原码但标待核。"""
        r = R.revise_code("07RED", unit_label="7号机组")
        self.assertEqual("ok", r.status)
        self.assertEqual("07RED", r.revised)
        self.assertEqual("pending", r.level)
        self.assertIn("系统编号", r.note)

    def test_part_suffix_hyphen_preserved(self):
        """`61EAB10GL001-M01` —— 设备码 + 连字符部件码，原始证据保留，缺约定时不自动认定合规。"""
        r = R.revise_code("61EAB10GL001-M01", unit_label="6、1号机组共用")
        self.assertEqual("", r.revised)
        self.assertEqual("unresolved", r.status)
        self.assertEqual("61EAB10GL001-M01", r.original)

    def test_part_suffix_glued_preserved(self):
        """`64GHA10GP001BT301` —— 设备码 + 直接拼接的部件/测点码（`BT301`=温度）。"""
        r = R.revise_code("64GHA10GP001BT301", unit_label="6、4号机组共用")
        self.assertEqual("", r.revised)
        self.assertEqual("64GHA10GP001BT301", r.original)

    def test_position_number_style_left_to_human(self):
        """设计院**位号**（`PQ1`、`XF-1-1`）不是 KKS 码 → 留空交人工，不自作聪明改码。"""
        for code in ("PQ1", "XF-1-1", "SF(B)-1-1"):
            with self.subTest(code=code):
                r = R.revise_code(code, unit_label="7号机组")
                self.assertEqual("unresolved", r.status)
                self.assertEqual("", r.revised)


class PointCodeNotEncodedTest(unittest.TestCase):
    """指令/测点行（原码 = 设备码 + 信号码）**不编码**（业务口径 2026-10-12）。

    口径：**这些是指令和测点，发出指令/信号的那个才是设备，只有设备才有 KKS 编码。**
    回归动机：旧做法把这类行**归一到设备本体码**，12 个不同测点全落成同一个 `01HLB10GQ001`，
    把一张本来无重复的表（输入 58,276 码 / 唯一 58,274）做成 **3,563 类、18,088 行重复**。
    """

    def test_signal_source_point_not_encoded(self):
        """`XB21`（信号标识 X₁=X 信号源 → 测点）不编码，只给出所服务的设备本体码。"""
        r = R.revise_code("01HLB10GQ001XB21", unit_label="1号机组")
        self.assertEqual("not_device", r.status)
        self.assertEqual("", r.revised)
        self.assertEqual("01HLB10GQ001", r.device_body)
        self.assertIn("01HLB10GQ001", r.note)

    def test_signal_application_command_not_encoded(self):
        """`YB21`（信号标识 X₁=Y 信号应用 → 启动/停止指令）同样不编码。"""
        r = R.revise_code("01HLB10GQ001YB21", unit_label="1号机组")
        self.assertEqual("not_device", r.status)
        self.assertEqual("", r.revised)
        self.assertEqual("01HLB10GQ001", r.device_body)

    def test_all_points_of_one_device_get_no_code(self):
        """同一设备的多个测点**一律留空**，绝不收敛成同一个设备码（重复值就是这里来的）。"""
        points = ["01HLB10GQ001",      # 设备本体行 → 有码
                  "01HLB10GQ001XB21", "01HLB10GQ001XB22", "01HLB10GQ001YB21",
                  "01HLB10GQ001XB41", "01HLB10GQ001XB61", "01HLB10GQ001XQ01"]
        revs = [R.revise_code(c, unit_label="1号机组") for c in points]
        self.assertEqual("ok", revs[0].status)
        self.assertEqual("01HLB10GQ001", revs[0].revised, "设备本体行必须保留自己的码")
        for c, r in zip(points[1:], revs[1:]):
            with self.subTest(code=c):
                self.assertEqual("", r.revised)
                self.assertEqual("not_device", r.status)

    def test_defers_nothing_to_human_review(self):
        """口径已定，不属"待核"——不能混进"需人工核对"的统计里。"""
        r = R.revise_code("01HAL11GP001XQ01", unit_label="1号机组")
        self.assertFalse(r.needs_review)

    def test_glued_part_code_unaffected(self):
        """⚠️ 只认 `X₁ ∈ {X,Y}` 的信号码。`BT301`（温度测量部件）首字母不是 X/Y，
        仍按"设备码 + 部件后缀"**保留整码**，不能被误判成指令/测点。"""
        r = R.revise_code("64GHA10GP001BT301", unit_label="6、4号机组共用")
        self.assertEqual("", r.revised)
        self.assertEqual("64GHA10GP001BT301", r.original)
        self.assertNotEqual("not_device", r.status)


class EmptyAndStructureTest(unittest.TestCase):
    """空值与非功能面标识的处置。"""

    def test_empty_code(self):
        """空编码 → 留空 + 说明。"""
        r = R.revise_code("", unit_label="1号机组")
        self.assertEqual("unresolved", r.status)
        self.assertEqual("原编码为空", r.note)

    def test_short_code_diagnosed(self):
        """长度不足 → 给可执行诊断（不是泛泛的"格式错误"）。"""
        r = R.revise_code("00123", unit_label="1号机组")
        self.assertEqual("unresolved", r.status)
        self.assertIn("5 位", r.note)

    def test_signal_identifier_not_revised_but_validated(self):
        """非功能面标识（信号/端子等）只做校验说明，不做修编。"""
        r = R.revise_code(";XG01", unit_label="1号机组")
        self.assertEqual("not_device", r.status)
        self.assertEqual("", r.revised)
        self.assertIn("信号标识", r.note)

    def test_bad_signal_unresolved(self):
        """信号标识格式错 → 留空 + 出处。"""
        r = R.revise_code(";XG07", unit_label="1号机组")
        self.assertEqual("unresolved", r.status)
        self.assertIn("5.2.13", r.note)


class BatchTest(unittest.TestCase):
    """批量接口与统计。"""

    def test_revise_rows_and_summary(self):
        """`revise_rows` / `revise_summary` 与单条口径一致。"""
        rows = [{"code": "01BMT11BC101", "unit_label": "1号机组"},
                {"code": "01BMT11BC701", "unit_label": "1号机组"},
                {"code": "01BMT11ZZ101", "unit_label": "1号机组"}]
        revs = R.revise_rows(rows)
        self.assertEqual(3, len(revs))
        s = R.revise_summary(revs)
        # `status=ok` 含两类：真正合规的，以及"待核但保留原码"的（§10.13 不判错）
        # —— 报表里靠 `pending_only` / `needs_review` 把后者单独列出，避免"待核"被算成"合规"
        self.assertEqual(2, s["ok"])
        self.assertEqual(1, s["unresolved"])            # BC701 号段锁定
        self.assertEqual(1, s["pending_only"])          # ZZ101 设备码未收录 → 待核保留
        self.assertEqual(2, s["needs_review"])          # unresolved + 待核


class UnitLabelInferenceTest(unittest.TestCase):
    """机组号上下文推断（只在证据足够时推，且绝不猜）。"""

    def test_infer_from_same_drawing(self):
        """同一「图纸名称或图号」的行互相补全机组号（强证据）。"""
        rows = [{"unit_label": "", "drawing": "图A", "system": ""},
                {"unit_label": "7号机组", "drawing": "图A", "system": ""},
                {"unit_label": "7号机组", "drawing": "图A", "system": ""}]
        self.assertEqual("7号机组", R.infer_unit_labels(rows)[0])

    def test_no_evidence_returns_empty(self):
        """证据不足 → 返回空串，保持"待人工"（**绝不猜**）。"""
        rows = [{"unit_label": "", "drawing": "图A", "system": ""},
                {"unit_label": "7号机组", "drawing": "图B", "system": ""},
                {"unit_label": "8号机组", "drawing": "图C", "system": ""}]
        self.assertEqual("", R.infer_unit_labels(rows)[0])


class DictionaryConsistencyTest(unittest.TestCase):
    """判据字典与引擎读表口径（总表 §10 / §13）。"""

    def test_engine_reads_project_dictionaries(self):
        """§13.2｜引擎读**项目可用码表**（国标 ∪ 中煤），不是只用国标附录。"""
        import identifiers as I
        self.assertGreater(len(I.SYS_DICT), len(I.GB_SYS_DICT))
        self.assertIn("PA", I.DEV_DICT)          # 中煤独创设备码（固定摄像机）
        self.assertIn("BA", I.DEV_DICT)          # 附录B 表B.0.2 电位（曾因 OCR 跨行丢行）

    def test_12_9_separate_tables_by_segment(self):
        """§12.9｜按段位分别建表：系统码查附录A、设备码查附录B，不混用。

        `BY` 是附录A 的系统码（控制和保护系统），**不是**设备分类码。
        """
        import identifiers as I
        self.assertIn("BY", I.SYS_DICT)
        self.assertNotIn("BY", I.DEV_DICT)


if __name__ == "__main__":
    unittest.main()
