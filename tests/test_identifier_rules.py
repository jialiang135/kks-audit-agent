# -*- coding: utf-8 -*-
"""段位化校验器的规则级回归测试。

**一条规则一个测试**，测试名直接标出所依据的总表规则号，便于对账
（`kks_gen/项目编码规则总表.md` §4 正则速查表 / §5 字符与禁用 / §8 热控测点 / §12 硬伤表）。

约定：`error` = 判错，`warn` = 提示，`pending` = **待核**（未收录等，**不得判为错误**，总表 §10.13）。
"""
from __future__ import annotations

import pathlib
import sys
import unittest

PROJECT = pathlib.Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT))
sys.path.insert(0, str(PROJECT / "kks_gen"))

import identifiers as I  # noqa: E402


def levels(code: str) -> dict[str, str]:
    """{segment: level}（同一段取最重级别）。"""
    weight = {"error": 3, "pending": 2, "warn": 1}
    out: dict[str, str] = {}
    for f in I.parse(code).findings:
        if f.segment not in out or weight[f.level] > weight[out[f.segment]]:
            out[f.segment] = f.level
    return out


class FunctionFaceTest(unittest.TestCase):
    """总表 §4 `4-01`~`4-11`：功能面标识段位与正则。"""

    def test_4_11_standard_12(self):
        """4-11｜标准 12 位设备码应无问题。"""
        self.assertEqual({}, levels("07AAB02UH001"))

    def test_4_11_internal_dot_is_error(self):
        """4-11 / §12.2｜功能面标识内部**不得**出现 `.` → 判错。"""
        self.assertEqual("error", levels("01HAC13.BP101").get("separator"))

    def test_4_04_h1_must_be_in_table(self):
        """4-04｜H₁ 仅表4.3.1-1 的 8 个取值（B/C/E/G/K/T/U/W）。"""
        self.assertTrue(I.re.match(rf"^[{I.H1_SET}]$", "G"))
        self.assertIsNone(I.re.match(rf"^[{I.H1_SET}]$", "A"))

    def test_4_05_plant_leading_zero(self):
        """4-05 / `4-08`｜编号必须写前导零：写 `7` 而非 `07` 就不成 12 位结构。"""
        self.assertEqual("unknown", I.parse("7AAB02UH001").kind)

    def test_4_07_system_class_known(self):
        """4-07｜系统分类码须命中附录A。"""
        self.assertEqual("full", I.system_known("AAB"))

    def test_12_3_two_letter_hit_is_pending_not_error(self):
        """§12.3｜中煤按 2 位定义系统码 → 仅前 2 位命中报**待核**，不得判错。"""
        code = "07CAX99UH001"          # CAX 应为 3 字母但只在 CA 段有名
        lv = levels(code).get("syscls")
        self.assertIn(lv, ("pending", None))

    def test_10_13_unknown_code_is_pending_not_error(self):
        """§10.13｜**「码未收录」只作待核线索，不得判为错误**（最易犯的错）。"""
        lv = levels("07ZZZ99ZZ999")
        self.assertNotEqual("error", lv.get("syscls"))
        self.assertNotEqual("error", lv.get("devcls"))


class BannedLetterTest(unittest.TestCase):
    """总表 §5.2 / §5.3：禁用字母必须**段位化**，不得对整串扫描（§12.1 硬伤 #1）。"""

    def test_12_1_A_type_segment_regex(self):
        """§12.1｜A 类型段的字符集就是 `[A-HJ-NP-Z]`（禁 I、O）。"""
        self.assertIsNone(I.re.match(I.LETTER, "I"))
        self.assertIsNone(I.re.match(I.LETTER, "O"))
        self.assertIsNotNone(I.re.match(I.LETTER, "J"))

    def test_12_14_terminal_allows_slash(self):
        """`4-50` / §12.14｜端子标识字符类型**开放**（含 `/`），不得据图注判 I/O 非法。"""
        self.assertEqual({}, levels(":X1/2.A3"))

    def test_5_2_4_terminal_not_judged_by_graph_note(self):
        """§5.2.4｜端子标识不得因含 I/O 被判非法（正文注中无此条）。"""
        self.assertNotEqual("error", levels(":I.O1").get("terminal"))

    def test_12_5_x2_bans_five_letters(self):
        """§12.5（5.2.12 注）｜X₂ 位禁 I、O、**X、Y、Z**，比通用 I/O 多禁 3 个。"""
        for banned in "IOXYZ":
            with self.subTest(x2=banned):
                self.assertNotIn(banned, I.X2_SET)

    def test_12_5_x1_uses_x_and_y(self):
        """§5.3.5｜X₁ 仅两个值：X=信号源、Y=信号应用（X、Y 在 X₁ 位**必用**、在 X₂ 位禁用）。"""
        self.assertEqual(("X", "Y"), I.X1_ALLOWED)


class SignalTest(unittest.TestCase):
    """总表 §8.5：信号标识（`4-42`~`4-49`）。"""

    def test_4_42_x1_only_xy(self):
        """`4-42`｜X₁ 仅 X/Y。"""
        self.assertEqual("error", levels(";AX01").get("signal"))

    def test_4_43_x2_banned_letter(self):
        """`4-43` / §12.5｜X₂ 取 X（在 X₂ 位禁用）→ 格式不符，判错。"""
        self.assertEqual("error", levels(";XX01").get("signal"))

    def test_4_49_x45_forced_when_bgh(self):
        """`4-45`/`4-46`（5.2.13）｜X₂∈{B,G,H} 时 X₄X₅ 必须落表5.2.13。"""
        self.assertEqual("error", levels(";XG07").get("x45"))
        self.assertIsNone(levels(";XG01").get("x45"))

    def test_4_45_x45_free_when_other_x2(self):
        """`4-46`｜X₂∉{B,G,H} 时 X₄X₅ 由工程约定 → 不判错。"""
        self.assertIsNone(levels(";XA07").get("x45"))


class PositionTest(unittest.TestCase):
    """总表 `4-37`/`4-39`（4.6.9）：R₁R₂ 两种模式。"""

    def test_4_37_indoor_floor_01_to_09(self):
        """`4-37`｜室内模式 R₁R₂ = 01~09（地上）。"""
        self.assertIsNone(levels("++07UBA01.03002").get("r12"))

    def test_4_37_underground_91_to_99(self):
        """`4-37` / §12.15｜地下负 1~9 层 = 91~99。"""
        self.assertIsNone(levels("++07UBA01.92002").get("r12"))

    def test_12_15_r12_illegal_value(self):
        """§12.15（4.6.9-2）｜R₁R₂ 合法集 = {01~09} ∪ {91~99}，`50` 非法 → 判错。"""
        self.assertEqual("error", levels("++07UBA01.50002").get("r12"))

    def test_4_39_area_mode_letters(self):
        """`4-39`｜区域模式 R₁R₂ 为字母（纵坐标）。"""
        self.assertIsNone(levels("++07UBA01.AB002").get("r12"))


class MeasurePointTest(unittest.TestCase):
    """总表 §8：热控测点专项（本项目核心）。"""

    def test_8_2_measure_point_recognised(self):
        """§8.2/§8.3（5.2.3-1、7.1.5）｜T₁=B 即测量回路/测点。"""
        self.assertTrue(I.is_measure_point("01BMT11BC101"))
        self.assertFalse(I.is_measure_point("01BMT11UH101"))

    def test_8_3_band_701_899_is_error(self):
        """§8.3 / `12.7`（编制说明 5.2.7 表5）｜701~899 为**锁定不可用** → 判错。"""
        self.assertEqual("error", levels("01BMT11BC701").get("devno"))
        self.assertEqual("error", levels("01BMT11BC801").get("devno"))

    def test_8_3_band_301_399_is_pending(self):
        """§8.3｜301~399 是"未用"而非"非法" → 待核，**不判错**。"""
        self.assertEqual("pending", levels("01BMT11BC301").get("devno"))

    def test_8_3_band_101_199_ok(self):
        """§8.3｜模拟量远传仪表号段 101~199 合法。"""
        self.assertIsNone(levels("01BMT11BC101").get("devno"))

    def test_8_3_band_999_combination_ok(self):
        """§8.3｜901~999 组合测量回路合法。"""
        self.assertIsNone(levels("01BMT11BC999").get("devno"))

    def test_8_4_t2_must_be_physical_quantity(self):
        """§8.4（附录B 表B.0.2）｜T₁=B 时 T₂ 应为物理量码；压力=BP、温度=BT、电流=BC。"""
        self.assertEqual("压力/真空", I.PHYSICAL_QUANTITY["BP"])
        self.assertEqual("温度", I.PHYSICAL_QUANTITY["BT"])
        self.assertEqual("电流", I.PHYSICAL_QUANTITY["BC"])

    def test_8_4_bb_not_invented(self):
        """§8.4 注 / §11.22｜`BB` 原表 OCR 漏行，**不得凭推定补**，故不列入物理量码。"""
        self.assertNotIn("BB", I.PHYSICAL_QUANTITY)

    def test_12_20_measure_system_is_measured_object(self):
        """§12.20｜测点系统码取**被测工艺系统**，不取仪控的 C（S₁=C 管的是仪控盘柜本身）。"""
        p = I.parse("01BMT11BC101")
        self.assertEqual("BMT", p.segments["sys"])
        self.assertNotEqual("C", p.segments["sys"][0])


class NumberingBandTest(unittest.TestCase):
    """总表 §7.14 / §7.18（表5.1.5、表5.2.7）：QM/WP 与开关类的**用途号段**。

    实测口径（热控点表 58,280 行）：本表 QM/WP 行 13,377 条里 10,734 在 001~299（通用）、
    2,627 在 401~499（疏水和排气用）——**本来就合表**，所以"表内段位一律不报"。
    真正要报的只有**表外段位**（约 96 行 / 0.16%），且一律是**待核**而非判错。
    """

    def test_7_14_qm_listed_bands_are_silent(self):
        """表5.1.5 表内四段：001~299 / 301~399 / 401~499 / 601~699 都不报。"""
        for code in ("01LAA10QM101", "01LAA10QM301", "01LAA10QM401", "01LAA10QM601"):
            with self.subTest(code=code):
                self.assertIsNone(levels(code).get("devno"), f"{code} 是表内段位，不应报")

    def test_wp_shares_the_same_bands(self):
        """表5.1.5 明写 QM 与 WP **共用同一套编号段**。"""
        self.assertIsNone(levels("01LAA10WP401").get("devno"))
        self.assertEqual("pending", levels("01LAA10WP501").get("devno"))

    def test_7_14_qm_out_of_band_is_pending_not_error(self):
        """表5.1.5 **未列** 501~599 / 701~999 → 报**待核**，不判错（总表 §8.3 缺口说明）。

        实测样本：`01LAA10QM501`（除氧器启动排汽电动门）落在 501。
        """
        lv = levels("01LAA10QM501")
        self.assertEqual("pending", lv.get("devno"))
        self.assertNotEqual("error", lv.get("devno"))

    def test_7_18_switch_codes_are_not_band_checked(self):
        """表5.2.7 **不查**（实测判定，见 `identifiers.QM_WP_BANDS` 上方注释）。

        依据：鄂州竣工版 259 行真开关类设备**全部落在表内**（规则零触发、零收益）；
        热控点表却用 QA 标电磁阀（附录B 的 QA 实例首项就是「电磁阀」，用法正确）、
        用 QB 标流量喷嘴 —— 套表5.2.7 会产出 82 条**误导性**提示：
        对着流量喷嘴讲"母线侧隔离开关号段"。宁可不报，也不报错方向。
        """
        for code in ("01ABC02QA401", "01ABC02QB501", "01ABC02QC500"):
            with self.subTest(code=code):
                self.assertIsNone(levels(code).get("devno"),
                                  f"{code} 不应套用表5.2.7 号段")

    def test_measure_points_still_use_their_own_bands(self):
        """守坑：测点（T₁=B）走**编制说明表5**那套，不能套用 QM/WP 的表5.1.5。

        `BC401` 在测点表里是"验收试验测点"（合法），若误套表5.1.5 会被判成表外段位。
        """
        self.assertIsNone(levels("01BMT11BC401").get("devno"))
        self.assertEqual("error", levels("01BMT11BC701").get("devno"))

    def test_band_check_silent_for_other_device_classes(self):
        """其它设备分类码（非测点、非 QM/WP/QA/QB/QC）不做号段判定。"""
        self.assertIsNone(levels("01LAB13GP101").get("devno"))


class ExcludedCodeTest(unittest.TestCase):
    """总表 §10.1 / §12.11：中煤剔除的核电码须排除。"""

    def test_12_11_jk_excluded(self):
        """§12.11（中煤细则 T25）｜`J`/`K` 前缀核电码在**中煤体系**下非法 → 判错。"""
        self.assertEqual("error", levels("07JAA01AA001").get("syscls"))

    def test_12_11_excluded_set_size(self):
        """§10.1｜剔除集 = 393 条（J 204 + K 189，含两个 1 字母主组码本身）。"""
        self.assertEqual(393, len(I.EXCLUDED))


class MiscIdentifierTest(unittest.TestCase):
    """其余标识类型。"""

    def test_4_56_document_format(self):
        """`4-56`｜文档标识 = `&` + 3 字母 + 3 位编号，无 `.` 无括号。"""
        self.assertIsNone(levels("&ABC001").get("document"))
        self.assertEqual("error", levels("&AB001").get("document"))

    def test_4_03_union_two_levels(self):
        """`4-03`｜联合标识两级，各级 ≤12 位，中间 `.`。"""
        self.assertIsNone(levels("#C1.U1").get("union"))

    def test_4_11_plus_suffix_is_warn_not_error(self):
        """§1.11 + 竣工版实测｜`07BFA01UC001.A1` 是"功能面 + 附加标识"的工程写法：
        国标 4.5.5 要求用 `-` 组合，但竣工版用 `.` → 报**警告**，不判错。"""
        lv = levels("07BFA01UC001.A1")
        self.assertEqual("warn", lv.get("separator"))
        self.assertIsNone(lv.get("h23"))

    def test_readme_2_phase_suffix(self):
        """README §2｜字母尾缀表分相（`A/B/C`），按功能面接收并提示，不判错。"""
        p = I.parse("07ABG01BA001A")
        self.assertEqual("function", p.kind)
        self.assertEqual("A", p.segments.get("phase"))
        self.assertFalse(p.has_error)

    def test_no_dot_in_function_face_but_dot_in_5_types(self):
        """§1.11 / §12.2｜`.` 只用于 5 类标识；功能面内部没有点。"""
        for ok in (":A.B", ";XG01.YY01", "++07UBA01.03002"):
            with self.subTest(code=ok):
                self.assertNotEqual("error", levels(ok).get("separator"))


class DictionaryGuardTest(unittest.TestCase):
    """判据字典缺失时必须**大声报错**，不许静默降级（否则每个码都会被判成"未收录"）。"""

    def test_missing_dictionary_raises(self):
        import pathlib as _p
        import tempfile
        saved = I.CONFIG
        try:
            with tempfile.TemporaryDirectory() as td:
                I.CONFIG = _p.Path(td)
                with self.assertRaises(RuntimeError) as cm:
                    I._require(I._load("project_system_codes.json"),
                               "project_system_codes.json", "见报错信息")
                self.assertIn("缺失或为空", str(cm.exception))
        finally:
            I.CONFIG = saved

    def test_dictionaries_loaded(self):
        """正常环境下字典非空（`config/` 未入库时会触发上面的守护）。"""
        self.assertGreater(len(I.SYS_DICT), 1000)
        self.assertGreater(len(I.DEV_DICT), 100)


if __name__ == "__main__":
    unittest.main()
