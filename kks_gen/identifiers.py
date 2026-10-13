# -*- coding: utf-8 -*-
"""KKS 标识段位化校验器 —— `kks_gen/项目编码规则总表.md` 第 4 节「正则速查表」的代码实现。

## 为什么单独成模块

修编引擎原来把"禁用字母 I/O"当**全局规则**对整串扫描（总表 §12.1 硬伤 #1），
于是 `01HAC13.BP101` 这类**含 `.` 的功能面标识**、端子标识里的合法字符都被误判。
正确做法是：**先按前缀符判标识类型 → 再按段位逐段校验**。
本模块就是这一层，`revise.py` 只负责"怎么改"，校验交给这里。

## 输出三级（对应总表情态动词映射 §0.5）

| 级别 | 含义 | 依据 |
|---|---|---|
| `error` | **判错**（"应/不应/不得/必须"级别的硬约束被违反） | 总表 §0.5 |
| `warn` | 异常但**不判错**（"宜"、源文件未规定、原文用词宽松处） | 总表 §0.5、§11 |
| `pending` | **待核**：码未收录、OCR 缺口、原文未定义段位等，**不得判为错误** | 总表 §10.13「未收录只作待核线索」、§12.8 |

## 依据（逐条可回溯到总表）

- 结构骨架 图4.4.1 / 4.7.1；`.` 只用于 5 类标识（总表 §1.11、§12.2）
- 正则 总表 §4 `4-01`~`4-63`
- 禁用字母段位化 总表 §5.2、§5.3（X₂ 位额外禁 X、Y、Z）
- 测点专项 总表 §8（T₁=B、T₃T₄T₅ 号段、701~899 锁定）
- R₁R₂ = {01~09} ∪ {91~99}  4.6.9-2（总表 `4-37`）
- 端子字符集开放（含 `/`）、各级 ≤12 位  4.8.1 注1（总表 `4-50`~`4-52`）
"""
from __future__ import annotations

import json
import pathlib
import re
from dataclasses import dataclass, field

CONFIG = pathlib.Path(__file__).resolve().parents[1] / "config"


def _load(name: str) -> dict:
    path = CONFIG / name
    if not path.is_file():
        return {}
    return json.loads(path.read_text(encoding="utf-8"))


def _require(table: dict, name: str, howto: str) -> dict:
    """判据字典缺失/为空时**立即报错**——绝不允许静默降级。

    🔴 为什么必须报错：字典若为空，`system_known()`/`device_known()` 会把**每一个码**
    都判成"未收录"，引擎表面上照样跑完、结果却全是"待核"，而且**没有任何报错**。
    这个失败模式极难发现（曾经就是这样把几千行合法码判成未核）。
    """
    if not table:
        raise RuntimeError(
            f"判据字典 `config/{name}` 缺失或为空，引擎无法工作。\n"
            f"  → 重新生成：{howto}\n"
            f"  → 注意 `.venv` 里没有 python-docx，读 .docx 要用 PyCharm 的 python。"
        )
    return table


# 项目可用码表（国标附录 ∪ 中煤细则）= 判据。总表 §13.1/13.2
_SYS_HOWTO = "python -B tools/extract_gb_dictionaries.py && python -B tools/extract_zhongmei_codes.py"
SYS_DICT: dict[str, dict] = _require(
    _load("project_system_codes.json") or _load("gb_system_codes.json"),
    "project_system_codes.json", _SYS_HOWTO)
DEV_DICT: dict[str, dict] = _require(
    _load("project_device_codes.json") or _load("gb_device_codes.json"),
    "project_device_codes.json", _SYS_HOWTO)
# 仅国标层，用于区分"未收录"与"中煤独创"（总表 §10）
GB_SYS_DICT: dict[str, dict] = _load("gb_system_codes.json")
EXCLUDED: dict[str, dict] = _load("zhongmei_excluded_codes.json")

# ---------------------------------------------------------------- 字符集
# 总表 §5.1：A = 大写英文字母（**图注：禁用 I、O**）→ A 类型段专用
LETTER = "[A-HJ-NP-Z]"          # A 类型（禁 I、O）
DIGIT = "[0-9]"
# 端子标识字符类型是**开放的** `A…/…N`（4.8.1 注1），**不得**据图注判 I/O 非法
TERMINAL_CHAR = "[A-Z0-9/]"
# 全厂码 H₁ 的 8 个取值（表4.3.1-1 / 总表 `4-04`）
H1_SET = "BCEGKTUW"
# 信号码 X₁ 仅两个值（表5.2.11）：X = 信号源、Y = 信号应用
X1_ALLOWED = ("X", "Y")
# 信号码 X₂ 可用集：A B C G H J K L M P Q R S U V W（15 个）；禁 I、O、X、Y、Z（5.2.12 注）
X2_SET = "ABCGHJKLMPQRSUVW"
# X₂ ∈ {B,G,H} 时 X₄X₅ 的强制取值（表5.2.13）
X45_FORCED = {"01", "02", "03", "04", "05", "06",
              "11", "12", "13", "21", "22", "23",
              "31", "32", "33", "51", "52", "53", "54", "55", "56", "99"}

# 功能面标识内部**无 `.`**（总表 §1.11、§12.2）
RE_FUNCTION = re.compile(
    rf"^(?P<h1>[{H1_SET}])?(?P<h23>{DIGIT}{{2}})(?P<h4>{DIGIT})?"
    rf"(?P<sys>{LETTER}{{3}})(?P<sysno>{DIGIT}{{2}})"
    rf"(?P<dev>{LETTER}{{2}})(?P<devno>{DIGIT}{{3}})$"
)
RE_SYSTEM_ONLY = re.compile(rf"^(?P<h1>[{H1_SET}])?(?P<h23>{DIGIT}{{2}})(?P<h4>{DIGIT})?"
                            rf"(?P<sys>{LETTER}{{3}})(?P<sysno>{DIGIT}{{2}})$")

# 宽松版功能面：允许编号**位数不足**（少写前导零）。国标 4.4.4-2 / 4.4.5-2 要求"每位上的 0 必须写出"，
# 所以 `07PAB02QM01` 这种是**可明确修复**的（编号 2 位 → 补成 3 位），不该直接判"无法解析"。
# 实测：鄂州四期 11 位码 218 行、5 位码 388 行都属这一类。
RE_FUNCTION_LENIENT = re.compile(
    rf"^(?P<h1>[{H1_SET}])?(?P<h23>{DIGIT}{{2}})(?P<h4>{DIGIT})?"
    rf"(?P<sys>{LETTER}{{3}})(?P<sysno>{DIGIT}{{1,2}})"
    rf"(?P<dev>{LETTER}{{2}})(?P<devno>{DIGIT}{{1,3}})$"
)
# 部件/附加后缀：设备码之后直接拼的分类码 + 编号（`KE01`、`BT301`、`-M01`），可带连字符。
# 属国标允许的"扩展/部件"层级（README §2、§5）；本引擎**不生成**部件，但**必须能识别并原样保留**。
RE_PART_SUFFIX = re.compile(rf"^-?{LETTER}{{1,2}}{DIGIT}{{2,3}}$")
# 产品面标识 `-`：P1 = P₁P₂(P₃)P₄P₅，P2 整级可重复（4-20~4-22）
RE_PRODUCT_SEG = re.compile(rf"^{LETTER}{{2}}({DIGIT})?{DIGIT}{{2}}$")
# 安装点标识 `+`（4-25~4-33）
RE_INSTALL = re.compile(
    rf"^\+?(?P<h1>[{H1_SET}])?(?P<h23>{DIGIT}{{2}})(?P<h4>{DIGIT})?"
    rf"(?P<sys>{LETTER}{{3}})(?P<sysno>{DIGIT}{{2}})"
    rf"(?:(?P<dev>{LETTER}{{2}})(?P<devno>{DIGIT}{{3}}))?"
    rf"\.(?P<e12>{LETTER}{{2}})(?P<e345>{DIGIT}{{3}})$"
)
# 位置标识 `++`（4-34~4-41）
# ⚠️ R₁R₂ 有两种模式（4.6.9-1 / 4.6.9-3）：**室内模式 = 数字 `NN`**（地上 01~09、地下 91~99），
#    **区域模式 = 字母 `AA`**（纵坐标）。正则必须两种都收，否则合法位置标识会被误判成格式错。
RE_POSITION = re.compile(
    rf"^\+\+(?P<h1>[{H1_SET}])?(?P<h23>{DIGIT}{{2}})(?P<h4>{DIGIT})?"
    rf"(?P<build>{LETTER}{{3}})(?P<buildno>{DIGIT}{{2}})"
    rf"\.(?P<r12>[A-Z0-9]{{2}})(?P<r345>{DIGIT}{{3}})$"
)
# 信号标识 `;`（4-42~4-49）
RE_SIGNAL = re.compile(
    rf"^;(?P<x1>[XY])(?P<x2>[{X2_SET}])(?P<x3>{DIGIT})?(?P<x45>{DIGIT}{{2}})"
    rf"(?:\.(?P<y12>{LETTER}{{2}})(?P<y34>{DIGIT}{{2}}))?$"
)
# 端子标识 `:`（4-50~4-52）：字符集开放，各级 ≤12 位，第 1 级可省
RE_TERMINAL = re.compile(rf"^:(?:({TERMINAL_CHAR}{{1,12}})\.)?({TERMINAL_CHAR}{{1,12}})$")
# 文档标识 `&`（4-54~4-56）
RE_DOCUMENT = re.compile(rf"^&{LETTER}{{3}}{DIGIT}{{3}}$")
# 联合标识 `#`（4-01~4-03）：两级，各级 ≤12 位，字母段在前数字段在后，中间 `.`
RE_UNION = re.compile(rf"^#[A-HJ-NP-Z0-9]{{1,12}}\.[A-HJ-NP-Z0-9]{{1,12}}$")
# 中煤设备码表里"视频监控/门禁"的 16 个码（用于识别中煤独创设备码，总表 §10.1）
ZM_UNIQUE_DEV = set("PA PB PC PD PE PV PU PR PT PL PP PN PM KB KC KD".split())

# 测点 T₃T₄T₅ 号段（编制说明 5.2.7 表5；总表 §8.3）
MEASURE_BANDS: list[tuple[int, int, str, str]] = [
    (1, 99, "开关量远传仪表", "ok"),
    (101, 199, "模拟量远传仪表", "ok"),
    (201, 299, "测试设备", "ok"),
    (301, 399, "未用", "pending"),
    (401, 499, "验收试验测点", "ok"),
    (501, 599, "就地测量回路", "ok"),
    (601, 699, "未用（可约定使用）", "pending"),
    (701, 799, "锁定（不可用）", "error"),
    (801, 899, "锁定（不可用）", "error"),
    (901, 999, "组合测量回路", "ok"),
]

# QM/WP 共用同一套编号段（**表5.1.5**；总表 `7.14`）。
# ⚠️ 501~599 与 701~999 **原表未列** → 遇此段报**待核**，**不判错**（总表 §8.3 缺口说明）。
QM_WP_BANDS: tuple[tuple[int, int, str], ...] = (
    (1, 299, "未进一步区分（通用）"),
    (301, 399, "测量用"),
    (401, 499, "疏水和排气用（含排污）"),
    (601, 699, "取样用"),
)

# ------------------------------------------------------------------ 为什么不查表5.2.7
# 表5.2.7（QA/QB/QC 的 101~199 进线 / 201~299 出线 / 301~399 联络 / 101~499 母线侧…）
# **看起来该查，实测却不该查**（2026-10-10，两份真实表量过）：
#   · 鄂州竣工版 259 行开关类设备（名称含「开关/隔离/接地/刀闸」）**全部落在表内**段位
#     → 该规则**一次都不触发**，加它等于零收益；
#   · 热控点表里 QA/QB/QC 被设计院拿去标**电磁阀**（`01LCE21QA001`，附录B 的 QA 实例首项
#     正是「电磁阀」，用法正确）和**流量喷嘴**（`01LAB70QB501`，QB 在附录B 是"电能电路隔离"）
#     → 套表5.2.7 会产出 82 条**误导性**提示（对着喷嘴讲"母线侧隔离开关号段"）。
# 结论：**不查**。要对"分类码与设备名不符"发话，得走 名称↔设备类 的匹配，那是另一条规则，
#       不能借号段表来说事（借了就是把"分类码用错"误报成"编号段用错"）。
# QA/QB/QC 只在 `_check_device_class` 里校验**是否收录、是否属剔除段**。

# 物理量码 → 中文（附录B 表B.0.2；总表 §8.4）
# ⚠️ `BB` 原表 OCR 漏行（BA→BC 之间），**不凭推定补**，故不列入
PHYSICAL_QUANTITY: dict[str, str] = {
    "BA": "电压/电位", "BC": "电流", "BD": "密度", "BE": "其他电气或电磁变量（如频率）",
    "BF": "流量", "BG": "位置/长度/振幅", "BJ": "功率", "BK": "时间", "BL": "液位（料位）",
    "BM": "水分/湿度", "BP": "压力/真空", "BQ": "成分/浓度/纯度", "BR": "辐射",
    "BS": "速度/机械频率/加速度", "BT": "温度", "BU": "多变量", "BW": "重量/力",
    "BX": "其他的量", "BZ": "事件数/计数",
}


@dataclass
class Finding:
    """一条校验结论。"""
    level: str                  # error / warn / pending
    segment: str                # plant / syscls / sysno / devcls / devno / …
    message: str                # 业务可读
    basis: str = ""             # 出处（条款号 / 总表规则号）
    suggestion: str = ""        # 建议改法（可执行的才填）

    def __str__(self) -> str:   # pragma: no cover - 便于调试
        return f"[{self.level}] {self.segment}: {self.message}（{self.basis}）"


@dataclass
class Parsed:
    """段位解析结果。"""
    kind: str                            # function/system/install/position/signal/terminal/document/union/product/unknown
    segments: dict[str, str] = field(default_factory=dict)
    findings: list[Finding] = field(default_factory=list)
    raw: str = ""

    @property
    def errors(self) -> list[Finding]:
        return [f for f in self.findings if f.level == "error"]

    @property
    def pendings(self) -> list[Finding]:
        return [f for f in self.findings if f.level == "pending"]

    @property
    def warnings(self) -> list[Finding]:
        return [f for f in self.findings if f.level == "warn"]

    @property
    def has_error(self) -> bool:
        return any(f.level == "error" for f in self.findings)


# ---------------------------------------------------------------- 判据查询
def system_known(code: str) -> str:
    """系统分类码收录情况：full（3 位命中）/ prefix（仅前 2 位命中）/ excluded（中煤剔除的核电码）/ "" 。"""
    code = str(code or "").upper()
    if not code:
        return ""
    if code in EXCLUDED:
        return "excluded"
    if code in SYS_DICT:
        return "full"
    # 中煤细则以 2 位定义系统码，第 3 位由项目约定（总表 §12.3）→ 前两位命中只作待核
    if len(code) >= 2 and code[:2] in SYS_DICT:
        return "prefix"
    return ""


def device_known(code: str) -> bool:
    return bool(code) and str(code).upper() in DEV_DICT


def band_of(sysno_or_devno: str) -> tuple[str, str] | None:
    """测点设备编号落在哪个号段 → (用途, level)。"""
    digits = re.sub(r"\D", "", sysno_or_devno or "")
    if not digits:
        return None
    n = int(digits)
    for lo, hi, use, level in MEASURE_BANDS:
        if lo <= n <= hi:
            return use, level
    return None


# ---------------------------------------------------------------- 段位校验
def _check_plant(seg: str, out: list[Finding]) -> None:
    """全厂码 H₂H₃（表4.3.1-2；总表 §3 / `4-05`）。"""
    n = int(seg)
    if 1 <= n <= 60:
        return                                   # 第 N 号机组
    if 61 <= n <= 90:
        return                                   # 两机共用
    if 91 <= n <= 99:
        return                                   # 三机及以上共用
    if n == 0:
        out.append(Finding("warn", "plant", "全厂码 00 = 全厂公用；请确认本行确为全厂公用设备",
                           "表4.3.1-2", "非公用设备应填机组号"))
        return
    out.append(Finding("error", "plant", "全厂码超出 00~99 范围", "表4.3.1-2"))


def _check_system_class(seg: str, out: list[Finding]) -> None:
    """系统分类码 S₁S₂S₃（附录A；总表 `4-07`、§12.3）。"""
    st = system_known(seg)
    if st == "full":
        return
    if st == "prefix":
        # 中煤按 2 位定义；仅 2 位命中属**待核**，不得判错（总表 §10.13）
        out.append(Finding("pending", "syscls",
                           f"系统分类码 {seg} 仅前 2 位 {seg[:2]} 在码表内，第 3 位由工程约定",
                           "总表 §12.3 / 中煤细则", "核对《工程约定与编码索引》"))
        return
    if st == "excluded":
        out.append(Finding("error", "syscls",
                           f"系统分类码 {seg} 属中煤剔除的核电码（J 核产热 / K 反应堆辅助）",
                           "中煤细则 T25", "核电码在本项目不可用"))
        return
    out.append(Finding("pending", "syscls",
                       f"系统分类码 {seg} 未收录（国标附录A 与中煤细则均无）",
                       "总表 §10.13", "**未收录只作待核线索，不判为错误**；核对编码索引"))


def _check_device_class(seg: str, out: list[Finding]) -> None:
    """设备分类码 T₁T₂（附录B；总表 `4-09`）。"""
    if device_known(seg):
        return
    out.append(Finding("pending", "devcls",
                       f"设备分类码 {seg} 未收录（国标附录B 与中煤细则均无）",
                       "总表 §10.13", "未收录只作待核线索；核对编码索引"))


def _check_device_numbering(devcls: str, devno: str, out: list[Finding]) -> None:
    """设备编号 T₃T₄T₅ 的**用途号段**校验（测点 + 阀门管道 + 开关类）。

    三套号段各不相同，判据也不同：

    | 对象 | 号段来源 | 表外段位怎么判 |
    |---|---|---|
    | 测点（T₁=B） | 编制说明 5.2.7 表5 | 有明确 "未用/锁定" 语义 → `pending`/`error` |
    | 阀门 QM / 管道 WP | **表5.1.5** | 501~599、701~999 **原表未列** → `pending`（不判错） |
    | 开关 QA/QB/QC | **查不了**（见下方"为什么不查表5.2.7"） | — |

    ⚠️ **实测（热控点表 58,280 行）**：QM/WP 行 13,377 条里 10,734 在 001~299（通用）、
      2,627 在 401~499（疏水和排气用）——**本来就合表5.1.5**，不能误报；
      真正落在表外段位的只有 14 行（`01LAA10QM501` 除氧器启动排汽电动门），
      低噪声、可解释，故保留。
    """
    if not devcls:
        return
    if devcls[0] == "B":                      # 测点：另一套号段
        _check_measure_point(devcls, devno, out)
        return
    if devcls not in ("QM", "WP"):            # 只管阀门与管道（表5.1.5）
        return
    digits = re.sub(r"\D", "", devno or "")
    if not digits:
        return
    n = int(digits)
    if any(lo <= n <= hi for lo, hi, _use in QM_WP_BANDS):
        return
    listed = "、".join(f"{lo:03d}~{hi:03d}" for lo, hi, _u in QM_WP_BANDS)
    out.append(Finding("pending", "devno",
                       f"{devcls} 的设备编号 {devno} 不在表5.1.5 号段内（表内为 {listed}）",
                       "5.1.5 + 表5.1.5 / 总表 §8.3",
                       "原表未定义该段位含义，需人工确认后约定使用"))


def _check_measure_point(devcls: str, devno: str, out: list[Finding]) -> None:
    """测点专项（总表 §8.2~§8.4、§12.7）。

    **T₁ = B** 即"测量回路/测点"的主组码（5.2.3-1、7.1.5）——注意 T₁ 是 `T₁T₂` 的
    **首字母**，不是整段等于 `B`（写成 `devcls != "B"` 就永远判不到 `BC`/`BP`/`BT`）。
    """
    if not devcls or devcls[0] != "B":
        return
    # T₂ 应按所测物理量查附录B 表B.0.2（5.2.3-1）
    if devcls not in PHYSICAL_QUANTITY:
        out.append(Finding("warn", "devcls",
                           f"T₁=B 表示测量回路，但 T₂ 码 {devcls} 不在附录B 表B.0.2 物理量码内"
                           f"（表内为 {'/'.join(sorted(PHYSICAL_QUANTITY))}）",
                           "5.2.3-1 / 附录B 表B.0.2",
                           "确认是否为『系统专用信号处理元件』（那种情形 T₁ 应为 K）"))
    band = band_of(devno)
    if band is None:
        return
    use, level = band
    if level == "error":
        out.append(Finding("error", "devno",
                           f"测量回路编号 {devno} 落在「{use}」号段（701~899），项目编制说明锁定不可用",
                           "编制说明 5.2.7 表5", "改到合法号段：开关量 001~099、模拟量 101~199"))
    elif level == "pending":
        out.append(Finding("pending", "devno",
                           f"测量回路编号 {devno} 落在「{use}」号段（原表未定义用途）",
                           "编制说明 5.2.7 表5", "确认是否属约定使用"))


def _check_no_dot(seg: str, out: list[Finding], kind: str) -> None:
    """功能面 / 产品面 / 文档标识内部**不得有 `.`**（总表 §1.11、§12.2）。"""
    if "." in seg:
        out.append(Finding("error", kind,
                           "该标识内部不应出现分隔符 `.`（`.` 只用于功能分配、安装点、位置、信号、端子 5 类）",
                           "图4.4.1 / 总表 §12.2", "去掉 `.`"))


def parse(code: str) -> Parsed:
    """按前缀符判标识类型并逐段校验。**不做任何改写**。"""
    raw = str(code or "").strip()
    up = raw.upper().replace(" ", "")
    p = Parsed(kind="unknown", raw=up)
    if not up:
        p.findings.append(Finding("error", "code", "编码为空", "—"))
        return p

    if up.startswith("#"):
        p.kind = "union"
        if not RE_UNION.match(up):
            # 不是语法错就判 警告：图注只写"分隔符"，未写不得省略（总表 `4-02`）
            p.findings.append(Finding("warn", "union",
                                      "联合标识格式不符合 `#<C1..C12>.<U1..U12>`（各级 ≤12 位）",
                                      "图4.2.1 / 4.2.2"))
        return p

    if up.startswith("++"):
        p.kind = "position"
        m = RE_POSITION.match(up)
        if not m:
            p.findings.append(Finding("error", "position",
                                      "位置标识格式不符合 `++全厂码+建(构)筑物码+编号.分层/房间码`",
                                      "图4.6.6 / 4.6.9"))
            return p
        p.segments = {k: v for k, v in m.groupdict().items() if v}
        r12 = m.group("r12")
        # R₁R₂：室内模式 {01~09} ∪ {91~99}；区域模式为字母（4.6.9-2，总表 `4-37`/`4-39`）
        if r12.isdigit():
            if not (("01" <= r12 <= "09") or ("91" <= r12 <= "99")):
                p.findings.append(Finding("error", "r12",
                                          f"楼层编号 {r12} 非法：地上 1~9 层=01~09，地下负 1~9 层=91~99",
                                          "4.6.9-2 / 总表 `4-37`", "改到 01~09 或 91~99"))
        return p

    if up.startswith("-"):
        p.kind = "product"
        segs = up[1:].split("-")
        if not RE_PRODUCT_SEG.match(segs[0]):
            p.findings.append(Finding("error", "product", "产品面标识首段格式不符（P₁P₂(P₃)P₄P₅）",
                                      "图4.5.1"))
        return p

    if up.startswith(";"):
        p.kind = "signal"
        m = RE_SIGNAL.match(up)
        if not m:
            p.findings.append(Finding("error", "signal",
                                      "信号标识格式不符 `;X₁X₂(X₃)X₄X₅[.Y₁Y₂Y₃Y₄]`",
                                      "图4.7.1 / 5.2.11~5.2.13"))
            return p
        p.segments = {k: v for k, v in m.groupdict().items() if v}
        x2, x45 = m.group("x2"), m.group("x45")
        if x2 in ("B", "G", "H") and x45 not in X45_FORCED:
            p.findings.append(Finding("error", "x45",
                                      f"X₂={x2} 时信号编号 {x45} 不在表5.2.13 取值内",
                                      "5.2.13 / 表5.2.13"))
        return p

    if up.startswith(":"):
        p.kind = "terminal"
        if not RE_TERMINAL.match(up):
            # ⚠️ 端子字符集开放（可含 `/`），正文**无** I/O 禁用注；只按长度/级数判
            p.findings.append(Finding("pending", "terminal",
                                      "端子标识不符合 `:<第1级>.<第2级>`（各级 ≤12 位，字符可为字母/数字/`/`）",
                                      "4.8.1 注1 / 总表 `4-50`"))
        return p

    if up.startswith("&"):
        p.kind = "document"
        if not RE_DOCUMENT.match(up):
            p.findings.append(Finding("error", "document", "文档标识应为 `&<3字母分类码><3位编号>`",
                                      "4.9.1 / 总表 `4-56`"))
        return p

    if up.startswith("="):
        up = up[1:]          # 前缀符不影响段位判读
        if up.startswith("="):
            up = up[1:]
    if up.startswith("+"):
        p.kind = "install"
        m = RE_INSTALL.match(up)
        if not m:
            p.findings.append(Finding("error", "install",
                                      "安装点标识格式不符 `+全厂码+系统码[+设备码].安装空间码`（`.` 不得省略）",
                                      "图4.6.2 / 4.6.2 注2"))
            return p
        p.segments = {k: v for k, v in m.groupdict().items() if v}
        _check_plant(m.group("h23"), p.findings)
        _check_system_class(m.group("sys"), p.findings)
        if m.group("dev"):
            _check_device_class(m.group("dev"), p.findings)
        return p

    # 余下：功能面标识（含 12 位设备码 / 7 位系统码）
    m = RE_FUNCTION.match(up)
    if m:
        p.kind = "function"
        p.segments = {k: v for k, v in m.groupdict().items() if v}
        _check_no_dot(up, p.findings, "function")
        _check_plant(m.group("h23"), p.findings)
        _check_system_class(m.group("sys"), p.findings)
        _check_device_class(m.group("dev"), p.findings)
        _check_device_numbering(m.group("dev"), m.group("devno"), p.findings)
        return p

    m = RE_SYSTEM_ONLY.match(up)
    if m:
        p.kind = "system"
        p.segments = {k: v for k, v in m.groupdict().items() if v}
        _check_no_dot(up, p.findings, "system")
        _check_plant(m.group("h23"), p.findings)
        _check_system_class(m.group("sys"), p.findings)
        return p

    # ---------- 含 `.` 的两种真实形态（必须分开处理）----------
    if "." in up:
        head, _sep, tail = up.partition(".")
        hm = RE_FUNCTION.match(head)
        if hm:
            # ① `07BFA01UC001.A1`：前段是**完整功能面标识**，`.` 后是附加/产品标识。
            #    竣工版用这种写法；国标 4.5.5 要求功能面+产品面用 `-` 组合，且
            #    **功能面内部无 `.`**（总表 §1.11）。→ 报**警告**（不判错），
            #    因为这不是"功能面内部插点"，而是"两段之间用点"的工程习惯。
            p.kind = "function"
            p.segments = {k: v for k, v in hm.groupdict().items() if v}
            _check_plant(hm.group("h23"), p.findings)
            _check_system_class(hm.group("sys"), p.findings)
            _check_device_class(hm.group("dev"), p.findings)
            _check_device_numbering(hm.group("dev"), hm.group("devno"), p.findings)
            p.findings.append(Finding(
                "warn", "separator",
                f"功能面标识后接附加标识用了 `.`（`{head}.{tail}`）：国标 4.5.5 要求"
                f"功能面与产品面组合时用产品面前缀符 `-`，且功能面**内部**不得出现 `.`",
                "4.5.5 / 总表 §1.11", f"建议写作 `={head}-{tail}`；须核对工程约定，不能直接判国标合规"))
            return p
        if RE_FUNCTION.match(up.replace(".", "")) or RE_SYSTEM_ONLY.match(up.replace(".", "")):
            # ② `01HAC13.BP101`：把点去掉就是合规功能面标识 → **功能面内部插点，判错**
            p.kind = "function"
            p.findings.append(Finding(
                "error", "separator",
                f"功能面标识内部不得出现分隔符 `.`（`.` 只用于功能分配、安装点、位置、信号、端子）",
                "图4.4.1 / 总表 §12.2", f"改为 `{up.replace('.', '')}`"))
            return p

    # ---------- 字母尾缀（分相）：`07ABG01BA001A` = 12 位设备码 + 1 位分相 A/B/C ----------
    # 竣工版实测形态（`kks_gen/README.md` §2）：A/B/C 是分相，**不单独成节点**，
    # 父级挂在 7 位系统码上。这里按功能面标识接收，并报**提示级**（不是错误）。
    if len(up) in (13, 14):
        head, tail = up[:12], up[12:]
        hm = RE_FUNCTION.match(head)
        if hm and tail.isalpha():
            p.kind = "function"
            p.segments = {k: v for k, v in hm.groupdict().items() if v}
            p.segments["phase"] = tail
            _check_plant(hm.group("h23"), p.findings)
            _check_system_class(hm.group("sys"), p.findings)
            _check_device_class(hm.group("dev"), p.findings)
            _check_device_numbering(hm.group("dev"), hm.group("devno"), p.findings)
            p.findings.append(Finding(
                "warn", "phase",
                f"末 {len(tail)} 位 `{tail}` 按分相/附加标识处理（竣工版实测形态，A/B/C 表分相）",
                "竣工版实测 / README §2", "分相后缀不单独成节点，父级挂在 7 位系统码上"))
            return p

    # ---------- 宽松功能面：编号位数不足（少写前导零）→ 可明确修复 ----------
    ml = RE_FUNCTION_LENIENT.match(up)
    if ml:
        p.kind = "function"
        p.segments = {k: v for k, v in ml.groupdict().items() if v}
        _check_plant(ml.group("h23"), p.findings)
        _check_system_class(ml.group("sys"), p.findings)
        _check_device_class(ml.group("dev"), p.findings)
        _check_device_numbering(ml.group("dev"), ml.group("devno"), p.findings)
        short = []
        if len(ml.group("sysno")) < 2:
            short.append(f"系统编号 {ml.group('sysno')} 应补足 2 位（0 必须写出）")
        if len(ml.group("devno")) < 3:
            short.append(f"设备编号 {ml.group('devno')} 应补足 3 位（0 必须写出）")
        if short:
            p.findings.append(Finding(
                "warn", "numbers", "；".join(short), "4.4.4-2 / 4.4.5-2（总表 `4-08`/`4-10`）",
                "补前导零至规定位数"))
        return p

    # ---------- 5 位系统级码：全厂码2 + 系统分类码3，**缺系统编号 S₄S₅** ----------
    m5 = re.match(rf"^(?:[{H1_SET}])?(?P<h23>{DIGIT}{{2}})(?P<sys>{LETTER}{{3}})$", up)
    if m5:
        p.kind = "system"
        p.segments = {k: v for k, v in m5.groupdict().items() if v}
        _check_system_class(m5.group("sys"), p.findings)
        p.findings.append(Finding(
            "pending", "sysno",
            f"5 位系统级编码，缺系统编号 S₄S₅（2 位数字）",
            "图4.4.1 / 4.4.4-2", "补系统编号后即为合规 7 位系统码"))
        return p

    # ---------- 12 位设备码 + 部件/附加后缀（`61EAB10GL001-M01`、`64GHA10GP001BT301`）----------
    # 实测占未判定行的大头（16 位 851 行 + 17 位 67 行）。国标允许扩展/部件层级，
    # 故**识别并原样保留**后缀，只校验前 12 位；本引擎**不生成**部件（实测 0% 精确率）。
    for cut in (12, 13, 14):
        if len(up) <= cut:
            continue
        head, tail = up[:cut], up[cut:]
        hm = RE_FUNCTION.match(head)
        if hm and RE_PART_SUFFIX.match(tail):
            p.kind = "function"
            p.segments = {k: v for k, v in hm.groupdict().items() if v}
            p.segments["part"] = tail
            _check_plant(hm.group("h23"), p.findings)
            _check_system_class(hm.group("sys"), p.findings)
            _check_device_class(hm.group("dev"), p.findings)
            _check_device_numbering(hm.group("dev"), hm.group("devno"), p.findings)
            p.findings.append(Finding(
                "warn", "part",
                f"末段 `{tail}` 按**部件/扩展层级**处理（前 12 位为设备码）",
                "README §2/§5（竣工版实测形态）",
                "部件级编码本引擎不生成（实测精确率 0%），仅原样保留"))
            return p

    p.findings.append(Finding("pending", "code",
                              f"无法按 GB/T 50549-2020 的标识结构解析（{len(up)} 位，未命中任何标识类型模式）",
                              "图4.4.1 / 4.7.1", "人工确认标识类型与段位"))
    return p


def is_measure_point(code: str) -> bool:
    """是否是「测量回路/测点」（T₁ = B，总表 §8.2/§8.3、`12.7`）。"""
    p = parse(code)
    return p.kind in ("function", "install") and p.segments.get("dev", "").startswith("B")
