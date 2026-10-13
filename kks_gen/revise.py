# -*- coding: utf-8 -*-
"""KKS 编码修编引擎（依据 GB/T 50549-2020 + 中煤《应用细则》）。

目标：把「与业务描述不符 / 不符合国标」的原设计编码，修编成**符合国标且符合业务描述**的编码；
判不出来的一律**留空**并给出简短、可执行的说明（**绝不猜**）。

## 分层（本模块只负责"怎么改"，"合不合规"交给 identifiers）

| 层 | 模块 | 职责 |
|---|---|---|
| 校验层 | `identifiers.py` | 段位化校验（总表 §4 正则速查表 / §5 字符禁用 / §8 测点专项），输出 error/warn/pending |
| 修编层 | **本模块** | 依据校验结论**改**编码；没依据就留空 |

## 判据来源（逐条可回溯到 `项目编码规则总表.md`）

| 段 | 依据 | 数据/规则 |
|---|---|---|
| 结构 功能面 = 全厂码(2) + 系统码(3字母+2数字) + 设备码(2字母+3数字) = 12 位 | 图4.4.1 | 总表 §2 |
| 全厂码 01~60 第N号机组 / 61~90 两机共用 / 91~99 三机以上 / 00 公用 | 表4.3.1-2 | `gb_rules.unit_label_to_plant_code`、总表 §3 |
| 系统分类码 S₁S₂S₃ 合法取值 | 附录A | `project_system_codes.json`（3 412+） |
| 设备分类码 T₁T₂ 合法取值 | 附录B | `project_device_codes.json`（238） |
| 禁用字母 I、O —— **段位化**，只适用于 `A` 类型段 | 图4.4.1 注1 等 5 处图注 | 总表 §5.2、**§12.1/§12.4 硬伤** |
| 「每位上的 0 必须写出」→ 编号补前导零 | 4.4.4-2、4.4.5-2 | 总表 `4-05`/`4-08`/`4-10` |
| **测点（测量回路）**：T₁=B、T₂=物理量、T₃T₄T₅ 落表5 号段 | 5.2.3-1、7.1.5、编制说明 5.2.7 表5 | 总表 §8、**§12.7** |
| 功能性双标识（工艺系统的电气仪控设备：功能面 + 安装点） | 5.2.1 | 总表 §7.1、**§12.6** |
| **「码未收录」只作待核线索，不判为错误** | `kks_code_index.json` `coverage_notes` | 总表 §10.13、**§12.8** |

## 输出状态

- `ok`          已符合国标（原码保留）
- `revised`     已修编（有国标依据，`revised` 是改后的码）
- `not_device`  **指令/测点行：不编码**（`revised` 为空，`device_body` 给出所服务的设备本体码）
- `unresolved`  **留空** + 说明（缺判定依据，交人工）

另有两个**正交**的标记（不改变 `status` 语义，供报表用）：

- `needs_review`  True = 需人工核对（含**待核**：码未收录、号段未定义、原文未规定处）
- `level`         本次发现的最重级别：`error` / `pending` / `warn` / `""`

## 🔴 两条踩过坑后固化的纪律（勿改）

1. **系统分类码绝不靠"设备名称关键词"去猜**。实测 656 行靠设备名改系统码的，
   证据强度**全部为零**（`01HNA10BP102` 描述「末级再热器出口烟气压力」——"再热器"只是测点所在部位）。
   → 系统码只在**「系统」列有证据**时才自动改，否则留空交人工（总表 §12.16）。
2. **设备分类码可以按描述匹配**——描述讲的就是被测物理量（「输出电流」→ 附录B「BC 电流」），
   这是直接语义证据（总表 §8.4）。
"""
from __future__ import annotations

import pathlib
import re
import sys
from collections import Counter
from dataclasses import dataclass, field

if __package__ in (None, ""):
    sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent))

if __package__ in (None, ""):
    import identifiers as ids
    from gb_rules import ambiguous_public_label, unit_label_to_plant_code
else:
    from . import identifiers as ids
    from .gb_rules import ambiguous_public_label, unit_label_to_plant_code

FUNCTION_KINDS = ("function", "system")


@dataclass
class Revision:
    """一条编码的修编结果。"""
    original: str
    revised: str = ""
    status: str = "unresolved"        # ok / revised / unresolved / not_device
    note: str = ""                    # 简短业务说明（用户直接看这列）
    details: list[str] = field(default_factory=list)
    level: str = ""                   # error / pending / warn / ""
    device_body: str = ""             # status=not_device 时：本行指令/测点所服务的**设备本体码**
    _findings: list[ids.Finding] = field(default_factory=list, repr=False)

    @property
    def changed(self) -> bool:
        return self.status == "revised"

    @property
    def findings(self) -> list[ids.Finding]:
        return self._findings

    @property
    def needs_review(self) -> bool:
        """需人工核对：未判定（unresolved）或含**待核**项（码未收录等，总表 §10.13）。

        用**动态属性**而不是字段：`status` 与 `_findings` 都在后面才定稿，
        写成字段会在 status 还是默认值 `unresolved` 时就把"已符合国标"也标成需人核（踩过）。
        """
        return self.status == "unresolved" or any(f.level == "pending" for f in self._findings)

    def adopt(self, parsed: ids.Parsed) -> None:
        """吸收校验结论（只记 findings 与最重级别；needs_review 见上）。"""
        self._findings = list(parsed.findings)
        for lv in ("error", "pending", "warn"):
            if any(f.level == lv for f in parsed.findings):
                self.level = lv
                break

    def fail(self, parsed: ids.Parsed, note: str) -> "Revision":
        """判定为**无法自动修编**：留空 + 说明。"""
        self.adopt(parsed)
        self.revised = ""
        self.status = "unresolved"
        self.note = note
        return self

    def keep(self, parsed: ids.Parsed, note: str) -> "Revision":
        """保留原码（不改动），但把结论写进说明。

        用于**待核**类：码未收录、号段原表未定义等 —— 总表 §10.13 要求"只作待核线索，
        不直接判定为错误"，所以既不能判错、也不该悄悄当"已符合"。
        """
        self.adopt(parsed)
        self.revised = self.original
        self.status = "ok"
        self.note = note
        return self


# ------------------------------------------------------------------ 语义选码
def _keywords(text: str) -> list[str]:
    """把中文释义/描述切成可用于匹配的词。"""
    text = re.sub(r"[（）()【】\[\]、，,；;：:·\-/]", " ", str(text or ""))
    return [w for w in text.split() if len(w) >= 2]


def _pick(table: dict, description: str, *, lengths: tuple[int, ...], limit: int) -> list[str]:
    if not description:
        return []
    desc = str(description)
    scored: list[tuple[int, str]] = []
    for code, item in table.items():
        if len(code) not in lengths:
            continue
        name = str(item.get("cn") or "")
        if not name:
            continue
        hits = [w for w in _keywords(name) if w in desc]
        if not hits:
            continue
        score = sum(len(w) for w in hits) * 10 + len(hits)   # 匹配词越长越可信
        scored.append((score, code))
    scored.sort(key=lambda x: (-x[0], x[1]))
    return [code for score, code in scored[:limit] if score == scored[0][0]] if scored else []


def pick_device_code(description: str, *, limit: int = 3) -> list[str]:
    """按业务描述从附录B 选设备/信号分类码（返回候选，按匹配度排序）。

    依据：附录B 的中文释义（如 `BC 电流`、`BP 压力、真空`、`BT 温度`）。
    """
    return _pick(ids.DEV_DICT, description, lengths=(2,), limit=limit)


def pick_system_code(description: str, *, limit: int = 3) -> list[str]:
    """按业务描述从附录A 选系统分类码（3 位）。"""
    return _pick(ids.SYS_DICT, description, lengths=(3,), limit=limit)


def pick_quantity_code(description: str, *, limit: int = 3) -> list[str]:
    """测点的 T₂：只从**附录B 表B.0.2 物理量码**里选（总表 §8.4）。

    与 `pick_device_code` 的区别：这里限定 `B*` 物理量闭集，
    避免把「控制阀 QN」这类**非测量**的设备码选成测点的 T₂。
    """
    table = {c: {"cn": cn} for c, cn in ids.PHYSICAL_QUANTITY.items()}
    return _pick(table, description, lengths=(2,), limit=limit)


def _pad_number(value: str, width: int) -> tuple[str, bool]:
    """编号补前导零（国标"每位上的 0 必须写出"）。"""
    digits = re.sub(r"\D", "", value or "")
    if not digits:
        return "", False
    padded = digits.zfill(width)
    return (padded[:width], padded != digits or len(digits) != width)


def _fmt_table(**seg: str) -> str:
    """按段位拼回功能面标识。

    ⚠️ 必须带上 `phase`（分相后缀，如 `07ABG01BA001A` 的末位 `A`）与 `part`
    （部件/扩展后缀，如 `61EAB10GL001-M01` 的 `-M01`），否则修编会**把合法后缀丢掉**
    （踩过：`07ABG01BA001A` 被改成 `07ABG01BA001`）。
    """
    return "".join(v for v in (seg.get("h1", ""), seg["h23"], seg.get("h4", ""),
                               seg["sys"], seg.get("sysno", ""), seg.get("dev", ""),
                               seg.get("devno", ""))
                   ) + seg.get("phase", "") + seg.get("part", "")


def _suffix_of(original: str) -> str:
    """取原码中 `.` 之后的附加标识（含 `.`），用于修编后原样保留。

    `07BFA01UC001.A1` 的 `.A1` 是竣工版的附加标识写法——国标 4.5.5 建议改用产品面前缀符 `-`，
    但**不能因此把它删掉**（删了就是丢信息）。parse 已对此报 warn（不判错）。
    """
    i = original.find(".")
    return original[i:] if i >= 0 else ""


# ------------------------------------------------------------------ 主流程
# 信号标识后缀（设计院把**信号码**直接拼在设备码后面的写法）
# 依据：表5.2.11 X₁ ∈ {X=信号源, Y=信号应用}；表5.2.12 X₂ 可用 15 字母；5.2.13 X₄X₅ 两位编号。
#   `XB21` = 信号源（测点）；`YB21` = 信号应用（指令）。
# ⚠️ **只认 `X₁ ∈ {X,Y}` 的形式**（有明确国标依据）。`RM`/`ZF`/`VO`/`MS` 这类首字母不是 X/Y 的
#    一律**不认**（它们不是信号码），保持原样 —— 宁可不判，也不乱判。
#    实测：本表 22,646 行带后缀，其中符合该形式的 16,672 行。
_SIGNAL_IN_TAIL = re.compile(rf"[XY][{ids.X2_SET}]\d{{2,3}}$")


def signal_body(code: str) -> str:
    """去掉**信号标识后缀**，返回设备本体码（无该后缀时原样返回）。

    用于判定"一个原码属于指令/测点（`设备码 + 信号码`）还是设备本体本身"
    —— 见 `revise_code` 的 `not_device` 口径。

    ⚠️ 两道安全守卫，避免误伤"末尾碰巧长得像信号码"的普通码：
        ① 只有长度 > 12 才可能有后缀（12 位设备码是固定的）；
        ② 剥离后必须**正好 12 位**，否则不剥。
        （否则 `01ABC12XY123` 这种末尾是 `XY123` 的 12 位码会被误剥成 `01ABC12`。）
        `64GHA10GP001BT301`（`BT` 不是信号码）因此也原样返回。
    """
    from kks_gen.row_policy import classify_row
    value = str(code or "").strip()
    role = classify_row(value)
    return role.body if role.is_point and role.body and role.body != value else value


def _structurally_ok(code: str) -> bool:
    """**国标结构闸门**：只有这些形态的编码才允许写进「KKS编码（修编后）」列。

    依据 GB/T 50549-2020：4.4.1（功能面标识 = F0+F1+F2）、4.4.2/4.4.3（段位与字符集）、
    3.2.x（可按阶段标识到级，故允许层级节点码）。

    允许的形态：
      · 完整功能面标识（12 位；含 H₁/H₄ 时 13/14 位）或 7 位系统码 —— `identifiers.parse` 认可且无错
      · **层级节点码**（2/3/5 位：全厂码 + 0~3 字母 + 0~2 数字）—— 树的分层节点，竣工版同样使用
      · 12 位设备码 + **部件/附加后缀**（`FF01`、`BC`、`FT1`…）—— 国标允许扩展层级

    🔴 为什么要有这道闸门：修编前的实现会把**源表里结构畸形的码原样写进交付文件**
    （实测 `016`、`0161BHE01ABC`、`00BE02QA03`、`00ABE02QA021FT1` 等 10,943 行）。
    要求是"**生成时就对**"——不合规的值根本不该出现在交付文件里，而不是事后检查出来。
    """
    from kks_gen.compliance import assess
    return assess(code).writable



def revise_code(raw: str, *, unit_label: str = "", description: str = "",
                device_type: str = "", system_name: str = "") -> Revision:
    """修编一个编码（**带国标结构闸门**，见 `_structurally_ok`）。"""
    result = _revise_code_inner(raw, unit_label=unit_label, description=description,
                               device_type=device_type, system_name=system_name)
    # ---------- 指令/测点行：**不编码**（业务口径，2026-10-12 用户定调）----------
    # 设计院把**信号码**直接拼在设备码后面（`01HLB10GQ001` + `XB21`）：
    #   `XB**` = 信号标识 X₁=X（信号源）→ 测点：已运行/已停止/开到位/频率反馈…
    #   `YB**` = 信号标识 X₁=Y（信号应用）→ 指令：启动/停止/频率给定/开度指令…
    # 口径：**这些是指令和测点，发出指令/信号的那个才是设备，只有设备才有 KKS 编码**
    #   → 本行「KKS编码（修编后）」**留空**，「修编说明」写明它所服务的设备本体码
    #     （设备本体码由**设备本体行**给出，不在这类行上重复）。
    # ⚠️ 为什么撤销旧的"归一到设备本体码"：归一后 12 个不同测点全落成同一个
    #    `01HLB10GQ001`，一张**本来无重复**的表（输入 58,276 码 / 唯一 58,274）
    #    被做成 **3,563 类、18,088 行重复** —— 编码列重复 = 交付物报废。
    # ⚠️ 为什么判在**最外层**而不是段位处理里：`_revise_code_inner` 有多条**早退路径**
    #    （系统码/设备分类码"未收录 → 只作待核、保留原码"、号段违规…），它们在到达段位
    #    处理之前就返回了。判在里面实测漏掉 **425 行**（如 `01MAV51AH001XB21`，`AH` 未收录）。
    # ⚠️ `signal_body` 自带两道守卫（长度 > 12 且剥后**正好 12 位**），
    #    所以 `64GHA10GP001BT301`（`BT` 不是信号码）这类不会被误判。
    from kks_gen.row_policy import classify_row
    role = classify_row(result.original, description, device_type, include_measurement=False)
    if role.is_point and not (result.original.startswith(";") and result.status == "unresolved"):
        result.device_body = role.body
        result.revised = ""
        result.status = "not_device"
        result.note = "本行是指令/测点，设备KKS编码留空（项目设备表口径）；" + role.reason
        result.note += (f"；其所服务的设备本体码为 {role.body}" if role.body else "；所服务设备须结合图纸确认，不猜本体码")
        result.details = [result.note]
        return result
    # ---------- 闸门：不合规的码**绝不写进交付文件** ----------
    if result.status in ("ok", "revised") and result.revised and not _structurally_ok(result.revised):
        original = result.original
        result.revised = ""
        result.status = "unresolved"
        result.level = result.level or "error"
        result.note = (f"编码 {original} 不符合国标标识结构（4.4.1：功能面标识 = 全厂码+系统码+设备码），"
                       f"且无法自动修编 → 留空待人工（不写入不合规的值）")
        result.details = list(result.details) + [result.note]
    if result.revised:
        from kks_gen.compliance import assess
        for message in assess(result.revised).pending:
            if not any(f.message == message for f in result._findings):
                result._findings.append(ids.Finding("pending", "agreement", message, "GB/T 50549-2020"))
                result.note += "；待核：" + message
    result.note = result.note.replace("已符合国标", "已通过自动结构校验")
    return result


def _revise_code_inner(raw: str, *, unit_label: str = "", description: str = "",
                       device_type: str = "", system_name: str = "") -> Revision:
    """修编一个编码。返回 `Revision`（status = ok / revised / unresolved）。"""
    original = str(raw or "").strip().upper().replace(" ", "")
    result = Revision(original=original)
    if not original:
        result.note = "原编码为空"
        return result          # status 默认 unresolved → needs_review 自动为 True

    parsed = ids.parse(original)

    # ---------- ① 无法解析结构 → 先尝试**确定性修复**，修不了才交人工 ----------
    # 要求是"**第一次就生成对**"：能在不猜的前提下修出来的，必须修出来，而不是一律留空。
    if parsed.kind == "unknown":
        plant_want = unit_label_to_plant_code(unit_label)
        # (a) **缺全厂码**：原码只有系统段（如 `CMD`、`ABC`），前面补机组号对应的全厂码
        #     —— 唯一确定、不猜（表4.3.1-2）。
        if re.match(r"^[A-HJ-NP-Z]", original) and len(original) + 2 in (5, 7, 12):
            if plant_want:
                cand = plant_want + original
                # ⚠️ 补完后必须是**完整标识**（12 位设备码 / 7 位系统码），**不能只是"层级节点形状"**：
                #    否则 `PQ1` 这种设计院位号会被误判成"缺全厂码"补成 `01PQ1`（实测踩过，测试变红）。
                cp = ids.parse(cand)
                if cp.kind in FUNCTION_KINDS and not cp.has_error:
                    result.revised = cand
                    result.status = "revised"
                    result.note = (f"原码缺 2 位全厂码 → 补 {plant_want}"
                                   f"（表4.3.1-2：{unit_label or '公用'}），得 {cand}")
                    result.details = [result.note]
                    return result
            else:
                why = ("本行「机组号」为空" if not (unit_label or "").strip()
                       else f"本行「机组号」只写「{unit_label}」，"
                            f"无法区分全厂公用（00）与两机共用（61~90）")
                result.note = (f"原码 {original} 缺 2 位全厂码，但{why}"
                               f" → 待人工（补上明确的机组号后可自动生成）")
                return result
        # (b) 修不了：说明**缺什么**（可执行），而不是只报"结构不符"
        missing = []
        if not (unit_label or "").strip() or not plant_want:
            missing.append("机组号（定全厂码）")
        if not (system_name or "").strip():
            missing.append("系统（定系统码）")
        hint = ("；需补充：" + "、".join(missing)) if missing else "；源码不是 KKS 码，无法在不编造的前提下还原"
        result.adopt(parsed)
        result.note = _diagnose_length(original) + hint
        return result

    # ---------- ② 非功能面/系统码（信号、端子、文档、位置、安装点、联合）----------
    #     本项目的修编范围是「有设计编码行」的功能面编码（业务口径），这些只做校验说明。
    if parsed.kind not in FUNCTION_KINDS:
        result.adopt(parsed)
        if parsed.has_error:
            bad = next(f for f in parsed.findings if f.level == "error")
            return result.fail(parsed, f"{bad.message}（{bad.basis}）"
                                      + (f"；建议：{bad.suggestion}" if bad.suggestion else ""))
        result.revised = original
        result.status = "ok"
        result.note = f"已符合国标（{_KIND_LABEL.get(parsed.kind, parsed.kind)}）"
        if parsed.pendings:
            result.note += f"；待核：{parsed.pendings[0].message}"
        return result

    seg = dict(parsed.segments)

    # ---------- ③ 结构性判错（如功能面内部插了 `.`）----------
    #    这类码**解析不出段位**，必须先处理，否则后面取 seg["sys"] 会 KeyError（踩过）。
    if parsed.has_error:
        bad = next(f for f in parsed.findings if f.level == "error")
        # 唯一能自动修的情形：功能面内部多了一个 `.` —— 国标依据明确，去掉即可（总表 §12.2）
        if bad.segment == "separator":
            fixed = original.replace(".", "")
            if ids.parse(fixed).kind in FUNCTION_KINDS and not ids.parse(fixed).has_error:
                result.adopt(parsed)
                result.revised = fixed
                result.status = "revised"
                result.note = "功能面标识内部不应出现分隔符 `.`（`.` 只用于功能分配、安装点、位置、信号、端子）→ 已去掉"
                return result
        return result.fail(parsed, f"{bad.message}（{bad.basis}）"
                                 + (f"；建议：{bad.suggestion}" if bad.suggestion else ""))

    if not seg.get("sys") or not seg.get("h23"):
        return result.fail(parsed, f"编码段位不完整（{parsed.kind}），无法判定，待人工")

    plant_want = unit_label_to_plant_code(unit_label)
    semantic = description or device_type
    # 系统分类码的**唯一合法证据**：系统列
    sys_evidence = pick_system_code(system_name) if system_name else []
    notes: list[str] = []

    if len(sys_evidence) > 1:
        return result.fail(parsed, "系统列对应多个分类码候选，不能自动选择：" + "、".join(sys_evidence))

    # ---------- ④ 全厂码：与机组号列不符 → 按表4.3.1-2 纠正 ----------
    if plant_want and seg["h23"] != plant_want:
        return result.fail(parsed, f"编码全厂码 {seg['h23']} 与机组列 {unit_label} 冲突；须依据图纸和工程约定裁定归属，不自动改为 {plant_want}")
    elif ambiguous_public_label(unit_label) and seg["h23"] not in ("", "00"):
        # 机组号列只写「公用」→ **判不出**是 00（全厂公用）还是 61（1、2号机共用）。
        # 🔴 实测（热控点表）：`61LAC10`（电动给水泵）**同一系统**下两种写法并存
        #    （一部分「公用」、一部分「1、2号机组共用」）→ 裸「公用」就是"共用"的简称。
        #    原来据此把 `61` 改成 `00`，等于把"1、2号机共用"改成"全厂公用"，
        #    还把同一系统的设备劈成两个分支（实测 127 行）——**已改为保留原码 + 提请人工**。
        notes.append(f"机组号列仅写「{unit_label}」：表4.3.1-2 里 00=全厂公用、61~90=两机共用，"
                     f"二者含义不同 → 保留原码 {seg['h23']}，需人工确认是否需要改为 00")

    # ---------- ⑤ 系统分类码 ----------
    sys_state = ids.system_known(seg["sys"])
    if sys_state == "excluded":
        # 中煤剔除的核电码：在中煤体系下非法
        if sys_evidence:
            notes.append(f"系统分类码 {seg['sys']}→{sys_evidence[0]}"
                         f"（属中煤剔除的核电码；依据「系统」列）")
            seg["sys"] = sys_evidence[0]
        else:
            return result.fail(parsed,
                               f"系统分类码 {seg['sys']} 属中煤剔除的核电码（J 核产热 / K 反应堆辅助，"
                               f"细则 T25）；「系统」列为空，无法判定应属哪个系统，待人工")
    elif sys_state == "":
        # **未收录**：只作待核线索，不判错（总表 §10.13）。有系统列证据才改。
        if sys_evidence:
            notes.append(f"系统分类码 {seg['sys']}→{sys_evidence[0]}"
                         f"（{ids.SYS_DICT[sys_evidence[0]]['cn'][:16]}，依据「系统」列）")
            seg["sys"] = sys_evidence[0]
        else:
            return result.keep(parsed,
                               f"系统分类码 {seg['sys']} 未收录（国标附录A 与中煤细则均无）——"
                               f"未收录只作待核线索，未改动；「系统」列为空，无法判定应属哪个系统")

    # ---------- ⑥ 设备分类码（含测点 T₂，总表 §8.4）----------
    #     ⚠️ 7 位系统码（`07BFA01`）是**合规的系统级标识**，本身就没有设备分类码——
    #        不能要求它"补设备码"（踩过：把 132 行合规系统码判成了"缺设备分类码"）。
    if parsed.kind == "function":
        dev_is_measure = seg.get("dev", "")[:1] == "B"
        if not ids.device_known(seg.get("dev", "")):
            if not seg.get("dev"):
                return result.fail(parsed, "缺设备分类码（2 字母，附录B），无法判定，待人工")
            # 测点优先按**物理量闭集**选 T₂，其它再按附录B 全表选
            cands = pick_quantity_code(semantic) if dev_is_measure else []
            if not cands:
                cands = pick_device_code(semantic)
            if len(cands) > 1:
                return result.fail(parsed, "设备描述对应多个同分分类码候选，须人工确认：" + "、".join(cands))
            if cands:
                notes.append(f"设备分类码 {seg['dev']}→{cands[0]}"
                             f"（附录B：{ids.DEV_DICT.get(cands[0], {}).get('cn', ids.PHYSICAL_QUANTITY.get(cands[0], ''))[:14]}）")
                seg["dev"] = cands[0]
            else:
                return result.keep(parsed,
                                   f"设备分类码 {seg['dev']} 未收录（国标附录B 与中煤细则均无），"
                                   f"且描述中找不到可判定的设备类别——待核，未改动")

        # ---------- ⑦ 测点：T₁=B 时校验 T₃T₄T₅ 号段（编制说明 5.2.7 表5）----------
        if seg.get("dev", "")[:1] == "B" and seg.get("devno"):
            band = ids.band_of(seg["devno"])
            if band and band[1] == "error":
                return result.fail(parsed,
                                   f"测量回路编号 {seg['devno']} 落在「{band[0]}」号段"
                                   f"（701~899 国标锁定不可用）；开关量测点应落 001~099、"
                                   f"模拟量测点应落 101~199，待人工确认后改号")

    # ---------- ⑧ 编号补前导零 ----------
    if seg.get("sysno"):
        new_sysno, padded = _pad_number(seg["sysno"], 2)
        if padded:
            notes.append(f"系统编号 {seg['sysno']}→{new_sysno}（补足 2 位）")
        seg["sysno"] = new_sysno
    if seg.get("devno"):
        new_devno, padded = _pad_number(seg["devno"], 3)
        if padded:
            notes.append(f"设备编号 {seg['devno']}→{new_devno}（补足 3 位）")
        seg["devno"] = new_devno

    revised = _fmt_table(**seg) + _suffix_of(original)
    result.adopt(parsed)
    result.revised = revised
    if revised == original:
        result.status = "ok"
        # ⚠️ 原码不变**不代表没有话说**：④ 的「机组号只写公用、无法区分 00/61」这类
        #    **提请人工**的说明必须带出去 —— 原来这里直接覆盖成"已符合国标"，
        #    把说明吞掉了（实测：127 行歧义提示全丢）。
        #    交付时 `gen_step1` 会把"已符合国标"这类**状态回显**过滤掉，只留真内容。
        result.note = "；".join(["已符合国标", *notes]) if notes else "已符合国标"
        pend = parsed.pendings
        if pend:
            result.note += f"；待核：{pend[0].message}"
    else:
        result.status = "revised"
        result.note = "；".join(notes) if notes else "已按国标调整"
        if _suffix_of(original) and not notes:
            result.note = "原码合规，仅保留附加标识；结构已按国标段位重排"
    result.details = notes
    return result


_KIND_LABEL = {
    "signal": "信号标识", "terminal": "端子标识", "document": "文档标识",
    "position": "位置标识", "install": "安装点标识", "union": "联合标识",
    "product": "产品面标识",
}


def _diagnose_length(code: str) -> str:
    """按国标结构给出简短的、可执行的诊断。"""
    n = len(code)
    if n < 2:
        return "编码过短：全厂码就需 2 位数字（1号机组 01、2号机组 02、公用 00）"
    if n == 3:
        return "仅 3 位：应是「全厂码2位 + 系统分类码第1位」，缺后续字符"
    if n == 5:
        return "仅 5 位：应是「全厂码2位 + 系统分类码3位」，缺系统编号(2位数字)与设备码"
    if n == 9:
        return ("9 位：全厂码2 + 系统码5 + 设备分类码2，**缺 3 位设备编号**"
                "（国标 001~999，不足补前导零）")
    if n == 10:
        return ("10 位：比标准 12 位少 2 位，**通常缺开头 2 位全厂码**"
                "（1号机组 01、2号机组 02、1、2号共用 61、公用 00）")
    if n == 11:
        return "11 位：比标准 12 位少 1 位，检查是否漏了 1 位编号（国标编号位数固定）"
    if n == 13:
        return ("13 位：若末位是 A/B/C 为分相后缀（合规）；否则是 12 位设备码 + 1 位附加，"
                "需确认附加位含义")
    if n in (14, 15, 16, 17, 18):
        if ids.parse(code).kind == "function":
            return (f"{n} 位：前段是合规设备码，后 {n - 12} 位是附加标识——"
                    "国标 4.5.5 要求功能面与产品面组合时用产品面前缀符 `-`（**功能面内部不得加点**）")
        return f"{n} 位：前段不是合规设备码，无法判定结构，待人工"
    return f"{n} 位：不符合国标三级编码结构（全厂码2 + 系统码5 + 设备码5 = 12），待人工"


def revise_rows(rows: list[dict]) -> list[Revision]:
    """批量修编。rows 每项需含 code / unit_label / description / device_type。"""
    return [revise_code(r.get("code", ""), unit_label=r.get("unit_label", ""),
                        description=r.get("description", ""),
                        device_type=r.get("device_type", ""),
                        system_name=r.get("system_name", ""))
            for r in rows]


def revise_summary(revisions: list[Revision]) -> dict[str, int]:
    """统计（供 CLI / 网页端进度用），含**待核**与**需人核**。"""
    c = Counter(r.status for r in revisions)
    return {
        "ok": c["ok"], "revised": c["revised"], "unresolved": c["unresolved"],
        "needs_review": sum(1 for r in revisions if r.needs_review),
        "pending_only": sum(1 for r in revisions
                            if r.status == "ok" and r.level == "pending"),
    }


# ------------------------------------------------------------------ 上下文推断机组号
def infer_unit_labels(rows: list[dict], *, window: int = 10,
                      min_share: float = 0.6) -> list[str]:
    """给「机组号为空」的行推断机组号。

    用户已确认可以采用（有依据但不是 100%，所以只在证据足够时才推）：

    证据优先级：
      ① **同一「图纸名称或图号」** 的其它行（同一张图基本同一机组）——强证据
      ② 同一「系统」的其它行（同一系统通常同机组）
      ③ 上下各 window 行的多数（同一区段连续录入）——次强证据

    只在多数占比 >= min_share 时才采纳；否则返回空串（保持"待人工"，**绝不猜**）。

    rows 每项需含：unit_label / drawing / system
    返回与 rows 等长的 list[str]。
    """
    units = [str(r.get("unit_label", "") or "").strip() for r in rows]
    drawings = [str(r.get("drawing", "") or "").strip() for r in rows]
    systems = [str(r.get("system", "") or "").strip() for r in rows]

    def majority(values: list[str]) -> str:
        clean = [v for v in values if v]
        if not clean:
            return ""
        top, count = Counter(clean).most_common(1)[0]
        return top if count / len(clean) >= min_share else ""

    by_drawing: dict[str, list[str]] = {}
    for unit, drawing in zip(units, drawings):
        if drawing and unit:
            by_drawing.setdefault(drawing, []).append(unit)

    by_system: dict[str, list[str]] = {}
    for unit, system in zip(units, systems):
        if system and unit:
            by_system.setdefault(system, []).append(unit)

    out: list[str] = []
    for i, unit in enumerate(units):
        if unit:
            out.append(unit)
            continue
        guess = majority(by_drawing.get(drawings[i], [])) if drawings[i] else ""
        if not guess and systems[i]:
            guess = majority(by_system.get(systems[i], []))
        if not guess:
            lo, hi = max(0, i - window), min(len(units), i + window + 1)
            guess = majority(units[lo:hi])
        out.append(guess)
    return out


def system_code_ok(code: str) -> bool:
    """系统分类码是否合法（按项目可用码表）。

    · 三位码本身在表内 → 合法
    · **前两位在表内** → 也合法（中煤细则以 2 位定义系统码，第 3 位由项目定）
    · 一位主组码在表内 → 合法（只到主组级的标识）
    · 中煤剔除的核电码（J/K）→ **不合法**
    """
    return ids.system_known(code) in ("full", "prefix")


def device_code_ok(code: str) -> bool:
    """设备分类码是否合法（按项目可用码表）。"""
    return ids.device_known(code)
