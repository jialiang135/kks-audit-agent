# -*- coding: utf-8 -*-
"""KKS 设备树生成：规则层（可复现、不调模型）。

输入：有设计编码的原始行（原名称 / 设计编码 / 专业 / 图纸名称或图号）
输出：完整设备树（设备名称 / KKS编码 / 上级设备名称 / 上级KKS编码）

设计原则（重要）：
1. **能确定的绝不用 AI**：层级截取、后缀拼接、机组号归一、括号半角化 —— 全部走规则，
   保证同一输入永远同一输出、可复现、可审计。
2. **AI 只补规则的盲区**：分类码语义选择、原名称为空时的名称反推。这些位置标 `需要AI`，
   由调用方决定是否送模型；本模块自己不做任何模型调用。
3. **每个值都带来源与置信度**，便于出"对照表"。

层级规则（GB/T 50549-2020 功能面标识，火电全厂码 H1 缺省）：
    G(1) | F0(1) | F1F2F3(3) | F4F5(2) | A1A2(2) | T3T4T5(3)
    位长： 1      2       5          7        9         12
    再加部件级扩展：设备码 + "-" + 部件分类码(2) + 编号(2~3)，如 07AAB02UH001-KF01
"""
from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass, field
from typing import Any, Iterable

# KKS 功能面层级的关键位长（用于向上展开父级链）
LEVEL_CUTS = (1, 2, 3, 5, 7, 9, 12)
LEVEL_NAMES = {
    1: "全厂码",
    2: "机组",
    3: "主组",
    5: "系统",
    7: "子系统",
    9: "设备组",
    12: "设备",
}

# ⚠️ 直接上级映射（**实测修正**）：设备分类码 A1A2（第8-9位）**不单独成节点**。
# 竣工版实测：07AAB02UH001（12位）的上级是 07AAB02（7位），不是 07AAB02UH（9位）。
# 所以 9 位与 12 位一律挂到 7 位子系统下；9 位节点只在它**自己就是设计编码**时才存在。
PARENT_LEN = {
    3: 2,
    5: 3,
    6: 5,
    7: 5,
    9: 7,
    12: 7,
}
# 带后缀的扩展码（如 07ABG01BA001A / .A1）：主干 + 尾缀。父级按主干长度映射。
EXT_LEN_MAP = {13: 12, 14: 13, 15: 12, 16: 12, 17: 12, 18: 12, 23: 12}
# 允许单独成节点的位长
NODE_LENS = (2, 3, 5, 7, 9, 12)


def _effective_len(base: str) -> int:
    """把带尾缀的扩展码折算成"主干位长"，用于父级映射。"""
    length = len(base)
    if length in PARENT_LEN:
        return length
    if length in EXT_LEN_MAP:
        return EXT_LEN_MAP[length]
    # 未知长度：按最长可识别前缀处理
    for cut in (12, 9, 7, 5, 3, 2):
        if length > cut:
            return cut
    return length
# 允许的编码字符（GB/T 50549-2020：大写字母+数字，禁用 I、O；连字符与点为项目扩展分隔符）
CODE_RE = re.compile(r"^[A-Z0-9][A-Z0-9.\-]*$")

# 机组号归一：本文件口径 —— 07=7号机组、08=8号机组（不是 1/2 号）
# 国标设备码形态：G F0 | F1F2F3 | F4F5 | A1A2 | T3T4T5
#               = 2位机组 + 3位系统 + 2位设备组 + 2位设备类型 + 3位序号
DEVICE_PATTERN = re.compile(r"^\d{2}[A-Z]{3}\d{2}[A-Z]{2}\d{3}$")

UNIT_BY_PREFIX = {
    "07": "7", "7": "7",
    "08": "8", "8": "8",
    "01": "1", "02": "2", "03": "3", "04": "4", "05": "5", "06": "6",
}
# 只指代机组的写法（"1号机"/"1号机组"），用于把错误的机组号纠正过来
UNIT_REF_RE = re.compile(r"([1-9]|1[0-6])\s*号\s*(机组|机)")
# 设备编号类写法（"1号主变"/"2号炉"），**不是**机组指代，不能改
EQUIP_NUM_KEEP = ("主变", "高厂变", "启备变", "炉", "机炉", "汽机", "锅炉")


@dataclass
class GenNode:
    """生成的一个设备树节点。"""
    kks: str                                  # 生成的 KKS 编码
    name: str                                 # 生成的设备名称
    parent_kks: str = ""                      # 上级 KKS 编码
    parent_name: str = ""                     # 上级设备名称
    source: str = "规则生成"                    # 原始 / 规则生成 / AI推断 / 待确认
    confidence: float = 1.0
    need_human: bool = False
    notes: list[str] = field(default_factory=list)
    # 回溯到原始行
    src_rows: list[int] = field(default_factory=list)
    profession: str = ""
    original_name: str = ""
    design_code: str = ""
    drawing: str = ""
    level: str = ""


def normalize_name(text: str) -> str:
    """名称文本归一：只做**确定无歧义**的清洗。

    ⚠️ 实测教训：不要在这里做 NFKC 全角→半角之外的改写，也不要把"号机组"缩写成"号机"。
    竣工版里 `7号机DEH（ATC）控制柜` 与 `7号机组NCS测控屏` 是**两种并存的写法**，
    统一缩写反而会把本来正确的名称改错（实测名称准确率因此从 40%→更低）。
    """
    if not text:
        return ""
    out = unicodedata.normalize("NFKC", text)          # 全角字母/数字 → 半角
    out = re.sub(r"[ \t\u3000]+", " ", out).strip()
    return out



def fix_unit_reference(name: str, kks: str) -> tuple[str, bool]:
    """把名称里的机组指代纠正为编码前缀对应的机组号。

    为什么需要：竣工版里 205+202 行的设备名称把机组号写成了「1号机/2号机」，
    而 KKS 前缀是 07/08（= 7/8 号机组）。用户已确认按 **07=7号机组** 理解。
    只改「N号机/机组」这种指代；「1号主变」这类设备编号保持原样。
    """
    if not name or len(kks) < 2:
        return name, False
    from kks_gen.compliance import plant_code
    want = UNIT_BY_PREFIX.get(plant_code(kks), "")
    if not want:
        return name, False
    changed = False

    def repl(match: re.Match[str]) -> str:
        nonlocal changed
        got = match.group(1)
        if got == want:
            return match.group(0)
        changed = True
        return f"{want}号{match.group(2)}"

    return UNIT_REF_RE.sub(repl, name), changed


# 「7号机组 / 7号机」→「7号机」的缩写（实测：竣工版把长写法统一缩成短写法）
UNIT_LONG_RE = re.compile(r"([1-9]|1[0-6])\s*号\s*机组")


def abbreviate_unit(name: str) -> str:
    """把「N号机组」缩写为「N号机」（幂等；「N号机」不受影响）。"""
    return UNIT_LONG_RE.sub(lambda m: f"{m.group(1)}号机", name or "")


def unit_fix_and_abbrev(name: str, kks: str) -> tuple[str, bool]:
    """一步到位：纠正机组号（按编码前缀）+ 缩写长写法。返回 (名称, 是否纠正过机组号)。"""
    fixed, changed = fix_unit_reference(name, kks)
    return abbreviate_unit(fixed), changed


def hierarchical_name(original: str, ancestors: dict[str, str]) -> tuple[str, str, bool]:
    """按「父名称 + 空格 + 原名称」生成层级化设备名称。

    ancestors: {父码: 父名称}。从最近的有名称的父级开始尝试拼接。
    返回 (名称, 采用的父码, 是否发生拼接)。这是实测里 ②④ 两类（前缀+原名称、
    原名称是子串）的通用形式——竣工版大量名称是"父设备名 + 空格 + 自身名"。
    """
    text = normalize_name(original)
    if not text:
        return "", "", False
    for parent_kks, parent_name in ancestors.items():
        if not parent_name:
            continue
        pname = normalize_name(parent_name)
        if not pname or text.startswith(pname):
            continue
        if len(pname) < 3:            # 父名太短（如"备用"）拼接意义不大
            continue
        return f"{pname} {text}", parent_kks, True
    return text, "", False


def code_level(code: str) -> int:
    """按主干长度判断层级（忽略 . 后缀）。"""
    base = code.split(".", 1)[0]
    return len(base)


def level_label(code: str) -> str:
    base = code.split(".", 1)[0]
    for cut in sorted(LEVEL_CUTS, reverse=True):
        if len(base) >= cut:
            return LEVEL_NAMES.get(cut, f"{cut}位")
    return f"{len(base)}位"


def ancestor_chain(code: str) -> list[str]:
    """返回该编码的全部上级（由近到远），逐级上溯到机组码为止。

    例：07AAB02UH001 → ['07AAB02', '07AAB', '07A', '07']
    （9 位设备组不是节点；2 位机组码到顶）
    """
    base = code.split(".", 1)[0]
    chain: list[str] = []
    seen: set[str] = set()
    current = direct_parent(base)
    while current and current not in seen:
        seen.add(current)
        chain.append(current)
        current = direct_parent(current)
    return chain


def direct_parent(code: str) -> str:
    """直接上级。三类后缀的挂法**实测不同**，不能一刀切：

    | 形态 | 例 | 上级 | 依据 |
    |---|---|---|---|
    | 点后缀 | `07BFA01UC001.A1` | `07BFA01UC001`（主干） | 竣工版行 559：上级是主干设备 |
    | 字母尾缀 | `07ABG01BA001A` | `07ABG01`（7位子系统） | 竣工版行 60：上级是子系统，A/B/C 是分相 |
    | 标准层级 | `07AAB02UH001` | `07AAB02` | 设备分类码不单独成节点 |

    另外：2 位机组码（`07`）**没有**上级，不能拼出 `"0"`。
    """
    raw = str(code or "").strip()
    if not raw:
        return ""
    from kks_gen import identifiers as ids
    if raw.startswith(("+", ";", ":", "&", "#", "-")):
        return ""  # Non-functional identifiers must not use functional cuts.
    match = ids.RE_FUNCTION.fullmatch(raw)
    if match:
        seg = match.groupdict()
        prefix = (seg.get("h1") or "") + seg["h23"] + (seg.get("h4") or "")
        return prefix + seg["sys"] + seg["sysno"]
    match = ids.RE_SYSTEM_ONLY.fullmatch(raw)
    if match:
        seg = match.groupdict()
        prefix = (seg.get("h1") or "") + seg["h23"] + (seg.get("h4") or "")
        return prefix + seg["sys"]
    partial = re.fullmatch(r"([BCEGKTUW]?[0-9]{2}[0-9]?)([A-HJ-NP-Z]{1,3})?", raw)
    if partial:
        prefix, group = partial.groups()
        return prefix + group[0] if group and len(group) == 3 else (prefix if group else "")
    # ① 点后缀：挂在主干下
    if "." in raw:
        base, _, tail = raw.partition(".")
        return base if base and tail else ""
    # ② 连字符形态：要分两种，不能都当"部件"
    #    · 部件：连字符前是**12 位设备码**（07AAB02UH001-KF01）→ 挂在设备下
    #    · 系统下设备：连字符前是**5 位系统码**（07CGQ-N001、07CGQ-D001）
    #      → 这类是"系统码 + - + 设备字母+编号"，上级是 5 位系统码本身
    #    （实测：07CGQ-N001 「7号炉1号炉管泄露探头」，上级应为 07CGQ）
    if "-" in raw:
        head, _, tail = raw.partition("-")
        if tail and len(head) == 5:
            return head
        return head if head and tail else ""
    # ③ 字母尾缀（13 位、末位字母）：挂到主干 12 位所对应的层级（= 7 位子系统）
    # ⚠️ 分相只可能是 A/B/C —— 限定范围。实测教训：放宽成"任意字母"时，
    #    `01BTM11JYZZGZ`（末位 Z，其实是别的写法）会被硬挂到 `01BTM11`，
    #    凭空造出一个错误上级。宁可留空交人工，也不要猜。
    if len(raw) == 13 and raw[-1] in ("A", "B", "C"):
        return direct_parent(raw[:12])
    # ④ 国标设备码 + 部件/序号后缀（现场汇总表常见，省略了连字符）：
    #    `01HAL11GP001XB21` = 设备 01HAL11GP001 + 部件 XB21 → 上级是那台设备
    #    `01HDU10BP101_1`    = 设备 01HDU10BP101 + 序号 1   → 上级是那台设备
    #    实测（鄂州某汇总表）：16 位里 16,504/16,675、14 位里 4,960/5,084 都属这种。
    base = raw.split(".", 1)[0]
    if len(base) > 12 and DEVICE_PATTERN.match(base[:12]):
        return base[:12]
    if "_" in base:
        head = base.split("_", 1)[0]
        if DEVICE_PATTERN.match(head) or (len(head) in PARENT_LEN and head.isascii()):
            return head
    # ⑤ 标准层级映射
    if len(raw) not in PARENT_LEN:
        return ""
    parent_len = PARENT_LEN[len(raw)]
    if parent_len <= 0 or parent_len >= len(raw):
        return ""
    return raw[:parent_len]


def split_design_suffix(design_code: str) -> tuple[str, str]:
    """拆出主干与后缀。

    - 点后缀：07BFA01UC003.A1 → ('07BFA01UC003', '.A1')
    - 分相/变体尾缀：07ABG01BA001A → ('07ABG01BA001', 'A')   ← 实测其父级是主干 12 位
    - 无后缀：07AAB02UH001 → ('07AAB02UH001', '')
    """
    code = str(design_code or "").strip().upper()
    if "." in code:
        base, _, tail = code.partition(".")
        return base, "." + tail
    if len(code) == 13 and code[-1].isalpha():
        return code[:12], code[12:]
    return code, ""


def _common_name_prefix(names: list[str], *, min_len: int = 3) -> str:
    """求一组名称的最长公共前缀（机组号先归一，避免 7号/1号 写法差异干扰）。"""
    cleaned = [abbreviate_unit(n) for n in names if n]
    if not cleaned:
        return ""
    prefix = cleaned[0]
    for name in cleaned[1:]:
        i = 0
        limit = min(len(prefix), len(name))
        while i < limit and prefix[i] == name[i]:
            i += 1
        prefix = prefix[:i]
        if len(prefix) < min_len:
            return ""
    return prefix.strip()


_CODE_INDEX_CACHE: dict | None = None


def _code_index() -> dict:
    """加载 `config/kks_code_index.json`（模块级缓存）。

    ⚠️ 原实现每次调用都重读一遍 157 KB 的 JSON；父级命名要处理几千个节点 × 3 轮，
    会读上万次。这里缓存一次。
    """
    global _CODE_INDEX_CACHE
    if _CODE_INDEX_CACHE is None:
        try:
            import json
            from pathlib import Path as _Path
            _CODE_INDEX_CACHE = json.loads(
                (_Path(__file__).resolve().parents[1] / "config" / "kks_code_index.json")
                .read_text(encoding="utf-8"))
        except (OSError, ValueError, TypeError):
            _CODE_INDEX_CACHE = {}
    return _CODE_INDEX_CACHE


def _unit_display(code2: str) -> str:
    """全厂码前两位 → 机组名（表4.3.1-2）。"""
    if not code2.isdigit():
        return ""
    n = int(code2)
    if n == 0:
        return "全厂公用"
    if 1 <= n <= 60:
        return f"{n}号机组"
    if 61 <= n <= 90:                      # 61↔(1,2)、62↔(3,4)…90↔(59,60)
        k = n - 60
        return f"{2 * k - 1}、{2 * k}号机组共用"
    return "多台机组共用"                     # 91~99


def _unit_prefix(code2: str) -> str:
    """机组前缀（用于 3/5 位父级名）：`7号机` / `1、2号机` / `全厂`。

    ⚠️ 不能拿 `_unit_display()` 的结果去 `[:-2]` 截断：`1、2号机组共用` 会被截成
    `1、2号机组`（`共用` 没了但 `组` 还在），`多台机组共用` 同理。直接按码值算。
    """
    if not code2.isdigit():
        return ""
    n = int(code2)
    if n == 0:
        return "全厂"
    if 1 <= n <= 60:
        return f"{n}号机"
    if 61 <= n <= 90:
        k = n - 60
        return f"{2 * k - 1}、{2 * k}号机"
    return "多台机组"


def _short_label(label: str, limit: int = 18) -> str:
    """把码表释义压成**短名**：去掉括号补充说明 → 只取第一个分句 → 限长。

    附录A 的 1 字母（主组）释义常常是整句话，例如
    `H` = 「热产生，由化石燃烧和可再生能源，以及天然能源的热回收」，
    直接拼进名字会长成 `1号机热产生，由化石燃烧和可再生能源，以及天然能源的热回收`。
    竣工版的同类名是短的（`7号机电网和配电系统`），所以这里取第一分句。
    """
    text = re.split(r"[（(]", str(label or ""), maxsplit=1)[0]
    text = re.split(r"[，,、;；]", text, maxsplit=1)[0].strip()
    return text[:limit]


def _name_from_code(kks: str) -> tuple[str, str]:
    """按**国标码表语义 + 机组号**给父级节点命名（业务确认口径，2026-10-10）。

    | KKS 长度 | 规则 | 例 |
    |---|---|---|
    | 2（机组） | 表4.3.1-2 全厂码 | `07`→`7号机组`、`61`→`1、2号机组共用`、`00`→`全厂公用` |
    | 3（全厂码+S₁） | 附录A 的 1 字母名 + 机组前缀 | `07A`→`7号机电网与配电系统` |
    | 5（+S₁S₂S₃） | 附录A 的 3 字母名 + 机组前缀 | `07HAC`→`7号机省煤器系统` |

    码表查不到 → 返回 `("", "")`，**不编造**（宁可留空）。
    """
    base = kks.split(".", 1)[0]
    match = re.fullmatch(r"[BCEGKTUW]?([0-9]{2})[0-9]?([A-HJ-NP-Z]{1,3})?", base)
    if not match:
        return "", ""
    idx = _code_index()
    unit2, classification = match[1], match[2] or ""
    if not classification:
        disp = _unit_display(unit2)
        return (disp, "表4.3.1-2 全厂码") if disp else ("", "")
    prefix = _unit_prefix(unit2)
    if not prefix:
        return "", ""
    if len(classification) == 1:
        seg = classification
        label = _short_label((idx.get("system_s1_groups") or {}).get(seg, ""))
        if label:
            return f"{prefix}{label}", f"附录A「{seg}」+机组号"
        return "", ""
    if len(classification) == 3:
        seg = classification
        label = _short_label(
            str((idx.get("system_codes") or {}).get(seg, "")).split(" / ", 1)[0])
        if label:
            return f"{prefix}{label}", f"附录A「{seg}」+机组号"
        return "", ""
    return "", ""


def _system_code_label(code: str) -> str:
    """查系统分类码字典，取该码的语义首词（用于拼父节点名称）。"""
    base = code.split(".", 1)[0]
    if len(base) < 5:
        return ""
    from kks_gen.identifiers import parse
    system = parse(base).segments.get("sys", "")
    if not system:
        match = re.fullmatch(r"[BCEGKTUW]?[0-9]{2}[0-9]?([A-HJ-NP-Z]{3})", base)
        system = match[1] if match else ""
    label = str((_code_index().get("system_codes") or {}).get(system, ""))
    if not label:
        return ""
    return label.split(" / ", 1)[0][:14]


def build_tree(seed_rows: Iterable[dict[str, Any]], *, name_resolver=None,
               progress_callback=None) -> dict[str, Any]:
    """由种子行生成完整设备树。

    seed_rows 每项需含：excel_row、design_code、original_name、profession、drawing。
    返回 {'nodes': {kks: GenNode}, 'ordered': [kks...], 'stats': {...}}
    """
    nodes: dict[str, GenNode] = {}
    stats = {
        "seed_rows": 0, "seed_without_code": 0,
        "parent_nodes_created": 0, "unit_fixed": 0,
        "need_ai_name": 0, "duplicate_kks": 0,
    }
    dup_seen: set[tuple[str, str]] = set()      # (最终码, 原始码) —— 用来区分"有意共享"与"真重复"
    body_named: set[str] = set()                # 名称已由「设备本体行」确定的节点（权威名）
    seed_rows = list(seed_rows)
    total_rows = len(seed_rows)
    # 预扫一遍种子：判断"某 12 位主干是否真的存在"要用到（见下面的 13 位回退）
    seen_design: set[str] = set()
    for row in seed_rows:
        code = str(row.get("design_code", "") or "").strip().upper()
        if code:
            seen_design.add(code)

    def ensure_parent(kks: str, child: GenNode) -> None:
        """确保父节点存在（父节点没有设计编码，是"后人补的"那部分）。"""
        if kks in nodes:
            if not nodes[kks].parent_kks:
                nxt = direct_parent(kks)
                if nxt:
                    nodes[kks].parent_kks = nxt
                    nodes[kks].parent_name = nodes[nxt].name if nxt in nodes else ""
            return
        node = GenNode(
            kks=kks,
            name="",                                  # 父节点名称待生成
            source="规则生成(父节点)",
            confidence=0.6,                            # 名称未知 → 降置信度
            need_human=True,
            notes=["父节点由编码层级推导，名称待生成"],
            src_rows=list(child.src_rows),
            profession=child.profession,
            level=level_label(kks),
        )
        nxt = direct_parent(kks)
        if nxt:
            node.parent_kks = nxt
        nodes[kks] = node
        stats["parent_nodes_created"] += 1
        if nxt:
            ensure_parent(nxt, node)

    for index, row in enumerate(seed_rows, start=1):
        if progress_callback is not None and index % 20000 == 0:
            # 建树阶段占总进度 10%~25%（前端进度条用）
            try:
                progress_callback({
                    "stage": "tree_building", "percent": 10 + int(index / max(1, total_rows) * 15),
                    "message": f"正在构建设备树（{index:,}/{total_rows:,} 行）",
                    "hint": f"已生成节点 {len(nodes):,} 个",
                })
            except Exception:      # 进度上报不能影响生成
                pass
        design = str(row.get("design_code", "") or "").strip().upper()
        if not design:
            stats["seed_without_code"] += 1
            continue
        stats["seed_rows"] += 1
        base, suffix = split_design_suffix(design)
        raw_name = str(row.get("original_name", "") or "").strip()
        # ⚠️ 名称栏里填的若是**标签串**（`汽消 汽电 汽电 汽源`，源表脏数据）→ 当作"没有名称"。
        #    **必须在源头拦**：否则它会被当"原始名称"，再被"父名+空格+自身名"拼成
        #    `…控制电源柜 汽消 汽电 汽电 汽源`（看着像名字，后续所有清洗都拦不住），
        #    最后靠兄弟行继承扩散到整组——实测 `01CSB02XA001~005` 两个机组共 10 行全中招。
        #    全表 58,280 行只命中 6 行（3 种写法），精度高。
        from name_rules import tag_junk
        junk_name = ""
        if tag_junk(raw_name):
            stats["name_tag_junk"] = stats.get("name_tag_junk", 0) + 1
            junk_name, raw_name = raw_name, ""

        # KKS 编码：主干 + 后缀原样保留（实测 3897 行带 .A1/.J1 这类后缀，直接沿用）
        kks = design
        # 名称处理，实测校准过的三步（顺序不能换）：
        #   ① 纠正错误机组号：`1号机NCS测控屏` → `7号机NCS测控屏`（07=7号机组）
        #   ② **仅当机组号本来就正确时**，把长写法缩写为短写法：`7号机组X` → `7号机X`
        #      依据：竣工版设备名称里短写法 25,602 处 vs 长写法仅 258 处（99:1）
        #      ⚠️ 若对"该纠正的长写法"也缩写（`1号机组X` → `1号机X`），会把源头错误留在结果里
        #        —— 上轮无条件缩写导致准确率 39.4%→38.6%，根因就是这个。
        name, unit_changed = fix_unit_reference(normalize_name(raw_name), base)
        if unit_changed:
            stats["unit_fixed"] += 1
        else:
            abbreviated = abbreviate_unit(name)
            if abbreviated != name:
                name = abbreviated
                stats["unit_abbreviated"] = stats.get("unit_abbreviated", 0) + 1

        node = GenNode(
            kks=kks,
            name=name,
            source="原始" if raw_name else "待确认",
            confidence=1.0 if raw_name else 0.4,
            need_human=not raw_name,
            src_rows=[int(row["excel_row"])] if row.get("excel_row") else [],
            profession=str(row.get("profession", "") or ""),
            original_name=raw_name,
            design_code=design,
            drawing=str(row.get("drawing", "") or ""),
            level=level_label(base),
        )
        if junk_name:
            stats["need_ai_name"] += 1
            node.notes.append(f"原名称栏填的是标签串「{junk_name}」，不是设备名——"
                              f"名称改按编码语义+图号推断，需人工确认")
        elif not raw_name:
            stats["need_ai_name"] += 1
            node.notes.append("原名称为空，设备名称需按编码语义+图号推断")
        # 重复判定：只把「**同一原始编码**再次出现」当真正的重复行。
        # ⚠️ 不能只看最终码 —— 指令/测点行（原码 = 设备码 + 信号码）**有意共享**设备本体码
        #    （它们不写交付码，见 `revise` 的 `not_device` 口径），那不是重复，
        #    误报会把它标成"需人工确认"（踩过）。
        orig_code = str(row.get("original_code") or design or "")
        if (kks, orig_code) in dup_seen:
            stats["duplicate_kks"] += 1
            node.notes.append("与其他行编码完全相同，需人工确认")
            node.need_human = True
            node.confidence = min(node.confidence, 0.5)
        dup_seen.add((kks, orig_code))

        parent = direct_parent(kks)
        # 回退：13 位扩展码默认挂到 12 位主干，但主干可能**不存在**（如 07LAB13BBP101
        # 的 12 位主干 07LAB13BBP10 表里没有）。此时退回 7 位子系统，与其它分相/扩展码一致。
        if parent and len(kks.split(".")[0]) == 13 and parent not in nodes and parent not in seen_design:
            fallback = parent.split(".")[0][:7]
            if len(fallback) == 7:
                parent = fallback
        if parent:
            ensure_parent(parent, node)
            node.parent_kks = parent
            # 名称层级化：原名称已含父名称时，在父名称后插空格断开；
            # 原名称孤立（不含机组/系统信息）时，用最近的父名称做前缀。
            # 实测依据（竣工版）：V3"父名后插空格"298 行、V2"父名+空格+原名"298 行。
            if name and not UNIT_REF_RE.search(name) and nodes.get(parent) and nodes[parent].name:
                pname = normalize_name(nodes[parent].name)
                if pname and name.startswith(pname) and len(name) > len(pname):
                    broken = f"{pname} {name[len(pname):].lstrip()}"
                    if broken != name:
                        node.name = broken
                        node.notes.append(f"名称在父级({parent})后断开")
                        node.source = "规则生成(父名后断开)"
                        node.confidence = min(node.confidence, 0.85)
                else:
                    composed, used_parent, joined = hierarchical_name(name, {parent: pname})
                    if joined and composed != name:
                        node.name = composed
                        node.notes.append(f"名称按父级({used_parent})补前缀")
                        node.source = "规则生成(父名前缀)"
                        node.confidence = min(node.confidence, 0.8)
            # 原名称为空时，走名称规则层（部件类型词 / 柜号 / 兜底待AI）
            if not name:
                from name_rules import generate_name
                generated, rule, conf, human = generate_name(
                    code=kks,
                    original_name="",
                    parent_name=nodes[parent].name,
                    parent_code=parent,
                )
                node.name = generated
                node.source = f"规则生成({rule})" if generated else "待AI推断"
                node.confidence = min(node.confidence, conf)
                node.need_human = node.need_human or human
                node.notes.append(f"名称由规则生成：{rule}")

        if kks in nodes and nodes[kks].design_code:
            # 同一编码出现多次：合并来源行，保留置信度较低者
            existing = nodes[kks]
            # 名称归属：**优先用「设备本体行」的名称**（原始码没有信号后缀、等于本体码）。
            # 测点行的名称常带运行状态（`A一次风机已运行`），不该成为设备名。
            # ⚠️ 不这么定就有**顺序依赖**：谁先出现谁定名，测点行在前时设备名会带状态词（踩过）。
            new_is_body = (orig_code == kks)
            if new_is_body and kks not in body_named and node.name:
                existing.name = node.name
                existing.source = node.source
                existing.confidence = node.confidence
                body_named.add(kks)
            existing.src_rows.extend(node.src_rows)
            existing.need_human = existing.need_human or node.need_human
            existing.confidence = min(existing.confidence, node.confidence)
            # ⚠️ **notes 也要并**：重复检测的说明写在**新节点**的 notes 上；
            #    原来只并了 src_rows/need_human/confidence，
            #    结果 `need_human=True` 但**原因丢失** —— 交付文件里看到"需人核"却不知道为什么（踩过）。
            existing.notes.extend(n for n in node.notes if n not in existing.notes)
            continue
        nodes[kks] = node
        if orig_code == kks:                       # 设备本体行 → 它的名字最有权威
            body_named.add(kks)

    # ---- 第二轮：父节点名称推导 ----
    # 优先级（业务确认口径，2026-10-10）：
    #   ① **2/3/5 位（机组/主组/系统）**：按**国标码表语义 + 机组号**确定性命名 —— 与竣工版
    #      `07A`→`7号机电网和配电系统` 同一路数，且完全可复算；
    #   ② **7/12 位**：子节点名称的公共前缀（**必须严格短于最短子名**，否则就是照抄某个子节点）；
    #   ③ 兜底：最近有名祖先 + 系统分类码语义；
    #   ④ 都给不出 → **留空**（宁缺勿错），交人工。
    derived = 0
    # 预建"父→子"索引：65k 节点下不能在对每个父节点循环时全表扫描
    children_of: dict[str, list[str]] = {}
    for kks, node in nodes.items():
        if node.parent_kks:
            children_of.setdefault(node.parent_kks, []).append(kks)
    for _ in range(3):                    # 自底向上多轮，让上层父节点也能拿到名称
        for node in nodes.values():
            if node.name or node.design_code:
                continue                  # 只处理"没有设计编码的父节点"
            # ① 2/3/5 位 → 码表语义 + 机组号
            from_code, why = _name_from_code(node.kks)
            if from_code:
                node.name = from_code
                node.source = f"规则生成({why})"
                node.confidence = max(node.confidence, 0.8)
                node.notes.append(f"名称按国标码表推导：{why} → {from_code}")
                derived += 1
                continue
            # ② 7/12 位 → 子节点名称公共前缀
            direct = children_of.get(node.kks, [])
            child_names = [nodes[c].name for c in direct if nodes[c].name]
            if not child_names:
                # 直接子节点都没名称（如主组层 07R）：用"孙子层"的名称求公共前缀。
                # 实测：07R 的子节点 07REK/07REQ 名称都以"7号脱硫…"开头 → 主组名 = "7号脱硫"。
                grand: list[str] = []
                for child in direct:
                    grand.extend(nodes[g].name for g in children_of.get(child, []) if nodes[g].name)
                child_names = grand
            # ②-b 7 位（子系统）且**只有一个子节点**：竣工版的做法是「子节点名 + 系统」
            #     （实测：`07AAB02` 的子是 `1号机NCS测控屏` → 竣工版名 `1号机NCS测控屏系统`）。
            #     这类占 7 位节点的 525 个；若不特殊处理会落到下面的兜底，拼出重复的
            #     `…操作和监控系统操作和监控系统`（实测踩过）。
            if len(node.kks.split(".", 1)[0]) == 7:
                only = [nodes[c].name for c in direct if nodes[c].name]
                if len(only) == 1:
                    base = only[0]
                    node.name = base if base.endswith(("系统", "装置")) else f"{base}系统"
                    node.source = "规则生成(唯一子节点名+系统)"
                    node.confidence = max(node.confidence, 0.7)
                    node.notes.append(f"名称由唯一子节点名「{base}」加『系统』推导")
                    derived += 1
                    continue
            if child_names:
                prefix = _common_name_prefix(child_names)
                # ⚠️ 必须**严格短于最短子名**：否则"公共前缀"等于某个子节点的整名，
                #    父节点就变成了照抄子行名（实测 1,310 例，如 `00A` 抄 `500kV GIS…断路器`）。
                if len(prefix) >= 3 and len(prefix) < min(len(n) for n in child_names):
                    node.name = prefix
                    node.source = "规则生成(子名称公共前缀)"
                    node.confidence = max(node.confidence, 0.75)
                    node.notes.append(f"名称由 {len(child_names)} 个子节点名称的公共前缀推导")
                    derived += 1
                    continue
            # 推不出前缀：父节点名称 = 最近的有名祖先 + 系统分类码语义
            label = _system_code_label(node.kks)
            ancestor_name = ""
            cur = node.parent_kks
            while cur and cur in nodes:
                if nodes[cur].name:
                    ancestor_name = nodes[cur].name
                    break
                cur = nodes[cur].parent_kks
            # ⚠️ 祖先名里**已经含**该码语义时不能再拼，否则会拼出
            #    `…操作和监控系统操作和监控系统` 这种重复名（实测踩过）。
            #    用「标签前 6 字是否已出现在祖先名里」判断——标签可能被截断/带逗号后缀，
            #    直接 `label in ancestor` 会漏判（实测：`中压主供电系统 1, 电压等` vs
            #    祖先 `7号机中压主供电系统 1` 判不出包含关系）。
            #    拼接加**空格分隔**，否则读起来是半截词。
            if label and ancestor_name and label[:6] not in ancestor_name:
                node.name = f"{ancestor_name} {label}"
                node.source = "规则生成(祖名+系统码语义)"
                node.confidence = max(node.confidence, 0.6)
                node.notes.append(f"名称由祖先名+系统分类码「{label}」推导")
                derived += 1
    stats["parent_names_derived"] = derived

    # ---- 第三轮：兄弟行继承名称（解决"原名称为空"的行）----
    # 实测依据（鄂州四期 12,268 行空名称行）：同父级+同设备分类码的兄弟名称，
    # 去掉末尾序号/分相后作为基线名，再按本行补序号或分相 → 完全命中 3,815 行。
    from name_rules import sibling_name

    by_parent_class: dict[tuple[str, str], list[str]] = {}
    for node in nodes.values():
        if not node.name or not node.parent_kks:
            continue
        base = node.kks.split(".", 1)[0]
        from kks_gen.identifiers import parse
        cls = parse(base).segments.get("dev", "")
        by_parent_class.setdefault((node.parent_kks, cls), []).append(node.name)

    inherited = 0
    for node in nodes.values():
        if node.name or node.design_code == "":
            continue                       # 已有名称，或不是种子行
        base = node.kks.split(".", 1)[0]
        cls = parse(base).segments.get("dev", "")
        names = by_parent_class.get((node.parent_kks, cls), [])
        cand, ok = sibling_name(node.kks, names)
        if ok and cand:
            node.name = cand
            node.source = "规则生成(兄弟继承)"
            node.confidence = max(node.confidence, 0.7)
            node.need_human = True         # 仍是推断，需人核
            node.notes.append("名称由同父级同类型兄弟行继承")
            inherited += 1
    stats["name_inherited_from_sibling"] = inherited

    # 补父级名称：父节点没有名称时，回填"父节点名称"字段供下游生成
    for node in nodes.values():
        if node.parent_kks and node.parent_kks in nodes:
            node.parent_name = nodes[node.parent_kks].name

    return {"nodes": nodes, "stats": stats}


def order_nodes(nodes: dict[str, GenNode]) -> list[str]:
    """按 机组 → 系统 → 子系统 → 设备 → 部件 的自然顺序输出。"""
    def key(item: tuple[str, GenNode]) -> tuple:
        kks, node = item
        return (node.profession or "zz", kks[:2], len(kks.split(".", 1)[0]), kks)
    return [kks for kks, _ in sorted(nodes.items(), key=key)]
