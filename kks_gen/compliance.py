"""Shared, conservative checks against GB/T 50549-2020 and project rules.

Passing these automated checks is not a certificate: drawings, engineering
agreements and plant-wide uniqueness still require evidence outside one sheet.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from . import identifiers as ids
from .gb_rules import unit_label_to_plant_code


@dataclass
class Assessment:
    kind: str
    errors: list[str] = field(default_factory=list)
    pending: list[str] = field(default_factory=list)

    @property
    def writable(self) -> bool:
        return self.kind != "unknown" and not self.errors


PLANT = re.compile(r"^[BCEGKTUW]?(?P<plant>[0-9]{2})(?:[0-9])?(?=[A-Z]|$)")
HIERARCHY = re.compile(r"^[BCEGKTUW]?[0-9]{2}(?:[0-9])?(?:[A-HJ-NP-Z]{1,3})?$")


def plant_code(code: str) -> str:
    match = PLANT.match(str(code or ""))
    return match.group("plant") if match else ""


def unit_conflicts(code: str, name: str = "", unit: str = "") -> list[str]:
    prefix = plant_code(code)
    if not prefix:
        return []
    result = []
    for label, value in (("名称", name), ("机组列", unit)):
        # Names can refer to the served equipment; only an explicit leading
        # unit qualifier is evidence of ownership.
        value = str(value or "")
        for cn, number in zip("一二三四五六七八九十", range(1, 11)):
            value = value.replace(cn + "号", str(number) + "号")
        if label == "名称" and not re.match(
            r"^\s*第?[0-9一二三四五六七八九十号#、,，及和~～\-\s]+(?:机组|机(?!械))", value
        ):
            continue
        expected = unit_label_to_plant_code(value)
        if expected and prefix != expected:
            # 91..99 are assigned by agreement; no universal mapping exists.
            if int(prefix) >= 91 and int(expected) >= 91:
                continue
            result.append(f"{label} {value!r} 与全厂码 H₂H₃={prefix} 的范围不一致（表4.3.1-2）；核对图纸和工程约定，不自动改归属")
    return result


def assess(code: str) -> Assessment:
    raw = str(code or "")
    # A combination is checked as two complete identifiers, never by dropping
    # everything after a separator. '=' is the optional functional prefix.
    combination = re.fullmatch(r"=?(.+?)(-[A-HJ-NP-Z]{2}[0-9]{2,3}(?:-[A-HJ-NP-Z]{2}[0-9]{2,3})*|;.+)", raw)
    if combination and ids.RE_FUNCTION.fullmatch(combination[1]):
        left, right = assess(combination[1]), assess(combination[2])
        return Assessment("combination", left.errors + right.errors, left.pending + right.pending)
    if raw.startswith("="):
        return assess(raw[1:])
    parsed = ids.parse(raw)
    out = Assessment(parsed.kind)
    if not raw or raw != raw.strip() or re.search(r"\s", raw) or raw != raw.upper():
        out.errors.append("标识为空或含空白/小写字符；应使用规定的大写字符（第4章）")
    if parsed.kind != "terminal" and any(c in raw for c in "IO"):
        out.errors.append("含规定字母段的禁用字母I/O（第4章相应结构图注）")
    if HIERARCHY.fullmatch(raw):
        out.kind = "hierarchy"
        # Partial classification nodes are an application tree convention,
        # not a complete equipment identifier.
        if len(raw) > 2:
            out.pending.append("分类层级节点须与工程编码索引一致（3.3）；不能据此认定完整设备标识合规")
    else:
        out.errors.extend(f"{f.message}（{f.basis}）" for f in parsed.errors)
        out.pending.extend(f"{f.message}（{f.basis}）" for f in parsed.pendings)
        out.pending.extend(f"{f.message}（{f.basis}）" for f in parsed.warnings if f.segment != "plant")
        if parsed.kind == "unknown":
            out.errors.append("未识别为国标标识结构；核对标识类型及工程约定（第4章）")
        if parsed.kind == "function" and not ids.RE_FUNCTION.fullmatch(raw.lstrip("=")):
            out.errors.append("功能面设备标识须完整写出系统编号两位、设备编号三位；附加写法须另核工程约定（4.4.4、4.4.5）")
        if parsed.kind == "system" and not ids.RE_SYSTEM_ONLY.fullmatch(raw.lstrip("=")):
            out.errors.append("系统标识须完整写出系统编号两位（4.4.4）")
        if parsed.kind == "product" and any(not ids.RE_PRODUCT_SEG.fullmatch(s) for s in raw[1:].split("-")):
            out.errors.append("产品面每一级均须符合分类码和编号结构（4.5）")
        if parsed.kind == "union" and not ids.RE_UNION.fullmatch(raw):
            out.errors.append("联合标识每一级不得超过12位，须符合两级结构（4.2）")
        if parsed.kind == "terminal" and not ids.RE_TERMINAL.fullmatch(raw):
            out.errors.append("端子标识每一级不得超过12位，须符合端子结构（4.8）")
    devno = parsed.segments.get("devno", "")
    if devno == "000":
        out.errors.append("设备编号000不在一般流水编号001～999范围；须提供适用条款或工程约定（4.4.5）")
    prefix = parsed.segments.get("h23") or plant_code(raw)
    if prefix and int(prefix) >= 91:
        out.pending.append("全厂码91～99对应的三机及以上共用范围须由各方约定（表4.3.1-2）；缺少约定不能判通过")
    return out
