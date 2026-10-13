# -*- coding: utf-8 -*-
"""抽取中煤《应用细则》全部分类码表，并合成项目可用码表。

表类型判定：**读表头 + 续表继承**，不按码长猜。

🔴 踩过的坑（四条，都在这里修掉）：
  1. **按码长路由**（2 位→设备码、3 位→系统码）会把 `XA 通风和空调系统` 这类**系统码**
     错放进设备表，得出"中煤系统码新增 0"的错误统计。→ 改为读表头。
  2. **只读 `cells[0]`**：很多表的码在**中间列**——正文表「序号 | 名称 | 编码原则 | 说明」
     的码在**第 3 列**；表8.11.3「系统码 | 视频监控类设备码 | 说明」的码在**第 2 列**，
     且首列是 `XSB 视频监控系统` 这种**合并列**；表8.11.4 首列是 `XSA01 厂区入口`。
     → 改为**逐格识别码格**，不再假定首列是码。
  3. **续表没有表头**：细则附录 A/B 各被拆成多张表（表26/27 有表头，表28~48 是续表；
     表49/50 有表头，表51~73 是续表，**首行直接是数据行**如 `B | 将输入变量…`）。
     原脚本把续表判为 `unknown` 后**整表跳过** → 中煤附录 A 大量码从未入库。
     → 续表**继承上一张有表头表的 kind**。
  4. **续表的"首格 1 位字母 + 次格中文"兜底规则是错的**：附录 B 续表的首行 `B | 将输入变量…`
     看起来像"主组分节表"，于是表51~73 被整片判成**系统表**，把 `BE/SF/TA/TR/TF…` 这些
     **T₂ 设备码错标成系统码**（这是 project_system_codes.json 里那批 2 字母"系统码"的来源，
     与 4.4.2「系统分类码固定 3 字母」冲突）。→ 删除该兜底规则，改由**继承 + 长度硬校验**决定。

硬校验（依 4.4.2/4.4.3）：**系统分类码 S₁S₂S₃ 固定 3 字母，设备分类码 T₁T₂ 固定 2 字母**
→ 见到 3 字母码一律判 system；1~2 字母码按所在表/列的类型判。长度 >3 的（如 `METS`/`MTSI`）
登记为**缺陷**，不入可用集。

产物（全部可复跑）：
  config/zhongmei_codes.json           {code: {cn, kind, kinds[], alt_cn, source}}
  config/project_system_codes.json     项目可用系统分类码 = 国标附录A ∪ 中煤系统码
  config/project_device_codes.json     项目可用设备分类码 = 国标附录B ∪ 中煤设备码
  config/zhongmei_unique_codes.json    中煤独创码（国标附录同段位查不到）
  config/zhongmei_excluded_codes.json  项目**禁用**码 = 国标附录A 中 S₁ ∈ {J,K} 的核电码
  config/zhongmei_star_codes.json      ⭐「可用」开放性槽位赋义码 + 国标外真新造码
  config/zhongmei_defects.json         抽到的**非法形态**（如 4 字母的 METS/MTSI），不入可用集
  config/gb_vs_zhongmei_diff.json      逐码差异明细（追溯用；由本脚本生成，避免与字典漂移）
"""
from __future__ import annotations

import json
import pathlib
import re
import sys

from docx import Document

# Windows 控制台默认 GBK，而本脚本会打印 `S₁` 等非 GBK 字符（下标、✅）；
# 不显式声明就会在**最后一行 print** 上抛 UnicodeEncodeError，把跑成功的抽取误报成失败。
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

CFG = pathlib.Path("config")
CFG.mkdir(exist_ok=True)

# 源文件定位：优先 argv[1]；不传则自动在「国标/」下寻找。
# ⚠️ 中文路径经 Git Bash 的 argv 传参会被 ANSI 代码页破坏（实测 PackageNotFoundError），
#    所以**推荐直接无参运行**：python -B tools/extract_zhongmei_codes.py
if len(sys.argv) > 1:
    SRC = pathlib.Path(sys.argv[1])
else:
    _cands = sorted(pathlib.Path("国标").glob("*A3.2*.docx"))
    if not _cands:
        raise SystemExit("用法: python -B tools/extract_zhongmei_codes.py [中煤细则.docx]（未找到 国标/*A3.2*.docx）")
    SRC = _cands[0]
print(f"数据源: {SRC}")
doc = Document(str(SRC))

CODE1 = re.compile(r"^[A-Z]$")
CODE2 = re.compile(r"^[A-Z]{2}$")
CODE3 = re.compile(r"^[A-Z]{3}$")
CJK = re.compile(r"[一-鿿]")
CODE_HEAD = ("系统分类码", "系统代码", "系统码", "设备分类码", "设备码", "设备编号",
             "产品分类码", "分类码", "编码原则", "编码")
NAME_HEAD = ("系统名称", "设备名称", "名称", "专业", "简述", "说明", "用途", "备注")
SYS_HINT = ("系统分类码", "系统代码", "系统码", "系统名称")
DEV_HINT = ("设备分类码", "设备码", "设备编号", "产品分类码")


def clean(cells: list[str]) -> list[str]:
    """去掉合并单元格造成的相邻重复。"""
    out: list[str] = []
    for c in cells:
        c = re.sub(r"\s+", " ", c).strip()
        if not out or c != out[-1]:
            out.append(c)
    return out


def header_kind(rows: list[list[str]]) -> str:
    """按表头判定 system / device / unknown（device 优先：表头常同时含两者）。"""
    for row in rows[:3]:
        text = " ".join(row)
        if any(h in text for h in DEV_HINT):
            return "device"
        if any(h in text for h in SYS_HINT):
            return "system"
    return "unknown"


def name_column(rows: list[list[str]], code_col: int) -> int | None:
    """名称列：优先「名称/专业」，其次其他 NAME_HEAD 列。"""
    for rank in (("系统名称", "设备名称", "名称", "专业"), NAME_HEAD):
        for row in rows[:3]:
            for i, cell in enumerate(row):
                if i != code_col and any(h in cell for h in rank):
                    return i
    return None


def codes_in(cell: str, allow_merged: bool = False) -> list[str]:
    """从一个单元格取出所有分类码。

    支持 `A` / `AB` / `ABC` / `ABC01`（系统码+编号）/ `A（说明）、B（说明）` 并列；
    另在 `allow_merged=True` 时支持 `XSB 视频监控系统`（码+空格+中文，**合并列**常见写法）。

    ⚠️ `allow_merged` **必须**只在「表头认定的码列」上打开。否则名称列里的
    `MIS 机柜、箱`、`DCS 控制、端子等 机柜`、`SIS 机柜` 会被误抽成分类码
    `MIS`/`DCS`/`SIS`（实测踩过：它们被当成"中煤独创系统码"混进了项目码表）。

    含中文且不符合上述形式 → 返回空（该格不是码格）。
    """
    txt = cell.strip()
    if not txt:
        return []
    if CJK.search(txt):
        if not allow_merged:
            return []
        m = re.match(r"^([A-Z]{2,3})\s+\S", txt)
        return [m.group(1)] if m else []
    codes: list[str] = []
    for piece in re.split(r"[、,，/\s]+", txt):
        piece = piece.strip("()（）*")
        if not piece:
            continue
        m = re.match(r"^([A-Z]{1,3})\d{0,2}$", piece)
        if not m:
            return []                       # 混入非码内容 → 整格不作为码格
        codes.append(m.group(1))
    return codes


def header_code_columns(rows: list[list[str]]) -> set[int]:
    """表头里含「码/编码」字样的列号集合（这些列允许「码+空格+中文」的合并写法）。"""
    for row in rows[:3]:
        cols = {i for i, c in enumerate(row) if any(h in c for h in CODE_HEAD)}
        if cols:
            return cols
    return set()


def code_cells(row: list[str], merged_cols: set[int]) -> list[tuple[int, list[str]]]:
    """行内所有码格 [(列号, [码…])]。"""
    found = []
    for j, cell in enumerate(row):
        cs = codes_in(cell, allow_merged=j in merged_cols)
        if cs:
            found.append((j, cs))
    return found


def pick_name(row: list[str], code_col: int, has_header: bool,
              rows: list[list[str]], kind: str) -> str:
    """取名称。有表头 → 按名称列；续表 → 优先紧邻左侧中文，否则右侧。"""
    if has_header:
        nc = name_column(rows, code_col)
        if nc is not None and nc < len(row) and CJK.search(row[nc]):
            return row[nc]
    for j in range(code_col - 1, -1, -1):          # 先往左（正文表「序号|名称|编码原则|说明」的名称在码左侧）
        if CJK.search(row[j]):
            return row[j]
    for j in range(code_col + 1, len(row)):        # 再往右（续表「码|名称|范围|英文名」的名称在码右侧）
        if CJK.search(row[j]):
            return row[j]
    return ""


collected: dict[str, dict] = {}      # code -> {cn, kind, kinds[], note, source, alt_cn}
by_kind: dict[str, dict[str, str]] = {"system": {}, "device": {}}   # kind -> code -> cn
defects: list[dict] = []
stats = {"hdr_system": 0, "hdr_device": 0, "cont": 0}
tables_used: list[tuple[int, str, int]] = []
cur_kind: str | None = None

for i, tb in enumerate(doc.tables):
    rows = [clean([c.text for c in r.cells]) for r in tb.rows]
    rows = [r for r in rows if any(r)]
    if not rows:
        continue

    hk = header_kind(rows)
    if hk != "unknown":
        kind = hk
        cur_kind = kind
        stats["hdr_system" if kind == "system" else "hdr_device"] += 1
        has_header = True
    else:
        # 续表：继承上一张有表头表的 kind（不再用"首格字母"兜底——那条规则会误判附录B续表）
        kind = cur_kind or "unknown"
        stats["cont"] += 1
        has_header = False

    # 「码+空格+中文」的合并写法只在**表头认定的码列**上打开（防名称列误抽 MIS/DCS/SIS）
    merged_cols = header_code_columns(rows) if has_header else set()
    got = 0
    for row in rows:
        for j, cs in code_cells(row, merged_cols):
            for code in cs:
                # 长度硬校验：3 字母必为系统码；设备分类码固定 2 字母
                if CODE3.match(code):
                    row_kind = "system"
                elif CODE2.match(code) or CODE1.match(code):
                    row_kind = kind
                else:
                    continue
                if row_kind not in by_kind:
                    # 尚无任何"有表头"表可继承（文档开头的非分类码表），无法判段位 → 不收
                    continue
                name = pick_name(row, j, has_header, rows, kind)
                if not name:
                    continue
                name = name[:60]
                # ⚠️ 同一个码可**同时**是系统码与设备码（段位不同、含义不同，
                #    如 `PA`：S₁S₂S₃ 位=主过程冷却水系统1，T₁T₂ 位=固定摄像机）。
                #    因此**按 kind 分别累积**，绝不能"保留名称更长的"跨 kind 覆盖
                #    ——那会把设备语义（如 `PC` 微距相机）盖成系统语义（后续标准预留）。
                got += 1
                prev = by_kind[row_kind].get(code)
                if prev is None or len(name) > len(prev):
                    by_kind[row_kind][code] = name
                ent = collected.setdefault(code, {"cn": name, "kind": row_kind,
                                                 "kinds": [], "note": "",
                                                 "source": f"中煤细则表{i + 1}", "alt_cn": ""})
                if row_kind not in ent["kinds"]:
                    ent["kinds"].append(row_kind)
                # 主 kind 取 device（更具体、判码时更相关）；两义并存时把另一义存进 alt_cn
                if row_kind == "device":
                    if ent["kind"] != "device":
                        ent["alt_cn"] = ent["cn"]
                    ent["cn"], ent["kind"] = name, "device"
                    ent["source"] = f"中煤细则表{i + 1}"
                elif ent["kind"] != "device":
                    if len(name) > len(ent["cn"]):
                        ent["cn"] = name
                        ent["source"] = f"中煤细则表{i + 1}"
                elif not ent["alt_cn"] or len(name) > len(ent["alt_cn"]):
                    ent["alt_cn"] = name
        # 登记非法形态：整格就是一个 ≥4 字母的大写串（如 METS / MTSI）
        for cell in row:
            tok = cell.strip().strip("()（）*")
            if re.fullmatch(r"[A-Z]{4,}", tok):
                defects.append({"code": tok, "table": i + 1,
                                "why": "长度 >3；系统码固定 3 字母、设备码固定 2 字母（4.4.2/4.4.3）"})
    if got:
        tables_used.append((i + 1, kind, got))

sys_zm = by_kind["system"]
dev_zm = by_kind["device"]
both = sorted(set(sys_zm) & set(dev_zm))

print(f"表格总数 {len(doc.tables)}：有表头 系统{stats['hdr_system']} / 设备{stats['hdr_device']}"
      f" ｜ 无表头续表 {stats['cont']}")
print(f"抽出分类码 {len(collected)} 条（去重后），来自 {len(tables_used)} 张表")
print(f"  中煤系统码 {len(sys_zm)} 条 ｜ 中煤设备码 {len(dev_zm)} 条 ｜ 两义并存 {len(both)} 条")
print("  中煤系统码按长度:", {n: sum(1 for c in sys_zm if len(c) == n) for n in (1, 2, 3)})
print("  中煤设备码按长度:", {n: sum(1 for c in dev_zm if len(c) == n) for n in (1, 2, 3)})
if both:
    print(f"  ⚠️ 两义并存（同一码既是系统码又是设备码，须按段位判读）: {both[:20]}"
          f"{' …' if len(both) > 20 else ''}")

(CFG / "zhongmei_codes.json").write_text(
    json.dumps(collected, ensure_ascii=False, indent=1), encoding="utf-8")
(CFG / "zhongmei_defects.json").write_text(
    json.dumps(defects, ensure_ascii=False, indent=1), encoding="utf-8")
print(f"  登记非法形态 {len(defects)} 条 → config/zhongmei_defects.json: "
      f"{sorted({d['code'] for d in defects})}")

# ---------- 合并成项目可用码表 ----------
gb_sys = json.loads((CFG / "gb_system_codes.json").read_text(encoding="utf-8"))
gb_dev = json.loads((CFG / "gb_device_codes.json").read_text(encoding="utf-8"))

proj_sys = dict(gb_sys)
proj_dev = dict(gb_dev)


def add(proj: dict, table: dict[str, str], kind: str) -> int:
    """把中煤码并入项目字典（已存在的不覆盖）。返回新增条数。"""
    n = 0
    for code, name in table.items():
        if code not in proj:
            rec = collected.get(code, {})
            proj[code] = {"cn": name, "en": "",
                          "note": rec.get("alt_cn", "") if rec.get("kinds") == ["system"] else "",
                          "source": rec.get("source", "中煤细则")}
            n += 1
    return n


add_sys = add(proj_sys, sys_zm, "system")
add_dev = add(proj_dev, dev_zm, "device")

(CFG / "project_system_codes.json").write_text(
    json.dumps(proj_sys, ensure_ascii=False, indent=1), encoding="utf-8")
(CFG / "project_device_codes.json").write_text(
    json.dumps(proj_dev, ensure_ascii=False, indent=1), encoding="utf-8")

print(f"\n项目可用系统分类码: 国标 {len(gb_sys)} + 中煤 {add_sys} = {len(proj_sys)}")
print(f"项目可用设备分类码: 国标 {len(gb_dev)} + 中煤 {add_dev} = {len(proj_dev)}")

# ---------- 派生产物：独创码 / 禁用码 ----------
# 独创码 = 中煤有、国标**同段位**没有的码（按段位分别判定，不跨段位）
unique: dict[str, dict] = {}
for code, name in sys_zm.items():
    if code not in gb_sys:
        unique[code] = {"cn": name, "kind": "system",
                        "source": collected.get(code, {}).get("source", "中煤细则")}
for code, name in dev_zm.items():
    if code not in gb_dev:
        unique[code] = {"cn": name, "kind": "device",
                        "source": collected.get(code, {}).get("source", "中煤细则")}
(CFG / "zhongmei_unique_codes.json").write_text(
    json.dumps(unique, ensure_ascii=False, indent=1), encoding="utf-8")

# 项目禁用码 = 国标附录A 中 S₁ ∈ {J, K} 的核电分类码（细则 T25 明文"省略"）
excluded = {c: gb_sys[c] for c in gb_sys if c[:1] in ("J", "K")}
(CFG / "zhongmei_excluded_codes.json").write_text(
    json.dumps(excluded, ensure_ascii=False, indent=1), encoding="utf-8")

n_us = sum(1 for v in unique.values() if v["kind"] == "system")
n_ud = sum(1 for v in unique.values() if v["kind"] == "device")
print(f"中煤独创码（国标同段位查不到）: {len(unique)} 条 (系统 {n_us} / 设备 {n_ud})")
print(f"项目禁用码（S₁ ∈ J/K，核电）: {len(excluded)} 条 "
      f"= J {sum(1 for c in excluded if c[0] == 'J')} + K {sum(1 for c in excluded if c[0] == 'K')}")
print(f"  校验: 国标附录A {len(gb_sys)} − 禁用 {len(excluded)} = {len(gb_sys) - len(excluded)}（应等于细则附录A）")

# ---------- ⭐「可用」槽赋义码 ----------
# 国标附录A 有大量 `可用（X 至 Y）` 的**开放槽位**（如 `CQA：可用（CQA 至 CQZ）`）。
# 中煤在这些槽里赋了具体含义（如 CUB=汽水取样系统、UYC=综合专业室外架空管道支架）。
# 这类码**合法但须登记**：它不是"国标外新造码"，而是"占用了国标预留的可用槽"。
# → 按区间解析后判别，替代原先的空壳文件。
avail_ranges: list[tuple[str, str, str]] = []
for _c, _v in gb_sys.items():
    _cn = _v.get("cn") or ""
    if "可用" not in _cn:
        continue
    _m = re.search(r"([A-Z]{1,3})\s*至\s*([A-Z]{1,3})", _cn)
    if _m:
        avail_ranges.append((_m.group(1), _m.group(2), _c))

star: dict[str, dict] = {}
true_new: dict[str, dict] = {}
for code, v in unique.items():
    if v["kind"] != "system":
        continue
    slot = next((r for r in avail_ranges if r[0] <= code <= r[1]), None)
    if slot:
        star[code] = {**v, "slot": f"{slot[0]}~{slot[1]}", "declared_at": slot[2]}
    else:
        true_new[code] = v

(CFG / "zhongmei_star_codes.json").write_text(
    json.dumps({"star": star, "true_new": true_new}, ensure_ascii=False, indent=1),
    encoding="utf-8")
print(f"\n⭐ 国标「可用」开放性槽位解析: {len(avail_ranges)} 段")
print(f"  中煤在**可用槽内**赋义（合法但须登记）: {len(star)} 条")
print(f"  中煤**完全在国标之外**新造: {len(true_new)} 条 → {sorted(true_new)}")
print("  → config/zhongmei_star_codes.json  {star: {...}, true_new: {...}}")

# ---------- 逐码差异明细（追溯用；此前是孤立产物，易与字典漂移，故在此一并生成）----------
def name_diff(left: dict[str, str], right: dict[str, str]) -> list[dict]:
    out = []
    for c in sorted(set(left) & set(right)):
        if left[c] != right[c]:
            out.append({"code": c, "gb": left[c][:80], "zm": right[c][:80]})
    return out


diff = {
    "gb_sys_total": len(gb_sys), "gb_dev_total": len(gb_dev),
    "proj_sys_total": len(proj_sys), "proj_dev_total": len(proj_dev),
    # 中煤带来的增量（= 项目可用集 减去 国标附录）
    "sys_only_zm": sorted(set(proj_sys) - set(gb_sys)),
    "dev_only_zm": sorted(set(proj_dev) - set(gb_dev)),
    # 项目**禁用**的国标码（核电 J/K）。`项目可用集` 是 `国标 ∪ 中煤`，所以"国标有、项目无"
    # 恒为空集；真正要追溯的是这个**主动剔除**集，即 S₁ ∈ {J, K}。
    # 实测 393 = J 204 + K 189（含 `J`、`K` 两个 1 字母主组码本身）；中煤附录A 里
    # 完全没有 J/K/A/I/O 这五个主组。→ 3 372 − 393 = 2 979 就是细则附录A 的规模。
    "sys_excluded_jk": sorted(excluded),
    # 同名不同义（两套文档对同一码的中文表述差异，追溯用）
    "sys_name_diff": name_diff({c: v["cn"] for c, v in gb_sys.items()},
                               {c: v for c, v in sys_zm.items()}),
    "dev_name_diff": name_diff({c: v["cn"] for c, v in gb_dev.items()},
                               {c: v for c, v in dev_zm.items()}),
}
(CFG / "gb_vs_zhongmei_diff.json").write_text(
    json.dumps(diff, ensure_ascii=False, indent=1), encoding="utf-8")
print(f"\n逐码差异明细 → config/gb_vs_zhongmei_diff.json")
print(f"  项目可用集 = 国标 ∪ 中煤：系统 {len(gb_sys)} ∪ {len(diff['sys_only_zm'])} = {len(proj_sys)}"
      f" ｜ 设备 {len(gb_dev)} ∪ {len(diff['dev_only_zm'])} = {len(proj_dev)}")
print(f"  项目禁用（核电 J/K）: {len(diff['sys_excluded_jk'])} 条"
      f" ｜ 同名不同义: 系统 {len(diff['sys_name_diff'])} / 设备 {len(diff['dev_name_diff'])}")
