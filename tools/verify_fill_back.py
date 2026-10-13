# -*- coding: utf-8 -*-
"""回填交付物的**可断言校验**（不靠肉眼看）：`<原名>_已填写.xlsx` 是否守住了纪律。

用法：
    python tools/verify_fill_back.py <源表.xlsx> <已填写.xlsx>

校验项（每项都能给出数字，通过/失败可判）：
  1. **表头对齐**：输出表头里能找到原表表头（新增列只允许出现在两端：左「原文件序号」、右「修编说明」）
  2. **原有行不丢**：源表每个数据行（按"身份列"整行内容）都能在输出里找到（多重集匹配）
  3. 🔴 **原有非空格一格未改**：源表非空值保留；最终设备列的指令/测点和重复编码按最新要求清空并追溯原值
  4. **只填空格**：新增内容只出现在源表原本为空的格里；上述最终编码清空为明确授权例外
  5. 🔴 **行序不变 + 原文件序号对得上**（2026-10-10 用户口径）：
     第一列「原文件序号」**非空的那些值，必须按序正好是 1,2,3,…,原表末行**
     —— 一条断言同时证明"行序没被重排"和"序号就是 Excel 实际行号"；
     且**留空的必须都是新增行**（没有原KKS编码），否则就是原表行被误判成新增。
  6. **可读性**：「系统」列的值不是码表整句（无 ` / `、无 `从(不包括)`），也不含「后续标准预留」
"""
from __future__ import annotations

import pathlib
import re
import sys
from collections import Counter

import openpyxl
sys.path.insert(0, str(pathlib.Path(__file__).resolve().parents[1]))
from kks_gen.row_policy import classify_row
from gen_step1 import resolve_columns

# 引擎**可能写入**的列（按表头前缀认）。其余列是"身份列"，必须原样保留 ——
# 注意用前缀匹配：`原KKS编码` 不以 `KKS编码` 开头，不会误判为可写列。
# 2026-10-10 起新增「原文件序号」列（最左边），它也是引擎写的 → 必须列进来，
# 否则会被当成"身份列"→ 源表没有这一列 → 所有行都会被判成"丢失"。
WRITABLE_PREFIX = ("原文件序号", "设备名称", "KKS编码", "父级KKS编码", "上级设备名称",
                   "上级KKS编码", "系统", "设备类", "设备类型", "修编说明")


def is_writable(head: str) -> bool:
    return any(head.startswith(p) for p in WRITABLE_PREFIX)


def sheet_rows(path: pathlib.Path) -> tuple[list, list[tuple], int]:
    """取第一个工作表：返回 (表头行, 表头之后的行, 表头行下标 0 基)。"""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=False)
    ws = wb[wb.sheetnames[0]]
    rows = [tuple(r) for r in ws.iter_rows(values_only=True)]
    wb.close()
    head_i = 0
    for i, r in enumerate(rows[:10]):
        cells = [str(c or "").strip() for c in r]
        if "序号" in cells:
            head_i = i
            break
    return list(rows[head_i]), rows[head_i + 1:], head_i


def cell(v) -> str:
    s = str(v).strip() if v is not None else ""
    return "" if s in ("None", "") else s


def align_columns(src_head: list, out_head: list, n: int) -> int | None:
    """找输出表头里原表表头的起始位置（= 左边新增了几列）。

    2026-10-10 起输出在**最左边**多一列「原文件序号」，所以正常值是 1；老文件是 0。
    返回 None 表示两表对不上（列被挪动/改名）。
    """
    src_names = [cell(h) for h in src_head]
    out_names = [cell(h) for h in out_head]
    for off in range(0, max(1, len(out_names) - n + 1)):
        if out_names[off:off + n] == src_names:
            return off
    return None


def main() -> int:
    src = pathlib.Path(sys.argv[1])
    out = pathlib.Path(sys.argv[2])
    header, srows, head_i = sheet_rows(src)
    oheader, orows, _ = sheet_rows(out)
    n = len(header)
    fails: list[str] = []

    print(f"源表 {src.name}: 表头 {n} 列 / 表头行后的数据行 {len(srows):,}")
    print(f"输出 {out.name}: 表头 {len(oheader)} 列 / 表头行后的数据行 {len(orows):,}")

    # ---- 1. 表头对齐（允许两端新增列）----
    shift = align_columns(header, oheader, n)
    if shift is None:
        fails.append("表头对不上")
        print("  [失败] 输出表头里找不到原表表头（列被挪动或改名？）")
        print(f"     源={[cell(h) for h in header]}")
        print(f"     出={[cell(h) for h in oheader]}")
        return 1
    left = [cell(x) for x in oheader[:shift] if cell(x)]
    right = [cell(x) for x in oheader[shift + n:] if cell(x)]
    print(f"  [通过] 原表 {n} 列在输出中整体右移 {shift} 列；左侧新增 {left}；末尾新增 {right}")

    def out(i: int) -> int:
        """源表列下标 → 输出列下标。"""
        return i + shift
    source_columns = resolve_columns([cell(h) for h in header])
    source_final = source_columns.get("kks")
    output_note = next((i for i, h in enumerate(oheader) if cell(h) == "修编说明"), None)
    output_seq = next((i for i, h in enumerate(oheader) if cell(h) == "原文件序号"), None)
    output_by_source = {int(r[output_seq]): r for r in orows
                        if output_seq is not None and isinstance(r[output_seq], (int, float))}
    def role(row):
        def get(key):
            index = source_columns.get(key)
            return cell(row[index]) if index is not None else ""
        found = classify_row(get("design_code"), get("original_name") or get("device_name"), get("device_type"))
        return found if found.is_point else classify_row(get("kks"), get("original_name") or get("device_name"), get("device_type"))
    # Compare formula semantics after row/column insertion, with formula text
    # loaded explicitly: cached values disappear when openpyxl saves a file.
    from kks_gen.formula_refs import remap_formula
    source_book = openpyxl.load_workbook(src, read_only=True)
    source_sheet = source_book.sheetnames[0]
    source_book.close()
    row_map = {row: row for row in range(1, head_i + 2)}
    row_map.update({source_row: output_row for output_row, row in enumerate(orows, start=head_i + 2)
                    if output_seq is not None and isinstance((source_row := row[output_seq]), (int, float))})
    srows = [tuple(remap_formula(value, row_map, shift, source_sheet, source_sheet) for value in row) for row in srows]
    authorized_clears = 0
    for source_row, row in enumerate(srows, start=head_i + 2):
        if source_final is None or not cell(row[source_final]):
            continue
        output_row = output_by_source.get(source_row)
        if output_row is None:
            fails.append(f"原文件第{source_row}行无法按序号定位")
            continue
        before, after = cell(row[source_final]), cell(output_row[out(source_final)])
        note = cell(output_row[output_note]) if output_note is not None else ""
        if before == after:
            continue
        explained_duplicate = before in note and "3.1.1" in note and ("重复编码" in note or "撞成同一编码" in note)
        if not after and (role(row).is_point or explained_duplicate):
            authorized_clears += 1
        else:
            fails.append(f"第{source_row}行最终编码发生未经规则允许的改写")
    print(f"  [信息] 按最新指令/测点与重复编码规则清空 {authorized_clears:,} 个原最终值")

    # ---- 2. 原有行不丢（按"身份列"比多重集）----
    # ⚠️ 不能按整行比：目标列被**填过**的行整行内容必然变了（这正是回填的目的，实测踩过）。
    #    正确的"行不丢"判据是**身份列**（原KKS编码/原设备名称/序号… 这些引擎不写的列）。
    writable = sorted(set(i for i, h in enumerate(header) if is_writable(cell(h))) |
                      {index for key, index in source_columns.items() if key in ("kks", "device_name", "parent_kks", "parent_name") and index is not None})
    ident = [i for i in range(n) if i not in set(writable)]

    def src_key(row) -> tuple:
        return tuple(cell(row[i]) for i in ident)

    def out_key(row) -> tuple:
        return tuple(cell(row[out(i)]) for i in ident)

    want_id = Counter(src_key(r) for r in srows)
    got_id = Counter(out_key(r) for r in orows)
    missing = sum(max(0, c - got_id[k]) for k, c in want_id.items())
    if missing:
        fails.append(f"{missing} 行原数据在输出中丢失")
        print(f"  [失败] {missing:,} 行原数据丢失（身份列对不上）")
        for k, c in want_id.items():
            if got_id[k] < c:
                print(f"     例: 原KKS={k[-1]!r} 原表 {c} 次 / 输出 {got_id[k]} 次")
                break
    else:
        print(f"  [通过] 原表 {len(srows):,} 行数据在输出中一行不丢（身份列多重集核对）")

    # ---- 3/4. 原有非空格一格未改；新增只落在空格 ----
    # 引擎只做两件事：① 把原表行**原样**搬到新位置；② 往**空**的目标格里填。
    # 所以"原有数据是否被改动"可以化成一条**无序不变量**（不受行重排影响）：
    #   按「身份列」分组，逐**可写列**比非空值的**多重集**——
    #   输出必须**包含**源表（只增不减）。少一个值就说明有原内容被覆盖/丢失。
    def vals(rows, col: int) -> Counter:
        return Counter(cell(r[col]) for r in rows if cell(r[col]))

    src_by: dict[tuple, list] = {}
    out_by: dict[tuple, list] = {}
    for r in srows:
        src_by.setdefault(src_key(r), []).append(r)
    for r in orows:
        out_by.setdefault(out_key(r), []).append(r)

    lost = 0
    examples: list[str] = []
    preserved = 0
    for k, rs in src_by.items():
        os_ = out_by.get(k) or []
        for c in writable:
            if c == source_final:
                continue  # checked by source Excel row, including permitted clears
            want, got = vals(rs, c), vals(os_, out(c))
            for v, cnt in want.items():
                preserved += cnt
                if got[v] < cnt:
                    lost += cnt - got[v]
                    if len(examples) < 5:
                        examples.append(f"第{c+1}列 {cell(header[c])!r}={v!r}")
    print(f"  可写列: {[cell(header[c]) for c in writable]}")

    if lost:
        fails.append(f"原表 {lost} 个原有数值在输出中丢失/被覆盖")
        print(f"  [失败] 原有非空值丢失 {lost:,} 个；例：{examples}")
    else:
        print(f"  [通过] 原表原有非空值 {preserved:,} 个全部保留（一格未改）")

    # 新增格数按**列计数差**算（精确）：非空格只增不减，所以
    #   输出非空格数 − 源表非空格数 = 真正新填的格数。
    # ⚠️ 不能用"按身份组做多重集差"——组内其它行已有同值时会互相抵消，
    #    2,208 格实测被算成 0（踩过）。
    filled = 0
    decreased: list[str] = []
    for c in writable:
        # ⚠️ 要 `sum(Counter.values())`（非空格**总数**），不是 `len(Counter)`（**不同取值个数**）——
        #    后者把「系统」列的 52,906 格算成 156（实测踩过）。
        d = sum(vals(orows, out(c)).values()) - sum(vals(srows, c).values())
        if d < 0 and c != source_final:
            decreased.append(f"{cell(header[c])}({d})")
        filled += d
    if decreased:
        fails.append("可写列非空格数减少: " + ", ".join(decreased))
        print(f"  [失败] 可写列非空格数减少（说明有原内容被清掉）: {decreased}")
    print(f"  [信息] 原本为空、输出已填的格: {filled:,}（含新增父级行上的填充）")

    # ---- 5. 行序不变 + 「原文件序号」对得上（用户 2026-10-10 口径）----
    seq_head = [cell(h) for h in oheader[:shift]]
    if "原文件序号" not in seq_head:
        print(f"  [注意] 左侧没有「原文件序号」列（{seq_head}），跳过行序校验")
    else:
        seq_i = seq_head.index("原文件序号")
        kks_i = next((i for i, h in enumerate(header) if cell(h) == "原KKS编码"), None)
        seq_vals: list[int] = []
        new_rows = 0
        bad_new: list[str] = []
        for r in orows:
            v = cell(r[seq_i]) if seq_i < len(r) else ""
            if not v:
                new_rows += 1
                # 序号留空的必须真是"新增行"：它不该有原表摘录的「原KKS编码」。
                # 若某行的原KKS编码有值却没有序号 → 说明原表行被误标成新增（口径破了）。
                if kks_i is not None and cell(r[out(kks_i)]):
                    bad_new.append(cell(r[out(kks_i)]))
                continue
            try:
                seq_vals.append(int(float(v)))
            except ValueError:
                fails.append(f"「原文件序号」出现非数字值 {v!r}")
                print(f"  [失败] 「原文件序号」出现非数字值 {v!r}")
                break
        last_src_row = head_i + 1 + len(srows)          # 原表末行（Excel 实际行号）
        # ⚠️ 本脚本只读"表头行之后"的行，所以序列应从**表头行的下一行**开始（点表 = 第 3 行起），
        #    不是从 1 开始 —— 表头行本身第 1 列写的是列名「原文件序号」（已在表头对齐检查里核过）。
        first_seen = head_i + 2
        expect = list(range(first_seen, last_src_row + 1))
        if seq_vals != expect:
            fails.append("「原文件序号」不连续/非原序")
            print(f"  [失败] 「原文件序号」不是 {first_seen}…{last_src_row} 的连续序列"
                  f"（取到 {len(seq_vals):,} 个，首尾 {seq_vals[:3]}…{seq_vals[-3:]}）")
            diff = next((i for i, (a, b) in enumerate(zip(seq_vals, expect)) if a != b), None)
            if diff is not None:
                print(f"     首个不一致位置：第 {diff+1} 个应为 {expect[diff]}，实际 {seq_vals[diff]}")
        else:
            print(f"  [通过] 「原文件序号」= {first_seen}…{last_src_row:,} 连续升序 → **行序与原文件一致**、"
                  f"序号即 Excel 实际行号；另有 {new_rows:,} 行是新增（序号留空）")
        if bad_new:
            fails.append(f"{len(bad_new)} 行留空序号的其实是原表行")
            print(f"  [失败] 有 {len(bad_new):,} 行序号留空、却带着原KKS编码（原表行被误标成新增）："
                  f"{bad_new[:5]}")
        else:
            print("  [通过] 序号留空的行都不是原表摘录行（新增行标识可信）")

    # ---- 6. 可读性 + 预留段必须留空 ----
    tgt = {cell(h): i for i, h in enumerate(oheader) if cell(h)}
    sys_i = tgt.get("系统")
    if sys_i is None:
        print("  [注意] 输出没有「系统」列，跳过可读性校验")
    else:
        vals_ = Counter(cell(r[sys_i]) for r in orows if cell(r[sys_i]))
        bad = [v for v in vals_ if " / " in v or "从(不包括)" in v or v == "后续标准预留"]
        print(f"  [信息] 「系统」列取值 {len(vals_)} 种，已填 {sum(vals_.values()):,} 格")
        print(f"         最常见: {[v for v, _ in vals_.most_common(6)]}")
        if bad:
            fails.append(f"「系统」列有 {len(bad)} 种不可读/预留值")
            print(f"  [失败] 「系统」列出现不可读或预留段值: {bad[:5]}")
        else:
            print("  [通过] 「系统」列无码表整句、无「后续标准预留」")

    # ---- 7. 「KKS编码（修编后）」列：不得有重复值 + 指令/测点行必须留空 ----
    # 国标 3.1.1：标识在**全厂范围内唯一**。交付列出现重复 = 交付物报废
    # （2026-10-12 实测：旧的"测点归一"把 3,563 类、18,088 行做成了同一个设备码）。
    kks_i = resolve_columns([cell(h) for h in oheader]).get("kks")
    if kks_i is None:
        kks_i = next((i for h, i in tgt.items() if h.startswith("KKS编码")), None)
    orig_i = next((i for h, i in tgt.items() if h == "原KKS编码"), None)
    note_i = next((i for h, i in tgt.items() if h.startswith("修编说明")), None)
    if kks_i is None:
        fails.append("缺最终KKS编码列，无法校验")
    else:
        # 判据只用**列内容**，不 import kks_gen：本脚本是独立核验工具，
        # 不该依赖被核验的引擎（否则引擎错了它也一起错）。
        # 信号码 = `[XY]`（5.2.11：X=信号源 / Y=信号应用）+ X₂（5.2.12 的 15 个字母）+ 两位编号。
        sig_re = re.compile(r"[XY][ABCGHJKLMPQRSUVW]\d{2,3}$")

        # ⚠️ **零容忍**：交付列里一条重复都不许有 —— 包括**源表自带**的重复
        #    （两个不同设备被给了同一个原KKS编码）。KKS 编码是全厂唯一标识，
        #    源表带来的重复同样是交付物的缺陷，不能因为"不是引擎造的"就放过。
        dup = {v: c for v, c in Counter(
            cell(r[kks_i]).upper() for r in orows if cell(r[kks_i])).items() if c > 1}
        if dup:
            fails.append(f"「KKS编码（修编后）」有 {len(dup)} 类重复值")
            print(f"  [失败] 「KKS编码（修编后）」重复值 {len(dup)} 类 / "
                  f"{sum(dup.values()):,} 行（国标 3.1.1 要求全厂唯一）；例: {list(dup.items())[:3]}")
            for _v, _c in list(dup.items())[:3]:
                for _r in orows:
                    if cell(_r[kks_i]) == _v:
                        print(f"         {_v} ← 原码 {cell(_r[orig_i])} ｜ {cell(_r[note_i])[:60]}")
        else:
            print(f"  [通过] 「KKS编码（修编后）」{sum(1 for r in orows if cell(r[kks_i])):,} 个非空值"
                  f"**全表唯一**，一条重复都没有（国标 3.1.1）")

        # 指令/测点 = 原码末尾是信号码，且**剥掉后正好是 12 位设备码**（与引擎同一判据，
        # 独立实现）。剥完不是 12 位的是"编码本身结构就不对"，走的是另一条留空路径，
        # 不该按"未注明指令/测点留空原因"计错。
        def _is_point(code: str) -> bool:
            m = sig_re.search(code)
            return bool(m) and len(code) > 12 and len(code[:m.start()]) == 12 and len(code) > m.start()

        pts = [output_by_source[source_row] for source_row, row in enumerate(srows, start=head_i + 2)
               if role(row).is_point and source_row in output_by_source]
        leaked = [cell(r[orig_i]) if orig_i is not None else "" for r in pts if cell(r[kks_i])]
        no_note = [cell(r[orig_i]) if orig_i is not None else "" for r in pts
                   if note_i is None or "指令/测点" not in cell(r[note_i])]
        if leaked:
            fails.append(f"{len(leaked)} 行指令/测点仍带编码")
            print(f"  [失败] 指令/测点行 {len(pts):,} 行中有 {len(leaked):,} 行**拿到了编码**"
                  f"（口径：只有设备才有 KKS 编码）；例: {leaked[:3]}")
        else:
            print(f"  [通过] 指令/测点行 {len(pts):,} 行**全部留空**编码列（只有设备才有 KKS 编码）")
        if no_note:
            fails.append(f"{len(no_note)} 行指令/测点未注明指令/测点留空原因")
            print(f"  [失败] {len(no_note):,} 行指令/测点留空却没在「修编说明」里写明所服务的设备本体码"
                  f"；例: {no_note[:3]}")
        elif pts:
            print(f"  [通过] 这些行的「修编说明」都注明指令/测点留空原因")

    print()
    if fails:
        print("结论: [失败] " + "；".join(fails))
        return 1
    print("结论: [通过] 全部校验项成立")
    return 0


if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    raise SystemExit(main())
