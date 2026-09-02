# -*- coding: utf-8 -*-
"""
KKS 编码成果审核 · 通用模板 (inspect -> probe -> check -> diag -> gen)
==============================================================
中煤多厂 KKS 审核复用骨架。覆盖八维框架 + 用户两批命名/完整性校验
（#24-26 名+完整性三盲区；#1/#27-#31 术语/文本卫生/机组三方/OCR/父子语义）。

用法：
    python audit_template.py <编码.xlsx> [--sheet 主表名]

列定位：按表头"名"动态定位，不写死索引（ADAPT HERE 处按厂微调）。
openpyxl 在托管 venv：
    C:\\Users\\Matt\\workbuddy\\binaries\\python\\envs\\default\\Scripts\\python.exe

本文件是"骨架 + 示例函数"，每个函数独立可跑；跑具体厂时复制函数到该厂的
inspect_xx.py / probe_xx.py / check_xx.py / diag_xx.py / gen_xx_report.py，
并按 ADAPT HERE 调列即可。
"""
import openpyxl, re, sys, itertools
from collections import Counter, defaultdict

# ============ 0. 列定位（按表头名，勿用固定索引） ============
# ADAPT HERE: 不同厂表头文案略有差异，按需加关键字
def locate_columns(header):
    """返回 dict: kks / old / parent / name / dtype 的列索引(0-based)，找不到=None"""
    def find(*keys, exclude=()):
        for i, h in enumerate(header):
            if h is None:
                continue
            s = str(h).strip()
            if any(x in s for x in exclude):
                continue
            for k in keys:
                if k in s:
                    return i
        return None
    return {
        # 新 KKS：排除"原/上级/设计/自编"等衍生列
        'kks':    find('KKS编码', 'KKS码', '新KKS', exclude=('原','上级','设计','自编','自定')),
        # 原 KKS：直接匹配"原KKS/旧KKS"，不排"原"（本就是原码列）
        'old':    find('原KKS', '旧KKS'),
        'parent': find('父级', '上级'),
        # 设备名称：优先"必填/新"名称列，排除"（原）/原/机组/系统/上级"
        'name':   (find('设备名称（必填）','名称（必填）','新名称','新设备名称')
                   or find('名称', '设备名称', exclude=('机组','系统','上级','原','（原)','(原)'))),
        'dtype':  find('设备类型', '类型'),
        'unit':   find('机组', '机组号', exclude=('名称',)),  # #28 机组三方一致
    }

def read_rows(fp, sheet=None, header_row=0, data_start=2):
    """读 sheet；返回 (cols_dict, rows_list)。多 sheet 时若不指定 sheet 取最大表。"""
    wb = openpyxl.load_workbook(fp, read_only=True, data_only=True)
    if sheet is None:
        sheet = max(wb.sheetnames, key=lambda s: (wb[s].max_row or 0))
    ws = wb[sheet]
    rows = list(ws.iter_rows(values_only=True))
    hdr = rows[header_row]
    cols = locate_columns(hdr)
    data = rows[data_start:] if data_start < len(rows) else []
    return cols, data, sheet

# ============ 1. 结构审视 ============
def inspect(fp):
    wb = openpyxl.load_workbook(fp, read_only=True, data_only=True)
    print("sheets:", wb.sheetnames)
    for ws in wb.worksheets:
        rows = list(ws.iter_rows(values_only=True))
        print(f"  [{ws.title}] rows={ws.max_row} cols={ws.max_column}")
        if rows:
            print("    hdr:", [str(c)[:14] if c is not None else None for c in rows[0]])
    # 提示多 sheet 陷阱
    if len(wb.sheetnames) > 1:
        print("⚠ 多 sheet：确认主表（导出表/草稿表常重叠但互不为准）")

# ============ 2. 特征探测 ============
def probe(cols, rows):
    k = cols['kks']; p = cols['parent']; o = cols['old']
    codes, parents = [], set()
    illegal = 0; lens = Counter()
    for r in rows:
        c = r[k] if k is not None else None
        if not c:
            continue
        c = str(c).strip()
        codes.append(c)
        lens[len(c)] += 1
        if re.search(r'[^A-Z0-9\-]', c):
            illegal += 1
        par = r[p] if p is not None else None
        if par:
            parents.add(str(par).strip())
    # 孤儿：父级不在码集中（跳过根哨兵 -1/空）
    orphans = [pp for pp in parents if pp not in ('-1','') and pp not in set(codes)]
    print("总行(有码):", len(codes), "| 非法字符:", illegal)
    print("码长分布:", dict(sorted(lens.items())))
    print("孤儿父级数:", len(orphans), orphans[:10])
    return codes, parents, lens, orphans

# ============ 3. 深度校验 + 语义重分类 ============
def classify_long(code):
    """鄂州订正：>12 码不可全判'信号点号'，须按后缀语义分类。返回类别。"""
    if re.search(r'-KF\d{2}$', code):
        return 'DCS点号(-KF)'
    if re.search(r'-[A-Z]{2}\d{2}$', code):
        return '部件级附加码(-XXnn)'
    if re.search(r'[A-Z]{2}\d{3}[ABC]$', code):
        return '末位A/B/C(分相/变体)'
    return '其它超长'

def old_base(o):
    """原KKS 归并基：剥离尾随位置后缀 A~E（同源多设备的位置变体），再取前9位。"""
    o = str(o).strip()
    b = re.sub(r'[A-Ea-e]$', '', o)
    if b and b != o:
        return b
    return o[:9] if len(o) >= 9 else o

def dup_split(cols, rows):
    """重复码二分：位置后缀丢失(同源多设备合一) vs 异设备真碰撞。
    判定：同新码的多行，其原KKS 前9位(系统+设备类型)全部相同 -> 同源后缀丢失；
    前9位不同(跨系统/跨设备类型) -> 异设备真碰撞。"""
    k = cols['kks']; o = cols['old']
    by_code = defaultdict(list)
    for r in rows:
        c = r[k] if k is not None else None
        if c:
            by_code[str(c).strip()].append(r)
    suffix_loss, collision = [], []
    for c, grp in by_code.items():
        if len(grp) <= 1:
            continue
        bases = set()
        for g in grp:
            if o is not None and g[o]:
                bases.add(old_base(g[o]))
        if len(bases) <= 1:
            suffix_loss.append((c, len(grp)))
        else:
            collision.append((c, len(grp), [str(g[o]).strip() for g in grp if g[o]]))
    return suffix_loss, collision

# ============ 4. 诊断归因（取样本，避误报） ============
def diag_prefix(cols, rows, orphans):
    """对孤儿父级取子节点样本，确认是'可挂上级'还是'旧格式未迁'。"""
    k = cols['kks']; p = cols['parent']
    samples = defaultdict(list)
    for r in rows:
        par = r[p] if p is not None else None
        if par and str(par).strip() in orphans:
            c = r[k] if k is not None else None
            samples[str(par).strip()].append(str(c).strip() if c else '')
    for par, kids in list(samples.items())[:8]:
        print(f"  父缺 {par}: 子样本 {kids[:3]} ...")
        if re.match(r'^(50|60|L0)', par):
            print(f"    -> 旧版前缀(G=5/6/L 合法，仅提示核对，非迁移错误)")

# ============ 5. 用户新增 3 条校验 ============
# ---- #24 命名模糊/歧义 ----
def naming_ambiguity(cols, rows):
    """检测名称中'X号Y'结构：每个数字须紧跟明确修饰名词(机/高加/泵/阀…)。
    返回疑似歧义 (行号, 名称)。"""
    k = cols['kks']; n = cols['name']
    AMBIG = []
    # 简化规则：名称含"号"且其后紧跟另一数字（数字相邻无名词分隔）
    pat = re.compile(r'(\d+)\s*号\s*(\d+)')  # "1号2号..." 相邻数字
    for i, r in enumerate(rows):
        nm = r[n] if n is not None else None
        if not nm:
            continue
        nm = str(nm).strip()
        if pat.search(nm):
            AMBIG.append((i, nm))
    return AMBIG

# ---- #25 同设备异码 / 命名风格不一 ----
SYN = {  # 同义词词典（ADAPT HERE 按厂扩充；长词在前，避免短词先吞）
    '带式输送机':'BELT','皮带机':'BELT','带式':'BELT','皮带':'BELT',
    '截止阀':'VALVE','闸阀':'VALVE','截门':'VALVE','阀门':'VALVE',
    '给水泵':'PUMP','循环水泵':'PUMP','水泵':'PUMP','泵':'PUMP',
    '电动机':'MOTOR','马达':'MOTOR','电机':'MOTOR',
    '送风机':'FAN','引风机':'FAN','风机':'FAN',
    '采样装置':'SAMPL','采样器':'SAMPL','采样头':'SAMPL','采样':'SAMPL',
}
def norm_name(nm):
    s = str(nm)
    for k, v in SYN.items():
        s = s.replace(k, v)
    s = re.sub(r'[ \#]', '', s)  # 去空格/#/风格符
    # 中文数字 -> 阿拉伯（防御性；多数名称用"1号"数字写法，此步仅兜底"一号"式）
    CN_DIGIT = {'一':1,'二':2,'三':3,'四':4,'五':5,'六':6,'七':7,'八':8,'九':9,'十':10}
    s = re.sub(r'[一二三四五六七八九十]',
               lambda m: str(CN_DIGIT.get(m.group(0), m.group(0))), s)
    return s
def near_dup_by_name(cols, rows):
    """同系统(前9位)内，不同编码映射到**完全相同**的归一名 -> 疑似同物多码。

    ★校准要点：归一键必须用完整归一名，**不可截断前缀**。
      太仓实证：截 12 字后"…长吹C1/C2/C3"全落进同一桶，603 组几乎全是误报；
      改全名精确相等后只剩真正重名。"""
    k = cols['kks']; n = cols['name']
    buckets = defaultdict(set)  # (系统前9位, 归一全名) -> {编码}
    for r in rows:
        c = r[k] if k is not None else None
        nm = r[n] if n is not None else None
        if not c or not nm:
            continue
        c, nm = str(c).strip(), str(nm).strip()
        key = norm_name(nm)          # 完整归一名，勿截断
        buckets[(c[:9], key)].add(c)
    sus = [(sysk, key, sorted(codes)) for (sysk, key), codes in buckets.items() if len(codes) > 1]
    return sus  # 提示交人核，可能误报

# ---- #25b 语义层同物异名（无词典相似度聚簇，跨码、只标待核）----
LOC_PAT = re.compile(r'[A-Z]\d+[A-Z]?|\d+[A-Z]|\d+号|\d+')   # 位置/编号 token：C7B、A1、8A、1号、序号

def _ngrams(s, n=2):
    s = str(s)
    if len(s) < n:
        return {s}
    return {s[i:i+n] for i in range(len(s) - n + 1)}

def _jacc(a, b):
    if not a or not b:
        return 0.0
    inter = len(a & b)
    if inter == 0:
        return 0.0
    return inter / (len(a) + len(b) - inter)

def _type_key(code):
    """设备类型字母(第8-9位)；变长时退化到系统字母/机组，作同物聚簇的'同类'锚。"""
    c = str(code).strip()
    if len(c) >= 9:
        return c[7:9]
    if len(c) >= 5:
        return c[2:5]
    return c[:2]

def _block_key(nm):
    """分块键：优先取设备/皮带编号(C7B, A1)这类字母+数字 token；
    退化到机组'1号'；再退化到名称前两字。同位置才比，避免跨设备(如C7B vs C8B)误伤。"""
    ms = LOC_PAT.findall(str(nm))
    for m in ms:
        if re.match(r'[A-Z]', m):
            return m
    return ms[-1] if ms else str(nm)[:2]

def _core(nm):
    """语义聚簇'核心名'：归一 + 去数字 + 机组指代归一(号机/号炉/号机组 -> 号)。
    同物异名经同义词典归并、仅序号/机组写法不同后应相等。"""
    s = norm_name(nm)
    s = re.sub(r'\d+', '', s)
    s = s.replace('号机组', '号').replace('号机', '号').replace('号炉', '号')
    return s

# 设备实例区分标记：真不同设备通常有 A/B/甲/乙/进/出/左/右…；真重复往往没有
DISAMB = re.compile(r'(\d+\s*(?:号|#)|[A-Z甲乙][\s/]?(侧|列|进|出|左|右|上|下)?|(?:一|二|三|1|2|3)次|(?:进|出|左|右|上|下|前|后|首|末)[侧列]?|[甲乙AB][侧区]?)')

def semantic_dup_by_name(cols, rows, thr=0.6):
    """#25b 语义层同物异名（实验性 / 低置信 / 仅人工抽检，不可作自动判重）。

    ⚠️ 方法论红线：KKS 设备名称**本身不唯一定位设备**——同一子系统下大量设备共用基础
    名、只靠编码序号区分。"同名不同码"在 KKS 里是常态(非重复)。因此**纯名称相似度无法
    可靠判重**：'同物异名(录两次)'与'同子系统两台真不同设备(不同叫法)'在数据层面统计不可分。
    本函数只产出**高可疑、需人工凭图号/位置证据裁定**的候选，绝不自动合并/判重。

    严格候选条件（多道精度闸，对应已实证踩过的坑）：
      1. 同子系统：编码前 5 位(机组+系统字母)相同 —— 收敛到同一工艺系统；
      2. 名称是'改写'而非'完全相同'：核心名不同 —— 剔除'同子系统多设备同名'的 10 万级噪声；
      3. 名称中**无设备实例区分标记**(A/B/甲/乙/进/出/左/右…) —— 真不同设备通常有、真重复常无；
      4. 核心名字符 n-gram 相似度 >= thr，且为不同编码。

    为什么这么设计（踩坑实证）：
      - 初版按'核心名相等'判 Tier1：鄂州实测 10万+ 全是'同子系统同名只差序号'(QM031~QM032)，
        因 KKS 名非唯一 -> 误报灾难，已弃；
      - 纯 n-gram 相似度：把'同位置同型号只差序号/A-B侧'打到 0.9+ 淹没真同物 -> 加条件2/3 收敛；
      - 按位置 token 分块(优先 C7B 类)：C7B 与 C8B 真不同设备隔离；同源不同叫法(同 C7B)进同池。
      - 块超 1200 跳过：避免大表 O(n^2) 卡死。

    校准（ADAPT: thr 按厂调，默认 0.6）：太仓/鄂州实测，严格模式把候选从 10万级降到可人审量级，
    且能抓用户示例(皮带机/带式输送机 + 采样器/采样装置 + 1号机/1号，同 C7B、无区分标记)。

    返回：[(i, j, 码i, 码j, 核心名i, 核心名j, 相似度, 共有ngram)] 提示交人核。
    """
    k = cols['kks']; n = cols['name']
    blocks = defaultdict(list)   # (type_key, block_key) -> [(行, 编码, 归一名, 核心名)]
    for i, r in enumerate(rows):
        c = r[k] if k is not None else None
        nm = r[n] if n is not None else None
        if not c or not nm:
            continue
        c, nm = str(c).strip(), str(nm).strip()
        nn = norm_name(nm)
        if len(nn) < 3:
            continue
        ms = LOC_PAT.findall(nm)
        if not ms:
            # 名称内无位置/编号 token（如纯"电机""泵"）：跳过语义聚簇，
            # 此类通用名若有精确同名已由 #25a 覆盖；避免退化成大块 O(n^2) 爆炸。
            continue
        core = _core(nm)
        blocks[(_type_key(c), _block_key(nm))].append((i, c, nn, core))
    cand = []
    for key, items in blocks.items():
        if len(items) < 2:
            continue
        if len(items) > 1200:
            continue
        for (i, ci, ci_nn, ci_core), (j, cj, cj_nn, cj_core) in itertools.combinations(items, 2):
            if ci == cj:
                continue
            if ci[:5] != cj[:5]:
                continue                      # 条件1：必须同子系统
            if ci_nn == cj_nn:
                continue                      # 条件2a：原名完全相同只差编码 -> 同子系统多设备(名非唯一)，噪声，跳
            if DISAMB.search(ci_nn) or DISAMB.search(cj_nn):
                continue                      # 先排除进/出、A/B、甲/乙等实例区分语义，再比较核心名
            if ci_core == cj_core:
                # 原名不同、但核心名(同义归并+去数字+机组归一)相等 -> 典型同物异名(高置信)
                cand.append((i, j, ci, cj, ci_core, cj_core, 1.0, []))
                continue
            sc = _jacc(_ngrams(ci_core), _ngrams(cj_core))
            if sc >= thr:
                shared = sorted(_ngrams(ci_core) & _ngrams(cj_core))
                cand.append((i, j, ci, cj, ci_core, cj_core, round(sc, 3), shared[:8]))
    cand.sort(key=lambda x: -x[6])
    return cand


# ---- #25c 结构身份键去重（可靠，推荐；依赖 1:1 原KKS/设计编码列）----
def identity_dup_by_key(cols, rows, key_col=None, name_col=None):
    """按 1:1 结构身份键找真重复：同一原KKS/设计编码下出现 **>1 个不同 KKS 编码** -> 同一设备被录两次。

    ⚠️ 为什么这是 KKS 语义去重的**可靠**路径，而名称相似度不是：
      KKS 设备名天然非唯一（同屏多回路 FC01~FC10、同系统多馈线 51~54 共用基础名、
      只靠编码后缀区分）。"同名不同码"是常态，纯名称相似度无法把'同物异名(录两次)'
      与'同子系统两台真不同设备'区分开（鄂州实测名称相似度层暴到 10 万级误报）。
      图号/位置标识通常是图纸级键，一张图可以挂很多台设备，不能直接作为设备身份键。
      只有原KKS或旧版设计设备编码等 1:1 键才可进入可靠重复判断。

    参数：
      key_col  结构身份键列名（推荐'原KKS'/'旧KKS'/'设计编码'等 1:1 设备键）。
               若该列缺失，函数返回空并提示改用 semantic_dup_by_name（实验性）。
      name_col 名称列（可选，用于回显不同叫法）。

    返回：[(身份键, [(行, 编码, 名称), ...]), ...] —— 每个身份键下 >1 个不同编码即疑似真重复。
    """
    k = cols['kks']
    kc = cols.get('key') if key_col is None else _col_index(cols, key_col)
    if kc is None:
        # 退化：尝试常见身份键列名
        for cand_name in ('原KKS', '原KKS码', '旧KKS', '旧KKS码', '设计编码', '原设备编码'):
            if cand_name in cols:
                kc = cols[cand_name]; break
    if kc is None:
        return []   # 无身份键列，无法可靠去重
    nc = cols.get('name') if name_col is None else _col_index(cols, name_col)
    from collections import defaultdict as _dd
    groups = _dd(list)
    for i, r in enumerate(rows):
        c = r[k] if k is not None else None
        if not c:
            continue
        keyv = r[kc] if kc is not None else None
        keyv = str(keyv).strip() if keyv is not None else ''
        if not keyv or keyv in ('-', '无', 'NA', 'None'):
            continue
        nm = str(r[nc]).strip() if (nc is not None and r[nc] is not None) else ''
        groups[keyv].append((i, str(c).strip(), nm))
    res = []
    for keyv, items in groups.items():
        codes = set(c for (_, c, _) in items)
        if len(codes) > 1:
            res.append((keyv, items))
    res.sort(key=lambda x: -len(x[1]))
    return res


def _col_index(cols, name):
    """按列名(含模糊包含)取列索引；cols 既可能是 {名:idx} 也可能是本模板约定键。"""
    if name in cols:
        return cols[name]
    for kk, vv in cols.items():
        if isinstance(kk, str) and name in kk:
            return vv
    return None

# ---- #26 应编未编（规则库驱动） ----
# 父设备类型(按KKS设备字母或名称关键词) -> 期望子部件关键词
EXPECT = {
    '开关柜': ['抽屉','回路','开关','断路器'],
    '配电柜': ['抽屉','回路','开关'],
    'MCC': ['回路','开关','抽屉'],
    'PC': ['回路','开关'],
    '阀门': ['执行机构','电动头','位置开关','电磁阀','MA','CA'],
    '泵': ['电机','前置泵'],
    '给水泵': ['前置泵','电机'],
    '变压器': ['开关','中性点','隔离开关'],
}
DEVICE_CODE_LEN = 12   # ADAPT HERE: 设备级码长（变长层级厂如舟山按实际调整）

def expected_missing(cols, rows):
    """父设备存在、但下挂子码名称中无任一期望子部件关键词 -> 提示应编未编。

    ★校准要点：只对**设备级父节点**(len==DEVICE_CODE_LEN)判定。
      系统级节点名常含"泵/风机"等字样（如"05HAG 启动循环泵过冷水系统"），
      按名称关键词命中会把系统当设备，太仓实证 36 条全是此类误报。"""
    k = cols['kks']; n = cols['name']; p = cols['parent']
    parent_names = {}      # 父码 -> 名称
    children_names = defaultdict(list)  # 父码 -> [子名称]
    for r in rows:
        c = r[k] if k is not None else None
        nm = r[n] if n is not None else None
        par = r[p] if p is not None else None
        if not c:
            continue
        c = str(c).strip()
        if nm:
            parent_names[c] = str(nm).strip()
        if par:
            par = str(par).strip()
            if nm:
                children_names[par].append(str(nm).strip())
    hints = []
    for par, kids in children_names.items():
        if len(par) != DEVICE_CODE_LEN:      # 仅设备级父节点，排除系统级误报
            continue
        pnm = parent_names.get(par, '')
        for pt, expect in EXPECT.items():
            if pt in pnm:
                if not any(any(e in kk for e in expect) for kk in kids):
                    hints.append((par, pnm, pt, expect))
                    break
    return hints

# ============ 6. 第二批 命名/完整性/数据卫生 校验 ============

# 设备术语词典（starter，ADAPT HERE 按厂扩充；用于 #27 语义可解性 P2）
EQUIP_TERMS = set('泵 阀 机 风机 电机 电动机 马达 压缩机 换热器 加热器 凝汽器 除氧器 锅炉 汽轮机 发电机 变压器 开关柜 配电柜 皮带机 刮板机 碎煤机 给煤机 磨煤机 空预器 除尘器 脱硫塔 吸收塔 烟囱 管道 容器 罐 箱 执行机构 传感器 变送器 液位计 温度计 压力表 流量计 MCC PC'.split())

# 机组一致性（全厂码 G：1-9 → 1~9 号机；A-G → 10~16 号机）
# 通用默认：20版编码机组位前缀数字 == 名称所写机组号（1↔1号, 2↔2号, A↔10号）
COMMON_CODES = {'J','K','L','M','N','P','Q','R','S','T','U','V','Y'}   # 期别公用 J-R / 多期公用 S-V / 全厂公用 Y
FREE_CODES = {'H','W','X','Z'}                                          # 自由使用（火电厂导则 5.1 表2 注3）
UNIT_OVERRIDE = {}                  # ADAPT: 个别厂 G 取值与名称不一致时显式覆写，例 {'E':14}
MAX_UNIT = 16                       # 机组号上限护栏（A-G 对应 10-16）
CN_NUM = {'一':1,'二':2,'三':3,'四':4,'五':5,'六':6,'七':7,'八':8,'九':9,'十':10}
UNIT_LETTER_MAP = {'A':10,'B':11,'C':12,'D':13,'E':14,'F':15,'G':16}

def _code_unit(g_char):
    """把全厂码 G（1 位）翻译成机组号；无法可靠翻译时返回 None（跳过，不误报）。
    ★护栏：滁州实证——10版残留前缀 '60' 被当成"60号机"，凭空产生 7 条误报。
      G 只取 1 位（1-9/A-G），J-R/S-V/Y 为公用，H/W/X/Z 自由使用。"""
    if not g_char:
        return None
    if g_char in COMMON_CODES or g_char in FREE_CODES:
        return None                      # 公用/自由，跳过
    if g_char in UNIT_OVERRIDE:
        return UNIT_OVERRIDE[g_char]
    if g_char.isdigit():
        v = int(g_char)
        return v if 1 <= v <= 9 else None
    return UNIT_LETTER_MAP.get(g_char.upper())

def name_text_hygiene(cols, rows):
    """#27 名称文本卫生 + 语义可解性。返回 (symbol_issues[P1], semantic_voids[P2])。

    ★校准要点（太仓实证）：
    1) 符号白名单必须含中文标点（，。；：、（）"" 等），否则"给水和蒸汽部分，…"全被误报；
       真正该报的是控制符、方框/替换符 □ ￼ 、异常制表符、私用区字符。
    2) 语义可解性 **不用设备术语词典**（用户已定：不建词典）。词典永远不全，
       "省煤器/水冷壁"这类正词会被判为"不可解"，太仓一次误报 3447 条。
       改用无词典硬规则：无任何中文/字母、纯数字或纯符号、长度<2 —— 这些才是真的不可解。"""
    n = cols['name']
    symbol_issues, semantic_voids = [], []
    # 允许：中日韩汉字、字母数字、常用中英文标点与单位符号
    sym_pat = re.compile(r'[^\u4e00-\u9fffA-Za-z0-9'
                         r'()（）\[\]【】.\-—_/\\,，、;；:：%#+*&~°℃㎡"\'“”‘’!！?？°\s]')
    ctrl_pat = re.compile(r'[\x00-\x1f\x7f\ufffd\uf000-\uf8ff]|□|￼')
    for i, r in enumerate(rows):
        nm = r[n] if n is not None else None
        if not nm:
            continue
        raw = str(nm)
        bad = set(sym_pat.findall(raw))
        if ctrl_pat.search(raw) or bad:
            symbol_issues.append((i, raw[:40], ''.join(sorted(bad)) or '控制符/替换符'))
        s = raw.strip()
        has_cjk = re.search(r'[\u4e00-\u9fff]', s) is not None
        has_alpha = re.search(r'[A-Za-z]', s) is not None
        if len(s) < 2 or (not has_cjk and not has_alpha):
            semantic_voids.append((i, raw[:40]))                     # 语义不可解(P2)
    return symbol_issues, semantic_voids

def unit_consistency(cols, rows):
    """#28 机组三方一致（原B升级）：机组列 ↔ 名称机组指代 ↔ 全厂码 G。"""
    k = cols['kks']; n = cols['name']; u = cols['unit']
    issues = []
    name_unit_pat = re.compile(r'(\d{1,2}|[一二三四五六七八九十])\s*号\s*(机|炉|机组)')
    for i, r in enumerate(rows):
        c = r[k] if k is not None else None
        if not c:
            continue
        c = str(c).strip()
        code_g = c[0]
        nm = r[n] if n is not None else None
        nu = None
        if nm:
            m = name_unit_pat.search(str(nm))
            if m:
                key = m.group(1)
                nu = CN_NUM.get(key, int(key)) if not key.isdigit() else int(key)
        uc = r[u] if u is not None else None
        uc_num = None
        if uc:
            um = re.search(r'\d{1,2}|[一二三四五六七八九十]', str(uc))
            if um:
                key = um.group(0)
                uc_num = CN_NUM.get(key, int(key)) if not key.isdigit() else int(key)
        if (uc and ('公用' in str(uc) or '公共' in str(uc))) or code_g in COMMON_CODES or code_g in FREE_CODES:
            continue                                        # 公用/自由白名单，勿误报
        expect = _code_unit(code_g)
        if expect is None:
            continue                                        # 根/非标准前缀跳过
        if nu is not None and nu != expect:
            issues.append((i, c, f'名称指{nu}号但编码全厂码G={code_g}(应{expect}号)'))
        if uc_num is not None and uc_num != expect:
            issues.append((i, c, f'机组列指{uc_num}号但编码全厂码G={code_g}(应{expect}号)'))
    return issues

# ---- 分段字符类型（KKS 12 位骨架）----
# 位1 全厂码G(字母/数字) | 位2 系统前缀号F0(数字) | 位3-5 系统分类码F1F2F3 | 位6-7 系统编号FN | 位8-9 设备分类码A1A2 | 位10-12 设备编号AN
SEG_12 = [(1, 1, 'ALNUM', '全厂码G'), (2, 2, 'DIGIT', '系统前缀号F0'),
          (3, 5, 'ALPHA', '系统分类码F1F2F3'), (6, 7, 'DIGIT', '系统编号FN'),
          (8, 9, 'ALPHA', '设备分类码A1A2'), (10, 12, 'DIGIT', '设备编号AN')]

def segment_type_check(cols, rows):
    """★P0 位置感知字符类型校验（补"非法字符"盲区）：
    仅查 [^A-Z0-9] 会漏掉「字母O冒充数字0 / 字母I冒充数字1」——它们都是 A-Z。
    本函数按 12 位骨架逐段校验字符类型，数字段出现字母即报错（并给出建议修正）。
    太仓实证：112 条第10位写成字母 O，同系统兄弟码可反证应为 0。"""
    k = cols['kks']
    FIX = {'O': '0', 'o': '0', 'I': '1', 'l': '1', 'i': '1', 'S': '5', 'B': '8', 'Z': '2'}
    res = []
    for i, r in enumerate(rows):
        c = r[k] if k is not None else None
        if not c:
            continue
        c = str(c).strip()
        if len(c) < 12:          # 只对完整设备级码做分段校验（系统级短码跳过）
            continue
        body = c[:12]
        for a, b, want, label in SEG_12:
            seg = body[a-1:b]
            if want == 'DIGIT' and not seg.isdigit():
                bad = [ch for ch in seg if not ch.isdigit()]
                sug = ''.join(FIX.get(ch, ch) for ch in seg)
                res.append((i, c, f'{label}(第{a}-{b}位)="{seg}" 含非数字 {bad}',
                            body[:a-1] + sug + body[b:] if sug != seg else ''))
            elif want == 'ALPHA' and not seg.isalpha():
                res.append((i, c, f'{label}(第{a}-{b}位)="{seg}" 含非字母', ''))
            elif want == 'ALNUM' and not seg.isalnum():
                res.append((i, c, f'{label}(第{a}-{b}位)="{seg}" 含非字母数字', ''))
    return res

def ocr_confusable(cols, rows):
    """#29 OCR 易混字符待核（位置无关的兜底扫描）。
    与 segment_type_check 配合：本函数只报"位置无法判定"的可疑项，避免重复。
    - 硬异常 I/i/O/o：KKS 字母表本就排除 I/O（防与 1/0 混淆），出现即错。
    - 软提示 小写 l：疑为数字 1（大写 L 合法，不报）。
    只标不自动改。"""
    k = cols['kks']
    hard = set('IiOo')
    soft = set('l')
    res = []
    for i, r in enumerate(rows):
        c = r[k] if k is not None else None
        if not c:
            continue
        c = str(c).strip()
        for ch in c:
            if ch in hard:
                res.append((i, c, f'含易混字符 {ch} (硬异常:KKS字母表排除I/O，小写非法)'))
                break
        else:
            for ch in c:
                if ch in soft:
                    res.append((i, c, f'含易混字符 {ch} (疑为数字1,待核)'))
                    break
    return res

def text_hygiene(cols, rows):
    """#30 文本卫生（码+名称）：首尾空白/换行/tab/全角空格/控制符 + 汉字间夹空格。
    ★汉字之间的空格是复制粘贴/PDF 提取的典型残留（太仓"锅 炉"），会破坏名称检索与去重。"""
    k = cols['kks']; n = cols['name']
    issues = []
    ctrl = re.compile(r'[\n\r\t\x00-\x1f\x7f]|\u3000|\xa0')
    cjk_gap = re.compile(r'[\u4e00-\u9fff][ \t\u3000\xa0]+[\u4e00-\u9fff]')
    for i, r in enumerate(rows):
        for label, idx in (('码', k), ('名称', n)):
            v = r[idx] if idx is not None else None
            if v is None:
                continue
            s = str(v)
            if s != s.strip() or ctrl.search(s):
                issues.append((i, label, '首尾空白/控制符', repr(s)[:40]))
            elif label == '名称' and cjk_gap.search(s):
                issues.append((i, label, '汉字间夹空格', repr(s)[:40]))
        # 编码内部任何空白都非法
        cv = r[k] if k is not None else None
        if cv and re.search(r'\s', str(cv).strip()):
            issues.append((i, '码', '编码内含空白', repr(str(cv))[:40]))
    return issues

def clean_text(v):
    """#30 清洗：trim + 去控制符 + 全角空格/NBSP 转半角。可写回清洗后的文件。"""
    if v is None:
        return v
    s = str(v).replace('\u3000', ' ').replace('\xa0', ' ')
    s = re.sub(r'[\n\r\t\x00-\x1f\x7f]', '', s)
    return s.strip()

def parent_child_name_conflict(cols, rows):
    """#31 父子名称语义一致性：父名标识 token 须与子名一致（父C8B、子C7B→矛盾）。限直接父子一级。"""
    k = cols['kks']; n = cols['name']; p = cols['parent']
    parent_name = {}
    for r in rows:
        c = r[k] if k is not None else None
        nm = r[n] if n is not None else None
        if c and nm:
            parent_name[str(c).strip()] = str(nm).strip()
    tok_pat = re.compile(r'[A-Z]\d+[A-Z]')                 # C8B 式"字母+数字+字母"标识（排除A1/A2位置变体）
    issues = []
    for r in rows:
        c = r[k] if k is not None else None
        par = r[p] if p is not None else None
        nm = r[n] if n is not None else None
        if not (c and par and nm):
            continue
        par = str(par).strip(); nm = str(nm).strip(); c = str(c).strip()
        pnm = parent_name.get(par)
        if not pnm:
            continue
        ptoks = tok_pat.findall(pnm)
        if not ptoks:
            continue
        ctoks = tok_pat.findall(nm)
        for pt in ptoks:
            base = re.match(r'([A-Za-z]\d+)', pt).group(1)   # C8
            for ct in ctoks:
                cb = re.match(r'([A-Za-z]\d+)', ct).group(1)
                if cb != base and cb[0] == base[0]:           # 同字母族(都带C)但号不同
                    issues.append((c, pnm, nm, f'父标识{base} vs 子标识{cb}'))
                    break
            else:
                continue
            break
    return issues

# ============ 主函数（示例） ============
if __name__ == '__main__':
    fp = sys.argv[1] if len(sys.argv) > 1 else None
    if not fp:
        print("usage: audit_template.py <xlsx> [--sheet NAME]")
        sys.exit(1)
    sheet = sys.argv[3] if len(sys.argv) > 3 and sys.argv[2] == '--sheet' else None
    inspect(fp)
    cols, rows, sh = read_rows(fp, sheet)
    print("\n[主表]", sh, "列定位:", cols, "数据行:", len(rows))
    codes, parents, lens, orphans = probe(cols, rows)
    sl, col = dup_split(cols, rows)
    print(f"\n重复码: 后缀丢失 {len(sl)} / 真碰撞 {len(col)}")
    seg = segment_type_check(cols, rows)
    print(f"[★P0 分段字符类型] 违规 {len(seg)} 条", seg[:3])
    print("\n[用户新增 #24] 命名歧义样本:", naming_ambiguity(cols, rows)[:5])
    print("[用户新增 #25] 同设备异码疑似:", len(near_dup_by_name(cols, rows)), "组")
    print("[用户新增 #26] 应编未编提示:", len(expected_missing(cols, rows)), "条")
    # 第二批（#1/#27-#31）
    sym, void = name_text_hygiene(cols, rows)
    print(f"\n[#27 名称文本卫生] 非常规符号 {len(sym)} / 语义不可解(P2) {len(void)}")
    print("[#28 机组三方一致] 不一致:", len(unit_consistency(cols, rows)), "条")
    print("[#29 OCR易混字符] 待核:", len(ocr_confusable(cols, rows)), "条")
    print("[#30 文本卫生] 首尾空白/控制符:", len(text_hygiene(cols, rows)), "处")
    print("[#31 父子名称语义] 矛盾:", len(parent_child_name_conflict(cols, rows)), "条")
