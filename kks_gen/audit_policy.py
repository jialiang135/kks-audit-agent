"""Read-only device-table checks; classification evidence never generates a code."""
import re
from functools import lru_cache
from . import identifiers as ids
from .row_policy import classify_row
from .revise import pick_system_code
from .compliance import unit_conflicts

DEVICE_PATTERN = re.compile(r"[0-9]{2}[A-HJ-NP-Z]{3}[0-9]{2}[A-HJ-NP-Z]{2}[0-9]{3}")
STAGES = ("编码格式", "设备类型", "系统归属", "名称语义", "编号合理性", "位置辅助")
# Subject terms and their functional classifications. Generic labels remain broad.
SUBJECTS = {
    # Appendix B explicitly includes solenoid valves under QA. Depending on
    # the identified function, fluid-valve classifications also remain valid.
    "电磁阀": {"QA", "FL", "RM", "QN", "QM"},
    "一次门": {"FL", "RM", "QN", "QM"}, "二次门": {"FL", "RM", "QN", "QM"},
    "调节门": {"QN"}, "关断门": {"QM"}, "电动门": {"FL", "RM", "QN", "QM"}, "气动门": {"FL", "RM", "QN", "QM"},
    "安全阀": {"FL"}, "止回阀": {"RM"}, "逆止阀": {"RM"},
    "调节阀": {"QN"}, "控制阀": {"QN"}, "截止阀": {"QM"}, "闸阀": {"QM"},
    "阀门": {"FL", "RM", "QN", "QM"}, "阀": {"FL", "RM", "QN", "QM"},
    "风机": {"GQ"}, "泵": {"GP"}, "电动机": {"MA"}, "电机": {"MA"}, "马达": {"MA"},
    "管道": {"WP"}, "变压器": {"TA"},
    "压力变送器": {"BP"}, "差压变送器": {"BP", "BF", "BL"}, "压力表": {"BP"}, "压力开关": {"BP"},
    "热电偶": {"BT"}, "热电阻": {"BT"}, "温度计": {"BT"}, "温度变送器": {"BT"},
    "流量计": {"BF"}, "流量变送器": {"BF"}, "液位计": {"BL"}, "液位变送器": {"BL"},
}


@lru_cache(maxsize=100000)
def subject(text):
    matches = [(match.end(), len(term), term, allowed)
               for term, allowed in SUBJECTS.items() for match in re.finditer(re.escape(term), str(text or ""))]
    if not matches:
        return None
    end, _, term, allowed = max(matches, key=lambda item: (item[0], item[1]))
    if re.search(r"(?:跳闸|联锁|跳|启动|停止)[^，,;；]*$", str(text)[:end - len(term)]):
        return None
    if re.search(r"系统|回路|供电|电源|控制柜|控制箱|执行机构|测量装置|连锁|联锁|反馈|状态|连锁位|位置", str(text)[end:]):
        return None
    # A device noun before another object or a measured property is context,
    # not a reliable assertion that the whole row identifies that device.
    tail = str(text)[end:]
    tail = re.sub(r"备用|本体|设备|组件|正常|一|二|三|四|五|六|七|八|九|十|号", "", tail)
    if re.search(r"[\u4e00-\u9fff]", tail):
        return None
    return term, allowed


def is_node(rec, referenced=()):
    code, name, dtype = rec.get("kks_code", ""), rec.get("name", ""), rec.get("device_type", "")
    if subject(name) or subject(dtype):
        return False
    return (len(code) <= 7 and bool(code) and
            (code in referenced or bool(re.search(r"系统|层级|分类|机组|区域|根节点", dtype))
             or bool(re.search(r"系统$|机组$|根节点$", name))))


@lru_cache(maxsize=10000)
def system_evidence(explicit):
    structural = bool(re.fullmatch(r"[A-Z]{3}|[0-9]{2}[A-Z]{3}[0-9]{2}", explicit))
    candidates = ([explicit if len(explicit) == 3 else explicit[2:5]] if structural
                  else pick_system_code(explicit, limit=len(ids.SYS_DICT)))
    exact = structural or any(explicit == str(ids.SYS_DICT.get(code, {}).get("cn", "")) for code in candidates)
    return candidates, exact


def findings(rec, referenced=()):
    """Return stage, message, basis, confirmed tuples in audit priority order."""
    code, name, dtype = rec.get("kks_code", ""), rec.get("name", ""), rec.get("device_type", "")
    kwargs = {"include_measurement": not rec.get("finished_product", False)}
    if not code or classify_row(rec.get("old_code"), name, dtype, **kwargs).is_point or classify_row(code, name, dtype, **kwargs).is_point:
        return []
    if is_node(rec, referenced):
        return []
    if not DEVICE_PATTERN.fullmatch(code):
        return [(0, "完整设备编码必须为十二位：2位数字、3位大写字母、2位数字、2位大写字母、3位数字；字母段不允许I/O。",
                 "项目十二位设备交付要求；GB/T 50549-2020 第4.4节", True)]
    result = []
    dev, system = code[7:9], code[2:5]
    body, typed = subject(name), subject(dtype)
    if body and typed and not body[1].intersection(typed[1]):
        result.append((1, "设备名称主体与设备类型字段冲突，无法确认设备类别。", "设备字段一致性；附录B", False))
    elif body or typed:
        evidence = body or typed
        if dev in ids.DEV_DICT and dev not in evidence[1]:
            result.append((1, f"设备分类码{dev}的分类含义与设备主体“{evidence[0]}”不一致。", "GB/T 50549-2020 第4.4.3节、附录B", True))
    else:
        result.append((1, "现有名称和设备类型不能唯一确认设备主体，设备分类一致性待核。", "GB/T 50549-2020 第4.4.3节、附录B", False))
    explicit = str(rec.get("system") or "").strip()
    if rec.get("finished_product"):
        parent = str(rec.get("parent_code") or "").strip()
        explicit = parent if re.fullmatch(r"[0-9]{2}[A-HJ-NP-Z]{3}[0-9]{2}", parent) else ""
    if explicit:
        candidates, exact = system_evidence(explicit)
        if exact and len(candidates) == 1 and candidates[0] in ids.SYS_DICT:
            if system != candidates[0]:
                result.append((2, "编码系统分类与明确的系统字段不一致。", "GB/T 50549-2020 第4.4.2节、附录A；系统字段=" + explicit, True))
        else:
            result.append((2, "系统字段无法唯一对应分类索引，工艺系统归属待核。", "GB/T 50549-2020 第3.3节、附录A；系统字段=" + explicit, False))
    else:
        result.append((2, "缺少明确工艺系统字段，需结合图纸或工程编码索引确认系统归属。", "GB/T 50549-2020 第3.3节、第4.4.2节", False))
    if not name.strip():
        result.append((3, "设备名称缺失，无法完成名称语义审查。", "设备台账完整性", True))
    else:
        result.append((3, "设备名称中的功能、用途、介质和位置与编码含义的一致性需进行语义复核。",
                       f"名称={name}；设备类型={dtype}；系统分类释义={ids.SYS_DICT.get(system, {}).get('cn', '未收录')}；设备分类释义={ids.DEV_DICT.get(dev, {}).get('cn', '未收录')}", False))
    if code[-3:] == "000":
        result.append((4, "设备编号000不符合一般设备流水编号范围；特殊用途须核对适用依据。", "GB/T 50549-2020 第4.4.5节：一般采用001至999，特殊约定待核", False))
    for message in unit_conflicts(code, name, rec.get("unit", "")):
        result.append((5, message, "GB/T 50549-2020 表4.3.1-2；机组归属需核对关联资料", False))
    return result
