"""Project delivery policy: signals/measurement points have no device KKS.

This is a device-table policy, not a claim that GB/T forbids signal identifiers.
Physical instruments remain devices; measured loops/values are point rows.
"""
import re
import unicodedata
from dataclasses import dataclass
from .identifiers import RE_FUNCTION, RE_SIGNAL


@dataclass(frozen=True)
class RowRole:
    is_point: bool = False
    body: str = ""
    reason: str = ""


def classify_row(code="", name="", device_type="", *, include_measurement=True):
    value = re.sub(r"\s+", "", unicodedata.normalize("NFKC", str(code or ""))).upper()
    label = str(name or "")
    dtype = str(device_type or "")
    normalized_label = re.sub(r"\s+", "", unicodedata.normalize("NFKC", label)).rstrip("。.;；")
    # Full and glued signal identifiers, including optional H1/H4.
    for split in range(12, 15):
        body, tail = value[:split], value[split:]
        if RE_FUNCTION.fullmatch(body) and tail and RE_SIGNAL.fullmatch(";" + tail.lstrip(";")):
            return RowRole(True, body, "编码包含信号源/信号应用标识")
    if RE_SIGNAL.fullmatch(value):
        return RowRole(True, "", "独立信号标识")
    signal_suffix = re.search(
        r"(?:指令|命令|请求|测点|信号|反馈|开到位|关到位|已运行|已停止|已启动|已关闭|"
        r"已开启|已打开|已合闸|已分闸|运行状态|停止状态|关闭状态|开启状态|联锁位|"
        r"模拟量|开关量|报警|启动|停止|开启|关闭|打开|复位|跳闸|合闸|分闸|"
        r"状态|故障|正常|异常|运行|停运|闭锁|允许|禁止|联锁)"
        r"[0-9一二三四五六七八九十号#]*\)?$", normalized_label)
    if signal_suffix:
        match = RE_FUNCTION.fullmatch(value)
        return RowRole(True, value if match and not match["dev"].startswith("B") else "", "名称末端表示控制动作、测量信号或状态反馈")
    action_label = re.sub(r"^[0-9一二三四五六七八九十]+号(?:机组|机|炉)", "", normalized_label)
    if re.match(r"^(?:启动|停止|开启|关闭|复位|联锁跳闸|开度给定|频率给定)", action_label):
        return RowRole(True, value if RE_FUNCTION.fullmatch(value) else "", "名称明确为控制指令")
    if re.fullmatch(r"(?:AI|AO|DI|DO)(?:点)?", dtype.strip(), re.I):
        return RowRole(True, "", "类型明确为指令/测点")
    # Measurement words may qualify a physical component. The final noun
    # identifies the component; a feedback/state suffix identifies a signal.
    physical_body = re.search(
        r"(?:一次门|二次门|取样阀|仪表阀|截止阀|针型阀|阀门|电动门|气动门|"
        r"阀|泵|风机|变送器|传感器|压力表|流量计|液位计|热电偶|热电阻|"
        r"压力开关|差压开关|液位开关|限位开关|流量开关|温度开关|选择开关|"
        r"继电器|执行机构|执行器|定位器|接线盒|端子箱|控制柜|电磁阀|探头|"
        r"测量装置|监测仪|测量仪|测速仪|转速表)"
        r"[A-Za-z0-9一二三四五六七八九十号#\s()（）\-]*$", normalized_label)
    if physical_body:
        return RowRole(False, value, "名称末端主体为实体设备，前面的测量/信号词为用途限定")
    explicit = re.search(r"指令|测点|信号|反馈|模拟量|开关量|报警", dtype)
    if explicit or re.fullmatch(r"(?:AI|AO|DI|DO)(?:点)?", dtype.strip(), re.I):
        match = RE_FUNCTION.fullmatch(value)
        body = value if match and not match["dev"].startswith("B") else ""
        return RowRole(True, body, "名称或类型明确为指令/测点")
    match = RE_FUNCTION.fullmatch(value)
    physical = re.search(
        r"传感器|变送器|探头|温度计|压力表|流量计|液位计|仪表|热电偶|热电阻|"
        r"压力开关|差压开关|液位开关|限位开关|流量开关|温度开关|流量孔板|"
        r"监测仪|检漏仪|测量装置|测量仪|测速仪|转速表", label + " " + dtype)
    # A quantity row may use an incorrect non-B classification in the source.
    # Generic measurement equipment alone is not enough: require a measured
    # quantity name and exclude explicit physical instruments.
    quantity = re.search(r"(?:流量|压力|差压|温度|液位|水位|浓度|露点|转速|振动|频率|开度)(?:\d+)?$", label)
    if include_measurement and quantity and not physical and ("测量设备" in dtype or "测点" in dtype):
        return RowRole(True, "", "名称为测量量且类型为测量设备/测点；未声明实体仪表")
    if include_measurement and match and match["dev"].startswith("B") and not physical:
        return RowRole(True, "", "B类测量回路/测量点；未声明为实体仪表")
    return RowRole()


def is_table_instruction(code="", name="", source_code=""):
    """Recognize the template's explicit column instructions, not equipment."""
    return (str(code or "").startswith("按照") and str(name or "").startswith("按照")
            and "提供" in str(source_code or "")
            and ("KKS" in str(source_code or "").upper() or "编码" in str(source_code or "")))
