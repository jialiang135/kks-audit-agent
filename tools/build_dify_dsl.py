# -*- coding: utf-8 -*-
"""把「AI 补设备名称」这一步生成为 **Dify 可导入的 DSL**（一次生成两个版本）。

| 输出文件 | 接法 | 何时用 |
|---|---|---|
| `dify/kks_设备名称生成.yml` | **Dify 原生 LLM 节点**（`type: llm`） | 推荐：模型在 Dify 里选，不用把 API Key 传进来 |
| `dify/kks_设备名称生成_HTTP兜底版.yml` | HTTP 节点直连 `chat/completions` | 保底：只用你样例里出现过的节点类型，导入最稳 |

## 为什么用脚本生成，而不是手写 yml

提示词（`NAME_PROMPT`）、8 条 few-shot、以及**本地复核规则**都来自 `kks_gen/ai_names.py`。
手写一份 yml 等于把它们复制一遍，**改一处忘一处**。这里直接 import 进来，
`tests/test_dify_dsl.py` 再断言"两边的提示词一致、复核规则逐例同判"，从根上防漂移。

## 生成的 Dify 应用做什么 / 不做什么

✅ **做**：一批（≤40 行）的**名称生成**——与我们引擎里 `ai_names.resolve_names` 的一次请求
   **等价**：同一份系统提示词、同一批 few-shot、同样的本地复核（退回规则）。
   引擎按 40 行一批调用它，Dify 侧保持"无状态的一次调用"，不负责分批/并发/进度。

❌ **不做**（Dify 做不了，理由见交接文档 §6 的评估）：
   · Excel 读写与版式保真（代码节点是**受限沙箱**，没有 openpyxl，写不出带样式/合并区/列宽的 xlsx）
   · 确定性规则引擎（63 条正则 + 3,390 条系统码 + 222 条设备码 + 建树/唯一性闸门）
   · 全表 58,275 行 → 48,134 节点的编排（要分 1,457 批，且模型侧实测不可复现）

## 用法

```bash
python tools/build_dify_dsl.py          # 生成上面两个文件
```
"""
from __future__ import annotations

import argparse
import json
import pathlib
import sys

ROOT = pathlib.Path(__file__).resolve().parents[1]
for _p in (str(ROOT), str(ROOT / "kks_gen")):
    if _p not in sys.path:
        sys.path.insert(0, _p)


# ------------------------------------------------------------------ YAML 输出
# 不依赖 PyYAML（.venv 里没有）：自己吐一份最小 YAML。
# 字符串一律走 json.dumps → **JSON 字符串是合法的 YAML 双引号标量**，转义天然安全；
# 多行的代码/提示词用「块标量」（`|-`）输出，人看着也清楚。


class Block(str):
    """标记：按 YAML 块标量（`|-`）输出，保留换行、去掉结尾空行。"""


def _scalar(value) -> str:
    if value is None:
        return "null"
    if isinstance(value, bool):
        return "true" if value else "false"
    if isinstance(value, (int, float)):
        return str(value)
    return json.dumps(str(value), ensure_ascii=False)


def _emit(obj, indent: int, out: list[str]) -> None:
    pad = " " * indent

    def _entry(key: str, value, prefix: str) -> None:
        if isinstance(value, Block):
            out.append(f"{prefix}{key}: |-")
            for line in str(value).rstrip("\n").split("\n"):
                out.append((" " * (indent + 4) + line) if line else "")
        elif isinstance(value, (dict, list)) and value:
            out.append(f"{prefix}{key}:")
            _emit(value, indent + 4, out)
        elif isinstance(value, (dict, list)):
            out.append(f"{prefix}{key}: {('{}' if isinstance(value, dict) else '[]')}")
        else:
            out.append(f"{prefix}{key}: {_scalar(value)}")

    if isinstance(obj, dict):
        for key, value in obj.items():
            if isinstance(value, Block):
                out.append(f"{pad}{key}: |-")
                for line in str(value).rstrip("\n").split("\n"):
                    out.append((" " * (indent + 2) + line) if line else "")
            elif isinstance(value, (dict, list)) and value:
                out.append(f"{pad}{key}:")
                _emit(value, indent + 2, out)
            elif isinstance(value, (dict, list)):
                out.append(f"{pad}{key}: {('{}' if isinstance(value, dict) else '[]')}")
            else:
                out.append(f"{pad}{key}: {_scalar(value)}")
    elif isinstance(obj, list):
        for item in obj:
            if isinstance(item, dict):
                first = True
                for key, value in item.items():
                    prefix = f"{pad}- " if first else f"{pad}  "
                    first = False
                    if isinstance(value, Block):
                        out.append(f"{prefix}{key}: |-")
                        for line in str(value).rstrip("\n").split("\n"):
                            out.append((" " * (indent + 4) + line) if line else "")
                    elif isinstance(value, (dict, list)) and value:
                        out.append(f"{prefix}{key}:")
                        _emit(value, indent + 4, out)
                    elif isinstance(value, (dict, list)):
                        out.append(f"{prefix}{key}: {('{}' if isinstance(value, dict) else '[]')}")
                    else:
                        out.append(f"{prefix}{key}: {_scalar(value)}")
            elif isinstance(item, (dict, list)):
                out.append(f"{pad}-")
                _emit(item, indent + 2, out)
            else:
                out.append(f"{pad}- {_scalar(item)}")


def to_yaml(obj) -> str:
    out: list[str] = []
    _emit(obj, 0, out)
    return "\n".join(out) + "\n"


# ------------------------------------------------------------------ 代码节点模板
# 用三引号**原样**装代码（占位符 @@X@@ 再替换）——之前用字符串拼接，转义把文件写坏过两次。
_BUILD_TMPL = '''"""把待命名行拼成提示（与 kks-audit-agent 的 ai_names 同源）。"""
import json

# ---- 系统提示词：与 kks_gen/ai_names.py 的 NAME_PROMPT 完全一致（有测试盯着）----
NAME_PROMPT = @@NAME_PROMPT@@

# ---- few-shot：全部来自中煤鄂州四期竣工版真实成果 ----
FEWSHOT = json.loads(@@FEWSHOT@@)

MAX_BATCH = 40          # 与引擎的 AI_NAME_BATCH_SIZE 对齐；超过请先分批

ROW_FIELDS = ("row", "kks", "original_name", "parent_name", "parent_kks", "profession",
              "drawing", "system_meaning", "device_meaning", "phase_suffix",
              "same_class_names", "sibling_names")


def _rows(rows_json):
    rows = json.loads(rows_json or "[]")
    if isinstance(rows, dict):
        rows = rows.get("rows") or []
    if not isinstance(rows, list) or not rows:
        raise ValueError("rows_json 必须是非空 JSON 数组")
    if len(rows) > MAX_BATCH:
        raise ValueError("本工作流一次最多 {} 行，收到 {} 行；请由调用方分批".format(MAX_BATCH, len(rows)))
    return rows


def _payload(rows):
    return {
        "task": "generate_device_name",
        "改写规律的真实样例": FEWSHOT,
        "待改写的行": [{k: r.get(k, "") for k in ROW_FIELDS} for r in rows],
    }


@@BUILD_BODY@@
'''

_BUILD_HTTP_BODY = '''def main(rows_json: str, model: str = "@@MODEL@@") -> dict:
    rows = _rows(rows_json)
    body = {
        # 模型名由调用方传（可在不改编排的情况下换模型做 A/B）；不给就用默认
        "model": (model or "@@MODEL@@").strip(),
        "temperature": 0,
        "messages": [
            {"role": "system", "content": NAME_PROMPT},
            {"role": "user", "content": json.dumps(_payload(rows), ensure_ascii=False)},
        ],
    }
    return {"request_json": json.dumps(body, ensure_ascii=False), "row_count": len(rows)}
'''

_BUILD_LLM_BODY = '''def main(rows_json: str) -> dict:
    rows = _rows(rows_json)
    return {"system_prompt": NAME_PROMPT,
            "user_payload": json.dumps(_payload(rows), ensure_ascii=False),
            "row_count": len(rows)}
'''

_VALIDATE_HEAD = '''"""解析模型返回 + 本地复核，只把"过审"的名称交出去。"""
import json
import re

ILLEGAL = re.compile(r"[\\x00-\\x1f\\x7f]")


def _unit_of(kks: str) -> str:
    """全厂码 → 机组号（表4.3.1-2）：01~60 = 第 N 号机；61~90 两机共用；00/91~99 非单机。"""
    head = str(kks or "")[:2]
    if not head.isdigit():
        return ""
    n = int(head)
    return str(n) if 1 <= n <= 60 else ""


def _extract_json(text: str) -> dict:
    """从模型回复里抠出 JSON 对象（容忍 ```json 围栏与前后废话）。"""
    t = (text or "").strip()
    t = re.sub(r"^```(?:json)?", "", t).strip()
    t = re.sub(r"```$", "", t).strip()
    try:
        return json.loads(t)
    except Exception:
        pass
    start, end = t.find("{"), t.rfind("}")
    if start >= 0 and end > start:
        try:
            return json.loads(t[start:end + 1])
        except Exception:
            return {}
    return {}


def _validate(item: dict, row: dict):
    name = str(item.get("name", "") or "").strip()
    if not name:
        return False, ""
    if ILLEGAL.search(name):
        return False, "含控制字符"
    if len(name) < 3:
        return False, "名称过短（{} 字）".format(len(name))
    if len(name) > 60:
        return False, "名称过长（{} 字）".format(len(name))
    want = _unit_of(row.get("kks", ""))
    if want:
        for bad in ("1号机", "2号机", "3号机", "4号机", "5号机", "6号机",
                    "7号机", "8号机", "9号机"):
            if bad in name and bad != want + "号机" and (want + "号机") not in name:
                return False, "机组号与编码前缀不符（出现「{}」，应为「{}号机」）".format(bad, want)
    old = str(row.get("original_name", "") or "").strip()
    if old and len(old) >= 4:
        words = re.findall(r"[\\u4e00-\\u9fff]{2,}", old)
        if words and not any(w in name for w in words):
            return False, "与原名称语义偏离过大（原名称中的中文词一个都没保留）"
        if len(name) < len(old) * 0.6:
            return False, "名称比原名称明显变短（原 {} 字 → 生成 {} 字）".format(len(old), len(name))
    return True, ""
'''

_VALIDATE_LLM_MAIN = '''

def main(text, rows_json) -> dict:
    rows = json.loads(rows_json or "[]")
    if isinstance(rows, dict):
        rows = rows.get("rows") or []
    by_row = {str(r.get("row", "")): r for r in rows}
    raw = text if isinstance(text, str) else str(text or "")
    payload = _extract_json(raw)
    code = 200
    where = "原生 LLM 节点"
'''

_VALIDATE_HTTP_MAIN = '''

def main(status_code, body, rows_json) -> dict:
    rows = json.loads(rows_json or "[]")
    if isinstance(rows, dict):
        rows = rows.get("rows") or []
    by_row = {str(r.get("row", "")): r for r in rows}
    try:
        code = int(status_code)
    except Exception:
        code = 0
    data = body if isinstance(body, dict) else _extract_json(str(body))
    raw = ""
    try:
        raw = data["choices"][0]["message"]["content"]
    except Exception:
        raw = ""
    payload = _extract_json(raw) if raw else {}
    where = "HTTP"
'''

_VALIDATE_TAIL = '''    items, rejected = [], {}
    for item in (payload.get("items") or []):
        if not isinstance(item, dict):
            continue
        row_no = str(item.get("row", "")).strip()
        row = by_row.get(row_no)
        if row is None:
            continue
        ok, why = _validate(item, row)
        if not ok:
            if why:
                rejected[why] = rejected.get(why, 0) + 1
            continue
        try:
            conf = float(item.get("confidence", 0) or 0)
        except Exception:
            conf = 0.0
        items.append({"row": row_no, "name": str(item["name"]).strip(),
                      "reason": str(item.get("reason", ""))[:120],
                      "confidence": conf, "kks": row.get("kks", "")})
    result = {
        "items": items,
        "batch": {"rows": len(rows), "named": len(items),
                  "rejected": sum(rejected.values()), "reject_reasons": rejected,
                  "http_status": code, "via": where,
                  "returned": len(payload.get("items") or []),
                  "model_text_empty": (str(raw).strip() == "")},
    }
    return {"items_json": json.dumps(result, ensure_ascii=False),
            "named_count": len(items),
            "rejected_count": sum(rejected.values()),
            "summary": "本批 {} 行 → 采纳 {} 条、退回 {} 条（{} {}）".format(
                len(rows), len(items), sum(rejected.values()), where, code)}
'''


def _render(tmpl: str, **repl) -> str:
    for key, value in repl.items():
        tmpl = tmpl.replace(f"@@{key}@@", value)
    left = [m for m in ("@@NAME_PROMPT@@", "@@FEWSHOT@@", "@@BUILD_BODY@@", "@@MODEL@@")
            if m in tmpl]
    assert not left, f"模板占位符未替换：{left}"
    return tmpl


def _code_build(variant: str, model: str) -> str:
    import ai_names

    fewshot = [{"原名称": o, "上级设备名称": p, "KKS编码": c, "正确的设备名称": n}
               for o, p, c, n in ai_names.FEWSHOT]
    return _render(
        _BUILD_TMPL,
        NAME_PROMPT=json.dumps(ai_names.NAME_PROMPT, ensure_ascii=False),
        FEWSHOT=json.dumps(json.dumps(fewshot, ensure_ascii=False), ensure_ascii=False),
        BUILD_BODY=(_BUILD_LLM_BODY if variant == "llm" else _BUILD_HTTP_BODY),
        MODEL=model,
    )


def _code_validate(variant: str) -> str:
    body = _VALIDATE_LLM_MAIN if variant == "llm" else _VALIDATE_HTTP_MAIN
    return _VALIDATE_HEAD + body + _VALIDATE_TAIL


# ------------------------------------------------------------------ DSL 组装
N_START, N_BUILD, N_MID, N_CHECK, N_ANSWER = (
    "1758000000001", "1758000000002", "1758000000003", "1758000000004", "1758000000005")
DEFAULT_MODEL = "deepseek-flash"
CHAT_COMPLETIONS_URL = "https://www.ai.atyou.cn/v1/chat/completions"


def _node(node_id: str, x: float, y: float, data: dict, *, height: int = 100) -> dict:
    return {"data": data, "height": height, "id": node_id,
            "position": {"x": x, "y": y}, "positionAbsolute": {"x": x, "y": y},
            "selected": False, "sourcePosition": "right", "targetPosition": "left",
            "type": "custom", "width": 241, "zIndex": 0}


def _edge(source: str, source_type: str, target: str, target_type: str) -> dict:
    return {"data": {"isInIteration": False, "isInLoop": False,
                     "sourceType": source_type, "targetType": target_type},
            "id": f"{source}-source-{target}-target", "selected": False,
            "source": source, "sourceHandle": "source", "target": target,
            "targetHandle": "target", "type": "custom", "zIndex": 0}


def build_dsl(variant: str = "llm") -> dict:
    """`variant`: `llm` = 原生大模型节点（推荐）；`http` = HTTP 节点直连（保底）。"""
    if variant not in ("llm", "http"):
        raise ValueError("variant 只能是 llm 或 http")
    is_llm = variant == "llm"

    start_vars = [
        {"label": "待命名行（JSON 数组，一次最多 40 行）", "max_length": 200000,
         "options": [], "required": True, "type": "paragraph", "variable": "rows_json"},
    ]
    if not is_llm:                      # 原生 LLM 节点的接法里，模型与密钥都在 Dify 里配，不用传
        start_vars += [
            {"label": "模型名（如 deepseek-flash）", "max_length": 128,
             "options": [], "required": False, "type": "text-input", "variable": "model"},
            {"label": "模型 API Key（Bearer）", "max_length": 512,
             "options": [], "required": True, "type": "text-input", "variable": "api_key"},
        ]
    start = _node(N_START, 68, 320, {
        "desc": "输入：一批待命名的行（≤40 行）。" + (
            "" if is_llm else "HTTP 版另需 model 与 api_key（由调用方传入，不落在 DSL 里）"),
        "selected": False, "title": "开始", "type": "start", "variables": start_vars,
    }, height=160 if is_llm else 200)

    build = _node(N_BUILD, 400, 320, {
        "code": Block(_code_build(variant, DEFAULT_MODEL)),
        "code_language": "python3",
        "desc": "拼提示：系统提示词与 few-shot 与 kks_gen/ai_names.py 同源（有测试盯着）",
        "outputs": ({"system_prompt": {"children": None, "type": "string"},
                     "user_payload": {"children": None, "type": "string"},
                     "row_count": {"children": None, "type": "number"}} if is_llm else
                    {"request_json": {"children": None, "type": "string"},
                     "row_count": {"children": None, "type": "number"}}),
        "selected": False, "title": "构造提示", "type": "code",
        "variables": ([{"value_selector": [N_START, "rows_json"], "variable": "rows_json"}]
                      if is_llm else
                      [{"value_selector": [N_START, "rows_json"], "variable": "rows_json"},
                       {"value_selector": [N_START, "model"], "variable": "model"}]),
    }, height=120)

    if is_llm:
        mid = _node(N_MID, 736, 320, {
            "context": {"enabled": False, "variable_selector": []},
            "desc": ("用与引擎相同的提示词生成设备名称。⚠️ 导入后**必须**在本节点里选一个你实例里"
                     "已配置的模型（下面填的 provider/name 只是占位）；温度建议 0。"),
            "model": {"completion_params": {"temperature": 0}, "mode": "chat",
                      "name": DEFAULT_MODEL, "provider": "openai_api_compatible"},
            "prompt_template": [
                {"id": "kks-name-system", "role": "system",
                 "text": f"{{{{#{N_BUILD}.system_prompt#}}}}"},
                {"id": "kks-name-user", "role": "user",
                 "text": f"{{{{#{N_BUILD}.user_payload#}}}}"},
            ],
            "selected": False, "title": "生成名称（原生大模型）", "type": "llm", "variables": [],
            "vision": {"enabled": False},
        }, height=160)
    else:
        mid = _node(N_MID, 736, 320, {
            "authorization": {"config": None, "type": "no-auth"},
            "body": {"data": [{"type": "text", "value": f"{{{{#{N_BUILD}.request_json#}}}}"}],
                     "type": "raw-text"},
            "desc": "调用 OpenAI 兼容的 chat/completions（URL 可在界面里改）",
            "headers": ("Content-Type:application/json\n"
                        f"Authorization:Bearer {{{{#{N_START}.api_key#}}}}"),
            "method": "post", "params": "",
            "retry_config": {"max_retries": 3, "retry_enabled": True, "retry_interval": 100},
            "selected": False, "ssl_verify": True,
            "timeout": {"max_connect_timeout": 0, "max_read_timeout": 180, "max_write_timeout": 0},
            "title": "调用模型（HTTP）", "type": "http-request", "url": CHAT_COMPLETIONS_URL,
        }, height=220)

    check = _node(N_CHECK, 1072, 320, {
        "code": Block(_code_validate(variant)),
        "code_language": "python3",
        "desc": "解析模型返回 + 本地复核（控制字符/长度/机组号一致/语义保留），只交出过审的名称",
        "outputs": {"items_json": {"children": None, "type": "string"},
                    "named_count": {"children": None, "type": "number"},
                    "rejected_count": {"children": None, "type": "number"},
                    "summary": {"children": None, "type": "string"}},
        "selected": False, "title": "解析与复核", "type": "code",
        "variables": ([{"value_selector": [N_MID, "text"], "variable": "text"},
                       {"value_selector": [N_START, "rows_json"], "variable": "rows_json"}]
                      if is_llm else
                      [{"value_selector": [N_MID, "status_code"], "variable": "status_code"},
                       {"value_selector": [N_MID, "body"], "variable": "body"},
                       {"value_selector": [N_START, "rows_json"], "variable": "rows_json"}]),
    }, height=140)

    answer = _node(N_ANSWER, 1408, 320, {
        "answer": Block(f"{{{{#{N_CHECK}.items_json#}}}}"),
        "desc": "输出 items_json（调用方按 JSON 解析）",
        "selected": False, "title": "返回结果", "type": "answer", "variables": [],
    }, height=110)

    via = "Dify 原生大模型节点" if is_llm else "HTTP 节点直连 chat/completions"
    return {
        "app": {
            "description": (f"KKS 设备名称生成（一步，{via}）。输入一批待命名的行（JSON，≤40 行），"
                            "用与 kks-audit-agent 引擎完全一致的提示词与复核规则生成规范设备名称。"
                            "只做「命名」这一步；Excel 读写、分批、进度由调用方负责。"),
            "icon": "🏷️", "icon_background": "#E8F0FE", "icon_type": "emoji",
            "mode": "advanced-chat",
            "name": "KKS 设备名称生成" if is_llm else "KKS 设备名称生成（HTTP 兜底版）",
            "use_icon_as_answer_icon": False,
        },
        "dependencies": [],
        "kind": "app",
        "version": "0.7.0",
        "workflow": {
            "conversation_variables": [],
            "environment_variables": [],
            "features": {
                "file_upload": {"allowed_file_extensions": [], "allowed_file_types": [],
                                "enabled": False,
                                "fileUploadConfig": {"file_size_limit": 15,
                                                     "workflow_file_upload_limit": 10},
                                "image": {"enabled": False, "number_limits": 3,
                                          "transfer_methods": ["local_file", "remote_url"]},
                                "number_limits": 3},
                "opening_statement": ("一次一批（≤40 行）的 KKS 设备名称生成。"
                                      "把待命名的行以 JSON 数组贴进对话框即可；API 调用请传 rows_json。"),
                "retriever_resource": {"enabled": False},
                "sensitive_word_avoidance": {"enabled": False},
                "speech_to_text": {"enabled": False},
                "suggested_questions": [],
                "suggested_questions_after_answer": {"enabled": False},
                "text_to_speech": {"enabled": False, "language": "", "voice": ""},
            },
            "graph": {
                "edges": [
                    _edge(N_START, "start", N_BUILD, "code"),
                    _edge(N_BUILD, "code", N_MID, "llm" if is_llm else "http-request"),
                    _edge(N_MID, "llm" if is_llm else "http-request", N_CHECK, "code"),
                    _edge(N_CHECK, "code", N_ANSWER, "answer"),
                ],
                "nodes": [start, build, mid, check, answer],
                "viewport": {"x": 0, "y": 0, "zoom": 0.8},
            },
            "rag_pipeline_variables": [],
        },
    }


OUTPUTS = {
    "llm": ROOT / "dify" / "kks_设备名称生成.yml",
    "http": ROOT / "dify" / "kks_设备名称生成_HTTP兜底版.yml",
}


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--out-dir", default=str(ROOT / "dify"))
    args = ap.parse_args()
    for variant, default_path in OUTPUTS.items():
        out = pathlib.Path(args.out_dir) / default_path.name
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text(to_yaml(build_dsl(variant)), encoding="utf-8", newline="\n")
        print(f"已生成 {out.relative_to(ROOT)}（{out.stat().st_size/1024:.1f} KB）"
              f"｜{'原生 LLM 节点' if variant == 'llm' else 'HTTP 兜底'}")
    print("导入：Dify → 工作室 → 导入 DSL 文件（原生版导入后先到 LLM 节点里选模型）")
    return 0


if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    raise SystemExit(main())
