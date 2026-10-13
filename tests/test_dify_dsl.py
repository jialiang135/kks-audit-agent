# -*- coding: utf-8 -*-
"""Dify DSL 的结构与"同源"测试（两个版本都测：原生 LLM 节点版 / HTTP 兜底版）。

为什么不只生成完就完事：DSL 里有三处**必须与引擎保持一致**，否则就是静默失真——
  ① 提示词与 few-shot（改一处忘一处 → Dify 侧与引擎侧生成风格不一致）
  ② **本地复核规则**（Dify 侧放行了引擎侧会退回的名称 → 成果里混进脏名字）
  ③ 原生 LLM 节点引用的**变量名**（引用名写错 → 节点拿到空串，静默生成不出东西）

所以这里把 DSL 里的代码**真的执行起来**，与 `kks_gen/ai_names.py` 的规则逐例比对。
"""
from __future__ import annotations

import json
import pathlib
import re
import sys
import unittest

PROJECT = pathlib.Path(__file__).resolve().parents[1]
for p in (str(PROJECT), str(PROJECT / "tools"), str(PROJECT / "kks_gen")):
    if p not in sys.path:
        sys.path.insert(0, p)

import build_dify_dsl  # noqa: E402
import ai_names  # noqa: E402

VARIANTS = ("llm", "http")
SAMPLE = PROJECT / "宁夏智能体（文件上传知识库）.yml"      # 用户实例里能导入的样例（对齐版本号）


def _nodes(dsl: dict) -> list[dict]:
    return dsl["workflow"]["graph"]["nodes"]


def _code_of(dsl: dict, title: str) -> str:
    return next(n["data"]["code"] for n in _nodes(dsl) if n["data"].get("title") == title)


def _load(code: str, title: str = "node") -> dict:
    ns: dict = {}
    exec(compile(code, f"<dify:{title}>", "exec"), ns)  # noqa: S102 - 就是要跑 DSL 里的代码
    return ns


class StructureTest(unittest.TestCase):
    def test_kind_and_version_align_with_user_instance(self):
        """版本号必须与用户实例里那份能导入的样例一致，否则导入会失败。"""
        for variant in VARIANTS:
            with self.subTest(variant=variant):
                dsl = build_dify_dsl.build_dsl(variant)
                self.assertEqual("app", dsl["kind"])
                if SAMPLE.is_file():
                    m = re.search(r"^version:\s*(\S+)", SAMPLE.read_text(encoding="utf-8"), re.M)
                    if m:
                        self.assertEqual(m.group(1), dsl["version"])

    def test_ids_unique_and_edges_connect_existing_nodes(self):
        for variant in VARIANTS:
            dsl = build_dify_dsl.build_dsl(variant)
            ids = [n["id"] for n in _nodes(dsl)]
            edges = dsl["workflow"]["graph"]["edges"]
            with self.subTest(variant=variant):
                self.assertEqual(len(ids), len(set(ids)), "节点 id 不能重复")
                for e in edges:
                    self.assertIn(e["source"], ids)
                    self.assertIn(e["target"], ids)

    def test_edge_types_match_node_types(self):
        """边的 sourceType/targetType 必须与实际节点类型一致（Dify 靠它渲染连线）。"""
        for variant in VARIANTS:
            dsl = build_dify_dsl.build_dsl(variant)
            type_of = {n["id"]: n["data"]["type"] for n in _nodes(dsl)}
            for e in dsl["workflow"]["graph"]["edges"]:
                with self.subTest(variant=variant, edge=e["id"]):
                    self.assertEqual(type_of[e["source"]], e["data"]["sourceType"])
                    self.assertEqual(type_of[e["target"]], e["data"]["targetType"])

    def test_node_types_are_known_to_dify(self):
        """只用用户样例里出现过的类型 + `llm`（用户指名要的原生大模型节点）。"""
        allowed = {"start", "if-else", "code", "http-request", "answer",
                   "document-extractor", "llm"}
        for variant in VARIANTS:
            for n in _nodes(build_dify_dsl.build_dsl(variant)):
                with self.subTest(variant=variant, node=n["id"]):
                    self.assertIn(n["data"]["type"], allowed)

    def test_only_llm_variant_uses_native_llm_node(self):
        llm_llm = build_dify_dsl.build_dsl("llm")
        http_llm = build_dify_dsl.build_dsl("http")
        self.assertIn("llm", [n["data"]["type"] for n in _nodes(llm_llm)])
        self.assertIn("http-request", [n["data"]["type"] for n in _nodes(http_llm)])
        self.assertNotIn("llm", [n["data"]["type"] for n in _nodes(http_llm)])


class CodeNodeTest(unittest.TestCase):
    def test_all_code_nodes_compile(self):
        for variant in VARIANTS:
            for n in _nodes(build_dify_dsl.build_dsl(variant)):
                if n["data"]["type"] != "code":
                    continue
                with self.subTest(variant=variant, title=n["data"]["title"]):
                    compile(n["data"]["code"], "<dify>", "exec")

    def test_declared_outputs_exist_in_code(self):
        """声明的输出名必须真在代码里返回（Dify 靠声明取名，写错就拿到空值）。"""
        for variant in VARIANTS:
            for n in _nodes(build_dify_dsl.build_dsl(variant)):
                if n["data"]["type"] != "code":
                    continue
                for out in (n["data"].get("outputs") or {}):
                    with self.subTest(variant=variant, node=n["data"]["title"], output=out):
                        self.assertIn(f'"{out}"', n["data"]["code"])

    def test_prompt_and_fewshot_come_from_the_engine(self):
        """提示词与 few-shot 必须与 `ai_names.py` 完全一致（防两处漂移）。"""
        for variant in VARIANTS:
            ns = _load(_code_of(build_dify_dsl.build_dsl(variant), "构造提示"))
            with self.subTest(variant=variant):
                self.assertEqual(ai_names.NAME_PROMPT, ns["NAME_PROMPT"])
                self.assertEqual(len(ai_names.FEWSHOT), len(ns["FEWSHOT"]))
                self.assertEqual(ai_names.FEWSHOT[0][3], ns["FEWSHOT"][0]["正确的设备名称"])

    def test_native_llm_node_references_the_same_prompt(self):
        """🔴 原生 LLM 节点的 prompt_template 必须引用代码节点产出的那两个变量。

        引用名写错 → 节点拿到空串、静默生成不出东西（这类错误导入时不报，跑起来才发现）。
        """
        dsl = build_dify_dsl.build_dsl("llm")
        llm = next(n for n in _nodes(dsl) if n["data"]["type"] == "llm")
        build_id = next(n["id"] for n in _nodes(dsl) if n["data"].get("title") == "构造提示")
        text = " ".join(p["text"] for p in llm["data"]["prompt_template"])
        self.assertIn(f"{{{{#{build_id}.system_prompt#}}}}", text)
        self.assertIn(f"{{{{#{build_id}.user_payload#}}}}", text)
        roles = [p["role"] for p in llm["data"]["prompt_template"]]
        self.assertEqual(["system", "user"], roles)
        self.assertEqual(0, llm["data"]["model"]["completion_params"]["temperature"])

    def test_build_node_llm_variant_outputs_prompt_parts(self):
        ns = _load(_code_of(build_dify_dsl.build_dsl("llm"), "构造提示"))
        rows = [{"row": 1, "kks": "07AAB02UH001", "original_name": "7号机组NCS测控屏",
                 "parent_name": "1号机NCS测控屏系统"}]
        out = ns["main"](json.dumps(rows, ensure_ascii=False))
        self.assertIn("编制工程师", out["system_prompt"])
        payload = json.loads(out["user_payload"])
        self.assertEqual(8, len(payload["改写规律的真实样例"]))
        self.assertEqual("07AAB02UH001", payload["待改写的行"][0]["kks"])

    def test_build_node_http_variant_produces_valid_body(self):
        ns = _load(_code_of(build_dify_dsl.build_dsl("http"), "构造提示"))
        rows = [{"row": 1, "kks": "07AAB02UH001", "original_name": "7号机组NCS测控屏"}]
        out = ns["main"](json.dumps(rows, ensure_ascii=False), "deepseek-flash")
        body = json.loads(out["request_json"])
        self.assertEqual("deepseek-flash", body["model"])
        self.assertEqual(0, body["temperature"])
        self.assertEqual(["system", "user"], [m["role"] for m in body["messages"]])

    def test_build_node_rejects_oversized_batch(self):
        for variant in VARIANTS:
            ns = _load(_code_of(build_dify_dsl.build_dsl(variant), "构造提示"))
            rows = [{"row": i, "kks": "07AAB02UH001"} for i in range(41)]
            with self.subTest(variant=variant), self.assertRaises(ValueError):
                ns["main"](json.dumps(rows))

    def test_check_node_parses_fenced_json(self):
        """模型把 JSON 包在 ```json 围栏里也要能解析出来（两种接法都要）。"""
        rows = [{"row": 7, "kks": "07BBA00UC004", "original_name": "烟气再循环风机7A变频器"}]
        content = "```json\n" + json.dumps(
            {"items": [{"row": 7, "name": "7A烟气再循环风机变频器", "reason": "分支号提前",
                        "confidence": 0.9}]}, ensure_ascii=False) + "\n```"
        want_rows = json.dumps(rows, ensure_ascii=False)

        ns = _load(_code_of(build_dify_dsl.build_dsl("llm"), "解析与复核"))
        got_llm = json.loads(ns["main"](content, want_rows)["items_json"])
        self.assertEqual(1, got_llm["batch"]["named"])
        self.assertEqual("7A烟气再循环风机变频器", got_llm["items"][0]["name"])

        ns = _load(_code_of(build_dify_dsl.build_dsl("http"), "解析与复核"))
        body = {"choices": [{"message": {"content": content}}]}
        got_http = json.loads(ns["main"](200, body, want_rows)["items_json"])
        self.assertEqual("7A烟气再循环风机变频器", got_http["items"][0]["name"])


class ValidatorParityTest(unittest.TestCase):
    """🔴 DSL 里的复核规则必须与引擎 `ai_names._validate_name` **逐例同判**（两个版本都要）。

    这条最要紧：两边若判得不一样，Dify 放行的脏名称会绕过引擎的退回规则直接写进成果。
    """

    CASES = [
        ("正常", "7A烟气再循环风机变频器", {"kks": "07BBA00UC004", "original_name": "烟气再循环风机7A变频器"}),
        ("空名", "", {"kks": "07BBA00UC004", "original_name": "烟气再循环风机7A变频器"}),
        ("过短", "屏", {"kks": "07BBA00UC004", "original_name": "屏"}),
        ("过长", "泵" * 61, {"kks": "07BBA00UC004", "original_name": "泵" * 61}),
        ("机组号写错", "1号机NCS测控屏", {"kks": "07AAB02UH001", "original_name": "7号机组NCS测控屏"}),
        ("跑题", "完全换了个东西", {"kks": "07AAB02UH001", "original_name": "7号机组NCS测控屏"}),
        ("明显变短", "屏", {"kks": "07AAB02UH001", "original_name": "7号机组NCS测控屏"}),
        ("补了父级前缀", "10kV工作7A段 母线PT", {"kks": "07BBA01UC003", "original_name": "母线PT"}),
    ]

    def test_same_verdict_as_engine(self):
        for variant in VARIANTS:
            ns = _load(_code_of(build_dify_dsl.build_dsl(variant), "解析与复核"), "解析与复核")
            for label, name, row in self.CASES:
                with self.subTest(variant=variant, case=label):
                    engine_ok, _ = ai_names._validate_name({"name": name}, row)
                    dsl_ok, _ = ns["_validate"]({"name": name}, row)
                    self.assertEqual(engine_ok, dsl_ok,
                                     f"两侧判定不一致：引擎={engine_ok} DSL={dsl_ok}")


class RoundTripTest(unittest.TestCase):
    """生成出来的 YAML 必须是合法 YAML（装不了 PyYAML 就跳过这条）。"""

    def test_yaml_parses_and_keeps_key_fields(self):
        try:
            import yaml
        except ImportError:                                     # pragma: no cover
            self.skipTest("环境里没有 PyYAML（生成器本身不依赖它，故仅此项跳过）")
        for variant in VARIANTS:
            with self.subTest(variant=variant):
                data = yaml.safe_load(build_dify_dsl.to_yaml(build_dify_dsl.build_dsl(variant)))
                self.assertEqual("app", data["kind"])
                self.assertEqual("advanced-chat", data["app"]["mode"])
                self.assertEqual(5, len(data["workflow"]["graph"]["nodes"]))
                self.assertEqual(4, len(data["workflow"]["graph"]["edges"]))

    def test_files_on_disk_match_the_generator(self):
        """磁盘上的两个 yml 必须是当前生成器的产物（改了生成器要重跑，别让文件过期）。"""
        for variant, path in build_dify_dsl.OUTPUTS.items():
            if not path.is_file():
                continue
            with self.subTest(variant=variant):
                self.assertEqual(build_dify_dsl.to_yaml(build_dify_dsl.build_dsl(variant)),
                                 path.read_text(encoding="utf-8"))


if __name__ == "__main__":
    unittest.main()
