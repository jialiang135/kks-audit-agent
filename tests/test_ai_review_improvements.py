import unittest
from unittest.mock import patch

import ai_review
from app_runtime import load_skill_context


class FakeClient:
    calls = []

    def __init__(self, config):
        self.config = config

    def choose_model(self):
        return "fake-model", "test"

    def review(self, model, prompt, payload):
        self.calls.append(payload)
        results = []
        for candidate in payload["candidates"]:
            if payload["review_type"] == "false_positive_verification":
                results.append({
                    "issue_index": candidate["issue_index"],
                    "decision": "needs_human",
                    "priority": candidate.get("priority", "P2"),
                    "confidence": 0.72,
                    "evidence": "第 3 行与第 4 行同父级，规则证据仍成立",
                    "reason": "初审证据不足以排除问题",
                    "suggestion": "人工确认同级编码和设备语义",
                })
            else:
                decision = "likely_false_positive" if candidate["issue_index"] == 0 else "confirmed_issue"
                results.append({
                    "issue_index": candidate["issue_index"],
                    "decision": decision,
                    "priority": candidate.get("priority", "P2"),
                    "confidence": 0.88,
                    "evidence": "当前行、父级行和同级行的编码关系",
                    "reason": "模拟模型已结合证据包判断",
                    "suggestion": "保留问题并人工确认",
                })
        return {"reviews": results}


class AiReviewImprovementTests(unittest.TestCase):
    def test_relevant_skill_context_is_selected(self):
        full = load_skill_context()
        selected = load_skill_context(
            rule_ids={"KKS-31"},
            categories={"父子名称语义矛盾"},
            keywords={"父子", "语义"},
            max_chars=12000,
        )
        self.assertIn("父子", selected)
        self.assertLess(len(selected), len(full))

    def test_evidence_packet_contains_relationship_context(self):
        result = {
            "issues": [],
            "_ai_context": {
                "records": [
                    {"excel_row": 2, "kks_code": "10ABC01", "parent_code": "", "name": "父设备"},
                    {"excel_row": 3, "kks_code": "10ABC01002", "parent_code": "10ABC01", "name": "当前设备"},
                    {"excel_row": 4, "kks_code": "10ABC01003", "parent_code": "10ABC01002", "name": "子设备"},
                    {"excel_row": 5, "kks_code": "10ABC01004", "parent_code": "10ABC01", "name": "同级设备"},
                ]
            },
        }
        issue = {
            "excel_row": 3,
            "kks_code": "10ABC01002",
            "parent_code": "10ABC01",
            "rule_id": "KKS-31",
            "category": "父子名称语义矛盾",
            "priority": "P1",
            "message": "父子名称需要核对",
            "suggestion": "人工确认",
        }
        packet = ai_review._build_evidence_packet(0, issue, result)
        self.assertEqual(packet["current_row"]["excel_row"], 3)
        self.assertEqual(packet["parent_rows"][0]["excel_row"], 2)
        self.assertEqual(packet["child_rows"][0]["excel_row"], 4)
        self.assertEqual(packet["same_level_rows"][0]["excel_row"], 5)
        self.assertTrue(packet["adjacent_rows"])

    def test_group_review_and_false_positive_verification(self):
        FakeClient.calls = []
        result = {
            "issues": [
                {"status": "needs_review", "rule_id": "KKS-31", "category": "父子名称语义矛盾", "priority": "P1", "excel_row": 3, "kks_code": "10ABC01002", "parent_code": "10ABC01", "message": "问题一", "suggestion": "确认"},
                {"status": "needs_review", "rule_id": "KKS-24", "category": "命名歧义", "priority": "P1", "excel_row": 5, "kks_code": "10ABC01004", "parent_code": "10ABC01", "message": "问题二", "suggestion": "确认"},
            ],
            "_ai_context": {"records": [{"excel_row": 3, "kks_code": "10ABC01002", "parent_code": "10ABC01", "name": "当前设备"}, {"excel_row": 5, "kks_code": "10ABC01004", "parent_code": "10ABC01", "name": "同级设备"}]},
        }
        config = ai_review.AIConfig("https://example.test/v1", "test-key", "fake-model", 10, True)
        with patch.object(ai_review, "load_config", return_value=config), patch.object(ai_review, "OpenAICompatibleClient", FakeClient):
            reviewed = ai_review.review_issue_candidates(result)
        self.assertEqual(reviewed["candidate_count"], 2)
        self.assertEqual(reviewed["group_count"], 2)
        self.assertEqual(reviewed["verification_count"], 1)
        self.assertEqual(reviewed["verified_count"], 1)
        self.assertEqual(result["issues"][0]["ai_decision"], "needs_human")
        self.assertEqual(result["issues"][0]["ai_initial_decision"], "likely_false_positive")
        self.assertTrue(result["issues"][0]["ai_evidence"])
        self.assertEqual(result["issues"][1]["ai_decision"], "confirmed_issue")
        self.assertEqual(len(FakeClient.calls), 3)


if __name__ == "__main__":
    unittest.main()
