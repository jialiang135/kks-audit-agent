import unittest
from pathlib import Path

from agent_runtime import KksAuditAgent, ToolRegistry


class FakeRulesTool:
    name = "kks.rules"

    def execute(self, input_path, output_dir, **kwargs):
        callback = kwargs.get("progress_callback")
        if callback:
            callback({
                "phase": "audit",
                "stage": "rules_completed",
                "percent": 62,
                "message": "本地规则审核完成",
                "hint": "测试事件",
            })
        return {"conclusion": "测试完成", "issues": []}


class FakeAiTool:
    name = "kks.ai_review"

    def execute(self, result, **kwargs):
        callback = kwargs.get("progress_callback")
        if callback:
            callback({
                "phase": "ai",
                "stage": "completed",
                "percent": 92,
                "message": "AI 复核完成",
            })
        result["ai_review"] = {"status": "completed"}
        return result


class FakeReportTool:
    name = "kks.report"

    def execute(self, input_path, output_dir, result, **kwargs):
        callback = kwargs.get("progress_callback")
        if callback:
            callback({
                "phase": "audit",
                "stage": "report",
                "percent": 95,
                "message": "正在生成审核报告",
            })
        return result


class AgentRuntimeTests(unittest.TestCase):
    def test_plan_is_fixed_and_exposes_only_allowlisted_tool(self):
        agent = KksAuditAgent(tools=ToolRegistry([FakeRulesTool(), FakeAiTool(), FakeReportTool()]))
        self.assertEqual(agent.tools.names(), ["kks.ai_review", "kks.report", "kks.rules"])
        self.assertEqual([item["step_id"] for item in agent.plan()], [
            "ingest", "structure", "rules", "semantic_review", "artifacts",
        ])

    def test_run_relays_agent_stage_and_attaches_runtime_metadata(self):
        agent = KksAuditAgent(tools=ToolRegistry([FakeRulesTool(), FakeAiTool(), FakeReportTool()]))
        events = []
        result = agent.run(
            Path("input.xlsx"),
            Path("outputs"),
            run_id="test-run",
            progress_callback=events.append,
        )
        self.assertEqual(result["agent_runtime"]["mode"], "workflow_first")
        self.assertEqual(result["agent_runtime"]["status"], "completed")
        self.assertTrue(any(event.get("agent_stage") == "rules" for event in events))
        self.assertEqual(events[-1]["agent_status"], "completed")


if __name__ == "__main__":
    unittest.main()
