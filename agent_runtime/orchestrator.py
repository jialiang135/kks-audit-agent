"""Workflow-first orchestration for the KKS audit Agent."""
from __future__ import annotations

import logging
from pathlib import Path
from typing import Any, Callable

from .contracts import AgentEvent, AgentPlanStep, AgentStatus, AgentTask
from .tools import ProgressCallback, ToolRegistry


LOGGER = logging.getLogger("kks-audit.agent")
AgentCallback = Callable[[dict[str, Any]], None]


class KksAuditAgent:
    """Run one read-only KKS audit as a traceable Agent task.

    The plan is intentionally fixed and domain-specific.  This is safer than
    allowing a model to invent file operations or skip mandatory audit stages.
    The workbook, AI review and report stages are separate allow-listed tools.
    The model can only participate through the AI review tool and cannot
    modify the source workbook or choose arbitrary file operations.
    """

    PLAN = (
        AgentPlanStep("ingest", "读取 Excel", "kks.rules", 8, 20),
        AgentPlanStep("structure", "识别结构", "kks.rules", 20, 35),
        AgentPlanStep("rules", "执行规则", "kks.rules", 35, 62),
        AgentPlanStep("semantic_review", "AI 语义复核", "kks.ai_review", 62, 92),
        AgentPlanStep("artifacts", "生成审核结果", "kks.report", 92, 100),
    )

    _AUDIT_STAGE_MAP = {
        "reading": "ingest",
        "structure": "structure",
        "rules_completed": "rules",
        "candidate_scan": "semantic_review",
        "global_review": "semantic_review",
        "model_ready": "semantic_review",
        "review_running": "semantic_review",
        "verification_running": "semantic_review",
        "final_adjudication": "semantic_review",
        "disabled": "semantic_review",
        "error": "semantic_review",
        "completed": "artifacts",
        "report": "artifacts",
    }

    def __init__(self, *, tools: ToolRegistry | None = None) -> None:
        self.tools = tools or ToolRegistry()

    @classmethod
    def plan(cls) -> list[dict[str, Any]]:
        return [step.as_dict() for step in cls.PLAN]

    @classmethod
    def _step_for_stage(cls, stage: str, phase: str = "") -> AgentPlanStep:
        # AI also emits a generic "completed" stage.  Keep that event in the
        # semantic step instead of incorrectly moving the UI to artifacts.
        if phase == "ai" and stage == "completed":
            step_id = "semantic_review"
        else:
            step_id = cls._AUDIT_STAGE_MAP.get(stage, "rules")
        return next(step for step in cls.PLAN if step.step_id == step_id)

    def run(
        self,
        input_path: Path,
        output_dir: Path,
        *,
        run_id: str = "local",
        requested_sheet: str | None = None,
        comparison_path: Path | None = None,
        progress_callback: AgentCallback | None = None,
    ) -> dict[str, Any]:
        task = AgentTask(run_id=run_id)

        def emit(
            status: AgentStatus,
            stage: str,
            message: str,
            hint: str = "",
            percent: int = 0,
            *,
            tool: str = "",
            data: dict[str, Any] | None = None,
        ) -> None:
            task.sequence += 1
            task.status = status
            task.current_step = stage
            event = AgentEvent(
                run_id=run_id,
                sequence=task.sequence,
                status=status,
                stage=stage,
                message=message,
                hint=hint,
                percent=max(0, min(100, int(percent))),
                tool=tool,
                data=data or {},
            ).as_dict()
            if progress_callback:
                progress_callback(event)

        emit(
            AgentStatus.PLANNING,
            "planning",
            "正在创建审核计划",
            "按 KKS 审核流程读取文件、执行规则、复核疑难问题并生成报告",
            3,
            data={"plan": self.plan(), "tools": self.tools.names()},
        )
        rules_tool = self.tools.get("kks.rules")
        ai_tool = self.tools.get("kks.ai_review")
        report_tool = self.tools.get("kks.report")
        emit(
            AgentStatus.RUNNING,
            "ingest",
            "审核任务已启动",
            "原始 Excel 以只读方式处理",
            5,
            tool=rules_tool.name,
        )

        def relay(event: dict[str, Any], tool_name: str) -> None:
            stage = str(event.get("stage", "running"))
            phase = str(event.get("phase", ""))
            step = self._step_for_stage(stage, phase)
            task.current_step = step.step_id
            is_completed = (
                (phase == "audit" and stage == "rules_completed")
                or (phase == "ai" and stage == "completed")
                or stage == "report"
            )
            if step.step_id not in task.completed_steps and is_completed:
                task.completed_steps.append(step.step_id)
            forwarded = dict(event)
            forwarded.update(
                {
                    "source_phase": event.get("phase", "audit"),
                    "agent_status": AgentStatus.RUNNING.value,
                    "agent_stage": step.step_id,
                    "agent_sequence": task.sequence + 1,
                    "tool": tool_name,
                    "plan_step": step.as_dict(),
                }
            )
            task.sequence += 1
            if progress_callback:
                progress_callback(forwarded)

        try:
            result = rules_tool.execute(
                input_path,
                output_dir,
                requested_sheet=requested_sheet,
                comparison_path=comparison_path,
                progress_callback=lambda event: relay(event, rules_tool.name),
            )
            result = ai_tool.execute(
                result,
                progress_callback=lambda event: relay(event, ai_tool.name),
            )
            result = report_tool.execute(
                input_path,
                output_dir,
                result,
                progress_callback=lambda event: relay(event, report_tool.name),
            )
            task.status = AgentStatus.COMPLETED
            task.current_step = "completed"
            task.completed_steps = [step.step_id for step in self.PLAN]
            result["agent_runtime"] = {
                **task.as_dict(),
                "plan": self.plan(),
                "tools": self.tools.names(),
                "mode": "workflow_first",
            }
            emit(
                AgentStatus.COMPLETED,
                "completed",
                "审核任务已完成",
                "HTML 报告和 Excel 问题清单已生成",
                100,
                tool=report_tool.name,
            )
            LOGGER.info("agent_task_done run_id=%s tools=%s", run_id, ",".join(self.tools.names()))
            return result
        except Exception as exc:
            task.status = AgentStatus.FAILED
            task.error = str(exc)
            emit(
                AgentStatus.FAILED,
                "failed",
                "审核任务失败",
                str(exc),
                0,
                tool=task.current_step,
            )
            LOGGER.exception("agent_task_failed run_id=%s", run_id)
            raise
