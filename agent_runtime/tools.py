"""Bounded tools exposed to the KKS audit Agent."""
from __future__ import annotations

from pathlib import Path
from typing import Any, Callable, Protocol


ProgressCallback = Callable[[dict[str, Any]], None]


class AgentTool(Protocol):
    name: str

    def execute(
        self,
        input_path: Path,
        output_dir: Path,
        *,
        requested_sheet: str | None = None,
        comparison_path: Path | None = None,
        progress_callback: ProgressCallback | None = None,
    ) -> dict[str, Any]:
        ...


class RulesTool:
    """Run only the read-only workbook and deterministic rule stage."""

    name = "kks.rules"

    def execute(
        self,
        input_path: Path,
        output_dir: Path,
        *,
        requested_sheet: str | None = None,
        comparison_path: Path | None = None,
        progress_callback: ProgressCallback | None = None,
    ) -> dict[str, Any]:
        from run_audit import audit_file

        return audit_file(
            input_path,
            output_dir,
            requested_sheet=requested_sheet,
            comparison_path=comparison_path,
            progress_callback=progress_callback,
            run_ai_review=False,
            write_outputs=False,
        )


class AiReviewTool:
    """Run the AI stage against the rule result, without touching files."""

    name = "kks.ai_review"

    def execute(
        self,
        result: dict[str, Any],
        *,
        progress_callback: ProgressCallback | None = None,
    ) -> dict[str, Any]:
        from run_audit import apply_ai_review_stage

        return apply_ai_review_stage(result, progress_callback=progress_callback)


class ReportTool:
    """Generate the two user-facing artifacts after all decisions are ready."""

    name = "kks.report"

    def execute(
        self,
        input_path: Path,
        output_dir: Path,
        result: dict[str, Any],
        *,
        progress_callback: ProgressCallback | None = None,
    ) -> dict[str, Any]:
        from run_audit import write_audit_artifacts

        return write_audit_artifacts(
            input_path,
            output_dir,
            result,
            progress_callback=progress_callback,
        )


class AuditTool:
    """Adapter around the existing deterministic KKS audit engine.

    Keeping this adapter separate means the rule engine can later be split
    into smaller tools without changing the Agent or HTTP service contracts.
    """

    name = "kks.audit_workbook"

    def execute(
        self,
        input_path: Path,
        output_dir: Path,
        *,
        requested_sheet: str | None = None,
        comparison_path: Path | None = None,
        progress_callback: ProgressCallback | None = None,
    ) -> dict[str, Any]:
        # Lazy import prevents a circular import during frozen-app startup.
        from run_audit import audit_file

        return audit_file(
            input_path,
            output_dir,
            requested_sheet=requested_sheet,
            comparison_path=comparison_path,
            progress_callback=progress_callback,
        )


class ToolRegistry:
    """Explicit allow-list of tools available to this domain Agent."""

    def __init__(self, tools: list[AgentTool] | None = None) -> None:
        self._tools = {
            tool.name: tool
            for tool in (tools or [RulesTool(), AiReviewTool(), ReportTool()])
        }

    def get(self, name: str) -> AgentTool:
        try:
            return self._tools[name]
        except KeyError as exc:
            raise ValueError(f"Agent 工具未注册：{name}") from exc

    def names(self) -> list[str]:
        return sorted(self._tools)
