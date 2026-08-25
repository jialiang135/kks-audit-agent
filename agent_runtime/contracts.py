"""Small, serialisable contracts shared by the Agent runtime and UI."""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from enum import Enum
from typing import Any


class AgentStatus(str, Enum):
    CREATED = "created"
    PLANNING = "planning"
    RUNNING = "running"
    COMPLETED = "completed"
    FAILED = "failed"


@dataclass(frozen=True)
class AgentPlanStep:
    """A visible business step, not an arbitrary model-generated plan."""

    step_id: str
    label: str
    tool: str
    percent_start: int
    percent_end: int

    def as_dict(self) -> dict[str, Any]:
        return {
            "step_id": self.step_id,
            "label": self.label,
            "tool": self.tool,
            "percent_start": self.percent_start,
            "percent_end": self.percent_end,
        }


@dataclass
class AgentTask:
    """In-memory task state used by the HTTP service and CLI runner."""

    run_id: str
    name: str = "kks-coding-audit"
    status: AgentStatus = AgentStatus.CREATED
    current_step: str = "created"
    sequence: int = 0
    completed_steps: list[str] = field(default_factory=list)
    error: str = ""

    def as_dict(self) -> dict[str, Any]:
        return {
            "name": self.name,
            "status": self.status.value,
            "current_step": self.current_step,
            "sequence": self.sequence,
            "completed_steps": list(self.completed_steps),
            "error": self.error,
        }


@dataclass(frozen=True)
class AgentEvent:
    """A progress event emitted by the Agent runtime."""

    run_id: str
    sequence: int
    status: AgentStatus
    stage: str
    message: str
    hint: str = ""
    percent: int = 0
    tool: str = ""
    data: dict[str, Any] = field(default_factory=dict)
    timestamp: str = field(default_factory=lambda: datetime.now().isoformat(timespec="seconds"))

    def as_dict(self) -> dict[str, Any]:
        payload: dict[str, Any] = {
            "phase": "agent",
            "agent_status": self.status.value,
            "agent_stage": self.stage,
            "sequence": self.sequence,
            "percent": self.percent,
            "message": self.message,
            "hint": self.hint,
            "timestamp": self.timestamp,
        }
        if self.tool:
            payload["tool"] = self.tool
        payload.update(self.data)
        return payload
