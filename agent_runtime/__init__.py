"""Runtime primitives for the KKS audit agent.

The runtime deliberately keeps the audit workflow deterministic.  The model
is a bounded tool used by the workflow, not the owner of file permissions or
report generation.
"""

from .contracts import AgentEvent, AgentPlanStep, AgentStatus, AgentTask
from .orchestrator import KksAuditAgent
from .tools import AiReviewTool, AuditTool, ReportTool, RulesTool, ToolRegistry

__all__ = [
    "AgentEvent",
    "AgentPlanStep",
    "AgentStatus",
    "AgentTask",
    "AuditTool",
    "AiReviewTool",
    "KksAuditAgent",
    "ReportTool",
    "RulesTool",
    "ToolRegistry",
]
