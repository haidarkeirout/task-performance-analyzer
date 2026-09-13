"""Source-independent company task-performance domain helpers.

This package is deliberately isolated from the existing Jira and ClickUp
analysis flows.  Source adapters can translate their collected records into
these dataclasses before the Company Performance UI is connected.
"""

from .models import (
    AssigneeGroup,
    ParentClassification,
    StatusTransition,
    TaskPeriodSnapshot,
    TaskRecord,
    UnifiedStatus,
)

__all__ = [
    "AssigneeGroup",
    "ParentClassification",
    "StatusTransition",
    "TaskPeriodSnapshot",
    "TaskRecord",
    "UnifiedStatus",
]
