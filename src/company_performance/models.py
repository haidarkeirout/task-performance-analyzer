"""Typed records used by Company Performance calculations."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, datetime
from enum import Enum
from typing import Iterable


class UnifiedStatus(str, Enum):
    NOT_STARTED = "Not Started"
    IN_EXECUTION = "In Execution"
    IN_REVIEW = "In Review"
    ON_HOLD = "On Hold"
    AT_RISK = "At Risk"
    COMPLETED = "Completed"
    CANCELLED = "Cancelled"
    REJECTED = "Rejected"
    UNKNOWN = "Unknown"

    @property
    def is_terminal(self) -> bool:
        return self in {self.COMPLETED, self.CANCELLED, self.REJECTED}

    @property
    def is_open(self) -> bool:
        return self in {
            self.NOT_STARTED,
            self.IN_EXECUTION,
            self.IN_REVIEW,
            self.ON_HOLD,
            self.AT_RISK,
        }


class ParentClassification(str, Enum):
    CONTAINER = "Container Parent"
    INDEPENDENT = "Independent / Operational Parent"
    SUBTASK = "Subtask"
    STANDALONE = "Standalone Task"

    @property
    def counted_in_kpis(self) -> bool:
        # Subtasks remain visible and are included in the overall task count,
        # but parent-level performance KPIs must not count them a second time.
        return self not in {self.CONTAINER, self.SUBTASK}


class AssigneeGroup(str, Enum):
    UNASSIGNED = "Unassigned"
    MULTIPLE = "Multiple Assignees"


@dataclass(frozen=True, order=True)
class StatusTransition:
    """One source status transition, stored without changing source labels."""

    changed_at: datetime
    from_status: str | None
    to_status: str | None
    performed_by: str | None = None


@dataclass
class TaskRecord:
    """A source-adapter-neutral task record.

    Dates are normalised to calendar dates in ``Asia/Damascus`` by the
    normalisation module.  The original status names remain available so that
    workflow exceptions retain their source-specific meaning.
    """

    source_tool: str
    task_id: str
    task_name: str | None = None
    issue_type: str | None = None
    source_space: str | None = None
    unified_project: str | None = None
    department: str | None = None
    raw_status: str | None = None
    initial_status: str | None = None
    priority: str | None = None
    assignees: tuple[str, ...] = ()
    created_date: date | None = None
    planned_start_date: date | None = None
    due_date: date | None = None
    parent_id: str | None = None
    parent_classification: ParentClassification = ParentClassification.STANDALONE
    collection_timestamp: datetime | None = None
    history_through: datetime | None = None
    workflow_history: tuple[StatusTransition, ...] = ()
    history_complete: bool = False
    data_quality_flags: set[str] = field(default_factory=set)

    @property
    def unique_key(self) -> tuple[str, str]:
        return (self.source_tool.strip().casefold(), str(self.task_id).strip())

    @property
    def counted_in_kpis(self) -> bool:
        return self.parent_classification.counted_in_kpis

    @property
    def assignee_group(self) -> str:
        if not self.assignees:
            return AssigneeGroup.UNASSIGNED.value
        if len(self.assignees) > 1:
            return AssigneeGroup.MULTIPLE.value
        return self.assignees[0]

    def with_flags(self, flags: Iterable[str]) -> "TaskRecord":
        self.data_quality_flags.update(flag for flag in flags if flag)
        return self


@dataclass(frozen=True)
class StatusInterval:
    """A calendar-date status interval already clipped to the analysis period."""

    status: UnifiedStatus
    start_date: date
    end_date: date

    @property
    def days(self) -> int:
        return max(0, (self.end_date - self.start_date).days)


@dataclass
class TaskPeriodSnapshot:
    """The historical view of one task for one selected period."""

    task: TaskRecord
    period_start: date
    period_end: date
    status_at_period_end: UnifiedStatus
    in_scope: bool
    history_available: bool
    actual_start_date: date | None = None
    final_completion_date: date | None = None
    status_intervals: tuple[StatusInterval, ...] = ()
    exception_events: tuple[str, ...] = ()
    reopen_events: tuple[date, ...] = ()
    reactivation_events: tuple[date, ...] = ()
    data_quality_flags: set[str] = field(default_factory=set)

    @property
    def counted_in_kpis(self) -> bool:
        return self.task.counted_in_kpis and self.in_scope

    @property
    def is_open(self) -> bool:
        return self.status_at_period_end.is_open
