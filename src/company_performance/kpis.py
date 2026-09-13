"""Small, source-neutral Company Performance KPI helpers."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from statistics import mean
from typing import Iterable

from .models import StatusInterval, TaskPeriodSnapshot, UnifiedStatus
from .normalization import normalize_priority


def safe_rate(numerator: int, denominator: int) -> float | None:
    return None if denominator == 0 else numerator / denominator * 100.0


def _average(values: Iterable[int]) -> float | None:
    usable = list(values)
    return None if not usable else float(mean(usable))


@dataclass(frozen=True)
class CoreKPIs:
    total_tasks: int
    completed_tasks: int
    cancelled_tasks: int
    rejected_tasks: int
    open_tasks: int
    current_wip: int
    overdue_open_tasks: int
    high_priority_open_tasks: int
    high_priority_overdue_tasks: int
    unassigned_open_tasks: int
    completion_rate: float | None
    overdue_open_rate: float | None
    on_time_completion_rate: float | None
    late_completion_rate: float | None
    average_time_to_start_days: float | None
    average_execution_duration_days: float | None
    average_lead_time_days: float | None


def calculate_core_kpis(snapshots: Iterable[TaskPeriodSnapshot]) -> CoreKPIs:
    tasks = [snapshot for snapshot in snapshots if snapshot.counted_in_kpis]
    completed = [item for item in tasks if item.status_at_period_end is UnifiedStatus.COMPLETED]
    cancelled = [item for item in tasks if item.status_at_period_end is UnifiedStatus.CANCELLED]
    rejected = [item for item in tasks if item.status_at_period_end is UnifiedStatus.REJECTED]
    open_items = [item for item in tasks if item.status_at_period_end.is_open]
    wip = [
        item
        for item in tasks
        if item.status_at_period_end in {UnifiedStatus.IN_EXECUTION, UnifiedStatus.IN_REVIEW}
    ]
    open_with_due = [item for item in open_items if item.task.due_date is not None]
    overdue = [item for item in open_with_due if item.task.due_date < item.period_end]
    high_open = [
        item
        for item in open_items
        if normalize_priority(item.task.source_tool, item.task.priority) in {"Critical", "High"}
    ]
    high_overdue = [
        item
        for item in overdue
        if normalize_priority(item.task.source_tool, item.task.priority) in {"Critical", "High"}
    ]
    unassigned = [item for item in open_items if not item.task.assignees]
    eligible = len(tasks) - len(cancelled) - len(rejected)

    completed_with_due = [
        item for item in completed if item.task.due_date is not None and item.final_completion_date is not None
    ]
    on_time = [item for item in completed_with_due if item.final_completion_date <= item.task.due_date]
    late = [item for item in completed_with_due if item.final_completion_date > item.task.due_date]

    time_to_start = [
        (item.actual_start_date - item.task.created_date).days
        for item in tasks
        if item.actual_start_date is not None and item.task.created_date is not None
    ]
    execution = [
        (item.final_completion_date - item.actual_start_date).days
        for item in completed
        if item.final_completion_date is not None and item.actual_start_date is not None
    ]
    lead = [
        (item.final_completion_date - item.task.created_date).days
        for item in completed
        if item.final_completion_date is not None and item.task.created_date is not None
    ]
    return CoreKPIs(
        total_tasks=len(tasks),
        completed_tasks=len(completed),
        cancelled_tasks=len(cancelled),
        rejected_tasks=len(rejected),
        open_tasks=len(open_items),
        current_wip=len(wip),
        overdue_open_tasks=len(overdue),
        high_priority_open_tasks=len(high_open),
        high_priority_overdue_tasks=len(high_overdue),
        unassigned_open_tasks=len(unassigned),
        completion_rate=safe_rate(len(completed), eligible),
        overdue_open_rate=safe_rate(len(overdue), len(open_with_due)),
        on_time_completion_rate=safe_rate(len(on_time), len(completed_with_due)),
        late_completion_rate=safe_rate(len(late), len(completed_with_due)),
        average_time_to_start_days=_average(time_to_start),
        average_execution_duration_days=_average(execution),
        average_lead_time_days=_average(lead),
    )


def average_days_in_status(
    snapshots: Iterable[TaskPeriodSnapshot], status: UnifiedStatus
) -> float | None:
    """Average clipped calendar days for tasks that entered one status."""
    values = [
        interval.days
        for snapshot in snapshots
        if snapshot.history_available
        for interval in snapshot.status_intervals
        if interval.status is status
    ]
    return _average(values)
