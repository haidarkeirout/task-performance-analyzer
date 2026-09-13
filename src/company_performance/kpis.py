"""Small, source-neutral Company Performance KPI helpers."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from statistics import mean, median
from typing import Iterable

from .models import StatusInterval, TaskPeriodSnapshot, UnifiedStatus
from .normalization import normalize_priority


def safe_rate(numerator: int, denominator: int) -> float | None:
    """Return a presentation-safe rate using the agreed one-decimal precision."""
    return None if denominator == 0 else round(numerator / denominator * 100.0, 1)


def _average(values: Iterable[int]) -> float | None:
    usable = list(values)
    return None if not usable else round(float(mean(usable)), 1)


def _median(values: Iterable[int]) -> float | None:
    usable = list(values)
    return None if not usable else round(float(median(usable)), 1)


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


@dataclass(frozen=True)
class StatusMetrics:
    """Evidence for one non-terminal workflow status in the selected period.

    All duration values are calendar days already clipped to the selected
    period.  A caller should only use these metrics when ``history_covered``
    is true; this makes ClickUp's snapshot-only coverage explicit rather than
    inventing historical timing.
    """

    status: UnifiedStatus
    tasks_passed_through: int
    average_days: float | None
    median_days: float | None
    total_days: int
    open_tasks_now: int
    overdue_open_tasks: int
    repeated_returns: int
    history_covered: bool


@dataclass(frozen=True)
class BottleneckCandidate:
    """A workflow stage that warrants follow-up, never a confirmed cause."""

    status: UnifiedStatus
    strength: str
    evidence: tuple[str, ...]
    metrics: StatusMetrics


@dataclass(frozen=True)
class Recommendation:
    """An evidence-linked, actionable recommendation for an executive report."""

    code: str
    severity: str
    title: str
    evidence: str
    suggested_action: str


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


def calculate_status_metrics(snapshots: Iterable[TaskPeriodSnapshot]) -> tuple[StatusMetrics, ...]:
    """Summarise historical time-in-status without creating synthetic history.

    Completed, Cancelled and Rejected are deliberately excluded from workflow
    bottleneck comparison.  Unknown status is handled through Data Quality,
    not as a chart category.
    """
    items = [snapshot for snapshot in snapshots if snapshot.counted_in_kpis]
    result: list[StatusMetrics] = []
    non_terminal = [status for status in UnifiedStatus if status.is_open]
    for status in non_terminal:
        durations = [
            interval.days
            for snapshot in items
            if snapshot.history_available
            for interval in snapshot.status_intervals
            if interval.status is status
        ]
        open_now = [item for item in items if item.status_at_period_end is status]
        overdue_now = [
            item for item in open_now
            if item.task.due_date is not None and item.task.due_date < item.period_end
        ]
        # Rework/replanning/re-evaluation are review-return signals.  They are
        # intentionally attached to In Review, their common source stage.
        returns = sum(
            len(item.exception_events) for item in items if status is UnifiedStatus.IN_REVIEW
        )
        result.append(
            StatusMetrics(
                status=status,
                tasks_passed_through=len(durations),
                average_days=_average(durations),
                median_days=_median(durations),
                total_days=sum(durations),
                open_tasks_now=len(open_now),
                overdue_open_tasks=len(overdue_now),
                repeated_returns=returns,
                history_covered=bool(durations),
            )
        )
    return tuple(result)


def identify_bottleneck_candidates(
    snapshots: Iterable[TaskPeriodSnapshot],
) -> tuple[BottleneckCandidate, ...]:
    """Find evidence-backed workflow bottleneck candidates.

    There is deliberately no fixed "five days means bottleneck" rule.  A
    stage becomes a candidate only when at least two relative signals agree:
    above-peer time in stage, concentrated open work, overdue work, or
    repeated review returns.  Three or more signals make it *Strong*; neither
    level claims root cause or a confirmed bottleneck.
    """
    metrics = calculate_status_metrics(snapshots)
    comparable_averages = [metric.average_days for metric in metrics if metric.average_days is not None]
    peer_average = _average(value for value in comparable_averages if value is not None)
    highest_open = max((metric.open_tasks_now for metric in metrics), default=0)
    candidates: list[BottleneckCandidate] = []
    for metric in metrics:
        evidence: list[str] = []
        if (
            metric.average_days is not None
            and peer_average is not None
            and metric.average_days >= peer_average
            and metric.tasks_passed_through > 0
        ):
            evidence.append("Above-peer average time in status")
        if metric.open_tasks_now > 0 and metric.open_tasks_now == highest_open:
            evidence.append("Highest current open-work concentration")
        if metric.overdue_open_tasks > 0:
            evidence.append("Open overdue tasks in status")
        if metric.repeated_returns > 0:
            evidence.append("Repeated workflow returns from review")
        if len(evidence) >= 2:
            candidates.append(
                BottleneckCandidate(
                    status=metric.status,
                    strength="Strong Bottleneck Candidate" if len(evidence) >= 3 else "Potential Bottleneck Candidate",
                    evidence=tuple(evidence),
                    metrics=metric,
                )
            )
    return tuple(candidates)


def generate_recommendations(snapshots: Iterable[TaskPeriodSnapshot]) -> tuple[Recommendation, ...]:
    """Generate recommendations from agreed service thresholds and evidence.

    Completion <70% and on-time completion <80% produce recommendations only;
    they are not bottleneck thresholds and do not prove a root cause.
    """
    items = tuple(snapshots)
    core = calculate_core_kpis(items)
    recommendations: list[Recommendation] = []
    if core.completion_rate is not None and core.completion_rate < 70.0:
        recommendations.append(
            Recommendation(
                code="completion-rate",
                severity="High",
                title="Improve completion flow",
                evidence=f"Completion rate is {core.completion_rate:.1f}%, below the 70.0% recommendation threshold.",
                suggested_action="Review the oldest open work and agree an owner and next step for each item.",
            )
        )
    if core.on_time_completion_rate is not None and core.on_time_completion_rate < 80.0:
        recommendations.append(
            Recommendation(
                code="on-time-rate",
                severity="High",
                title="Protect delivery dates",
                evidence=f"On-time completion rate is {core.on_time_completion_rate:.1f}%, below the 80.0% recommendation threshold.",
                suggested_action="Review due dates for active work and escalate items that cannot meet their committed date.",
            )
        )
    if core.high_priority_overdue_tasks > 0:
        recommendations.append(
            Recommendation(
                code="priority-overdue",
                severity="High",
                title="Escalate overdue high-priority work",
                evidence=f"{core.high_priority_overdue_tasks} high-priority open task(s) are overdue.",
                suggested_action="Assign an accountable owner and recovery date, then track the item to closure.",
            )
        )
    elif core.overdue_open_tasks > 0:
        recommendations.append(
            Recommendation(
                code="open-overdue",
                severity="Medium",
                title="Recover overdue open work",
                evidence=f"{core.overdue_open_tasks} open task(s) are past their due date.",
                suggested_action="Validate scope, due date and next action for each overdue task.",
            )
        )
    missing_history = sum(
        1 for item in items if "Missing Workflow History" in item.data_quality_flags
    )
    if missing_history:
        recommendations.append(
            Recommendation(
                code="history-coverage",
                severity="Medium",
                title="Improve historical workflow coverage",
                evidence=f"{missing_history} task(s) cannot be used in historical status KPIs because workflow history is missing.",
                suggested_action="Collect complete workflow history before using historical status comparisons for these tasks.",
            )
        )
    return tuple(recommendations)
