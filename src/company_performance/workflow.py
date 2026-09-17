"""Historical status reconstruction for Company Performance."""

from __future__ import annotations

from dataclasses import replace
from datetime import date, datetime
from typing import Iterable

from .models import StatusInterval, StatusTransition, TaskPeriodSnapshot, TaskRecord, UnifiedStatus
from .normalization import calendar_date, normalize_status


JIRA_EXCEPTION_LABELS = {
    ("in review", "in progress"): "Rework",
    ("in review", "to do"): "Replanning",
    ("in review", "in triage"): "Re-evaluation",
}


def _ordered_events(events: Iterable[StatusTransition]) -> tuple[StatusTransition, ...]:
    return tuple(sorted(events, key=lambda event: event.changed_at))


def _same_day(left: date, right: date) -> bool:
    return left == right


def _history_available(task: TaskRecord, period_end: date, collection_date: date | None) -> bool:
    if task.history_complete:
        coverage_date = calendar_date(task.history_through)
        return bool(
            (task.initial_status or task.workflow_history)
            and coverage_date is not None
            and coverage_date >= period_end
        )
    # A same-day collection can only use a marked snapshot fallback.  It is not
    # treated as historical workflow coverage.
    return False


def _status_as_of(task: TaskRecord, period_end: date, collection_date: date | None, flags: set[str]) -> UnifiedStatus:
    coverage_date = calendar_date(task.history_through)
    if (
        task.history_complete
        and (task.initial_status or task.workflow_history)
        and coverage_date is not None
        and coverage_date >= period_end
    ):
        status = normalize_status(task.source_tool, task.initial_status)
        if task.initial_status is None:
            first = _ordered_events(task.workflow_history)[0]
            status = normalize_status(task.source_tool, first.from_status)
            flags.add("Missing Initial Status")
        for event in _ordered_events(task.workflow_history):
            event_date = calendar_date(event.changed_at)
            if event_date is None:
                flags.add("Invalid Workflow Timestamp")
                continue
            if event_date <= period_end:
                status = normalize_status(task.source_tool, event.to_status)
        if status is UnifiedStatus.UNKNOWN:
            flags.add("Unmapped Status")
        return status

    if collection_date == period_end and task.raw_status:
        flags.add("Snapshot Status Fallback")
        status = normalize_status(task.source_tool, task.raw_status)
        if status is UnifiedStatus.UNKNOWN:
            flags.add("Unmapped Status")
        return status

    if task.history_complete and coverage_date is not None and coverage_date < period_end:
        flags.add("Analysis Period Exceeds History Coverage")
    else:
        flags.add("Missing Workflow History")
    return UnifiedStatus.UNKNOWN


def _active_life_overlaps(task: TaskRecord, period_start: date, period_end: date) -> bool:
    """Determine whether active work existed in the selected date range.

    A terminal status closes the active interval.  A later non-terminal status
    begins a new active interval, which is what makes a reopening in-period
    visible without incorrectly including a task closed before the period.
    """
    if task.created_date is None:
        task.with_flags(["Missing Created Date"])
        return False
    current = normalize_status(task.source_tool, task.initial_status)
    active_start: date | None = task.created_date if not current.is_terminal else None
    events = _ordered_events(task.workflow_history)

    def overlaps(start: date, end: date | None) -> bool:
        right = end or period_end
        return start <= period_end and right >= period_start

    for event in events:
        changed = calendar_date(event.changed_at)
        if changed is None:
            continue
        next_status = normalize_status(task.source_tool, event.to_status)
        if active_start is not None and next_status.is_terminal:
            if overlaps(active_start, changed):
                return True
            active_start = None
        elif active_start is None and not next_status.is_terminal:
            active_start = changed
        current = next_status
    return active_start is not None and overlaps(active_start, None)


def _actual_start(task: TaskRecord, flags: set[str]) -> date | None:
    for event in _ordered_events(task.workflow_history):
        if normalize_status(task.source_tool, event.to_status) is UnifiedStatus.IN_EXECUTION:
            result = calendar_date(event.changed_at)
            if result is not None:
                return result
    # Completion with no actual execution transition must remain unavailable.
    if any(
        normalize_status(task.source_tool, event.to_status) is UnifiedStatus.COMPLETED
        for event in task.workflow_history
    ):
        flags.add("Missing Actual Start")
    return None


def _final_completion(task: TaskRecord, period_end: date) -> date | None:
    """Completion counts only if the reconstructed final period-end state is Done."""
    completed: date | None = None
    for event in _ordered_events(task.workflow_history):
        changed = calendar_date(event.changed_at)
        if changed is None or changed > period_end:
            continue
        if normalize_status(task.source_tool, event.to_status) is UnifiedStatus.COMPLETED:
            completed = changed
    return completed


def _clipped_intervals(
    task: TaskRecord,
    period_start: date,
    period_end: date,
    flags: set[str],
) -> tuple[StatusInterval, ...]:
    if task.created_date is None or not task.history_complete:
        return ()
    status = normalize_status(task.source_tool, task.initial_status)
    if task.initial_status is None:
        flags.add("Missing Initial Status")
        return ()
    interval_start = task.created_date
    result: list[StatusInterval] = []
    for event in _ordered_events(task.workflow_history):
        changed = calendar_date(event.changed_at)
        if changed is None:
            flags.add("Invalid Workflow Timestamp")
            continue
        if changed > period_end:
            break
        clipped_start = max(interval_start, period_start)
        clipped_end = min(changed, period_end)
        if clipped_start <= clipped_end:
            result.append(StatusInterval(status, clipped_start, clipped_end))
        status = normalize_status(task.source_tool, event.to_status)
        interval_start = changed
    clipped_start = max(interval_start, period_start)
    if clipped_start <= period_end:
        result.append(StatusInterval(status, clipped_start, period_end))
    return tuple(result)


def _in_period_events(
    task: TaskRecord, period_start: date, period_end: date
) -> tuple[tuple[str, ...], tuple[date, ...], tuple[date, ...]]:
    exceptions: list[str] = []
    reopened: list[date] = []
    reactivated: list[date] = []
    for event in _ordered_events(task.workflow_history):
        changed = calendar_date(event.changed_at)
        if changed is None or not (period_start <= changed <= period_end):
            continue
        before = normalize_status(task.source_tool, event.from_status)
        after = normalize_status(task.source_tool, event.to_status)
        if task.source_tool.strip().casefold() == "jira":
            exception = JIRA_EXCEPTION_LABELS.get(
                (str(event.from_status or "").strip().casefold(), str(event.to_status or "").strip().casefold())
            )
            if exception:
                exceptions.append(exception)
        if before is UnifiedStatus.COMPLETED and after.is_open:
            reopened.append(changed)
        if before is UnifiedStatus.CANCELLED and after.is_open:
            reactivated.append(changed)
    return tuple(exceptions), tuple(reopened), tuple(reactivated)


def reconstruct_task(
    task: TaskRecord,
    period_start: date,
    period_end: date,
    *,
    collection_date: date | None = None,
) -> TaskPeriodSnapshot:
    """Produce one historical period snapshot without changing source history."""
    if period_end < period_start:
        raise ValueError("period_end must not be before period_start")
    flags = set(task.data_quality_flags)
    in_scope = _active_life_overlaps(task, period_start, period_end)
    history_available = _history_available(task, period_end, collection_date)
    status = _status_as_of(task, period_end, collection_date, flags)
    actual_start = _actual_start(task, flags) if history_available else None
    final_completion = _final_completion(task, period_end) if status is UnifiedStatus.COMPLETED else None
    intervals = _clipped_intervals(task, period_start, period_end, flags) if history_available else ()
    exceptions, reopened, reactivated = _in_period_events(task, period_start, period_end)
    return TaskPeriodSnapshot(
        task=task,
        period_start=period_start,
        period_end=period_end,
        status_at_period_end=status,
        in_scope=in_scope,
        history_available=history_available,
        actual_start_date=actual_start,
        final_completion_date=final_completion,
        status_intervals=intervals,
        exception_events=exceptions,
        reopen_events=reopened,
        reactivation_events=reactivated,
        data_quality_flags=flags,
    )
