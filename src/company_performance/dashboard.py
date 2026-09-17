"""Presentation-ready Company Performance dashboard model.

This module intentionally contains no source collection and does not change
the existing Jira or ClickUp dashboards.  It turns already reconstructed
``TaskPeriodSnapshot`` objects into a small, explicit view model which a
Streamlit page, Excel writer, or another UI can render consistently.
"""

from __future__ import annotations

from collections import Counter
from dataclasses import dataclass
from datetime import date
from typing import Any, Iterable, Sequence

from .adapters import SourceCoverage
from .kpis import CoreKPIs, calculate_core_kpis
from .models import TaskPeriodSnapshot, UnifiedStatus
from .normalization import assignee_group, normalize_priority


_KNOWN_STATUS_ORDER = (
    UnifiedStatus.NOT_STARTED,
    UnifiedStatus.IN_EXECUTION,
    UnifiedStatus.IN_REVIEW,
    UnifiedStatus.AT_RISK,
    UnifiedStatus.ON_HOLD,
    UnifiedStatus.COMPLETED,
    UnifiedStatus.CANCELLED,
    UnifiedStatus.REJECTED,
)
_PRIORITY_ORDER = ("Critical", "High", "Medium", "Low", "Unknown")


@dataclass(frozen=True)
class DashboardFilters:
    """Optional drill-down filters.

    Empty tuples mean "all".  Filtering occurs after the analysis period has
    been reconstructed, so it cannot fabricate historical task states.
    """

    source_tools: tuple[str, ...] = ()
    unified_projects: tuple[str, ...] = ()
    statuses: tuple[UnifiedStatus, ...] = ()
    assignee_groups: tuple[str, ...] = ()
    priorities: tuple[str, ...] = ()


@dataclass(frozen=True)
class FilterOptions:
    source_tools: tuple[str, ...]
    unified_projects: tuple[str, ...]
    statuses: tuple[UnifiedStatus, ...]
    assignee_groups: tuple[str, ...]
    priorities: tuple[str, ...]


@dataclass(frozen=True)
class KpiCard:
    key: str
    title: str
    value: str
    supporting_text: str


@dataclass(frozen=True)
class ChartPoint:
    label: str
    value: int


@dataclass(frozen=True)
class ExecutiveChart:
    key: str
    title: str
    chart_type: str
    points: tuple[ChartPoint, ...]
    note: str | None = None


@dataclass(frozen=True)
class SourceCoverageView:
    source_tool: str
    source_space: str | None
    unified_project: str | None
    source_available: bool
    task_count: int
    history_mode: str
    reason: str | None
    flags: tuple[str, ...]


@dataclass(frozen=True)
class DataQualityItem:
    flag: str
    task_count: int


_QUALITY_REASON_EXPLANATIONS = {
    "Missing Workflow History": "The source did not provide a complete status history through the period end.",
    "Analysis Period Exceeds History Coverage": "The available history ends before the selected period end.",
    "Unmapped Status": "The source status could not be mapped to the approved unified workflow statuses.",
    "Invalid Workflow Timestamp": "At least one workflow event has a missing or invalid timestamp.",
    "Missing Initial Status": "The history has events but no reliable initial status from which to reconstruct the timeline.",
    "ClickUp Chronological History Unavailable": "ClickUp did not provide chronological status history for this task.",
    "Jira History Unavailable": "Jira did not provide verified status history for this task.",
}


def explain_quality_flag(flag: str) -> str:
    """Return a user-facing explanation while preserving the original flag."""
    return _QUALITY_REASON_EXPLANATIONS.get(flag, flag)


def data_quality_task_rows(task_details: Iterable["TaskDetail"]) -> tuple[dict[str, Any], ...]:
    """Build one row per task with every recorded quality reason."""
    rows: list[dict[str, Any]] = []
    for detail in task_details:
        flags = tuple(dict.fromkeys(detail.data_quality_flags))
        if detail.final_status != UnifiedStatus.UNKNOWN.value and not flags:
            continue
        reasons = flags or ("Unknown status without a recorded reason",)
        if detail.final_status == UnifiedStatus.UNKNOWN.value:
            impact = "Excluded from Completion Rate, On-Time, WIP and open/overdue status KPIs"
        else:
            impact = "See recorded reason(s); status-dependent KPIs remain unchanged"
        rows.append({
            "Source Tool": detail.source_tool,
            "Source Space": detail.source_space or "N/A",
            "Project": detail.unified_project or "Unmapped Project",
            "Task ID": detail.task_id,
            "Task Name": detail.task_name or "Untitled task",
            "Status at Period End": detail.final_status,
            "Reason(s)": "; ".join(f"{flag}: {explain_quality_flag(flag)}" for flag in reasons),
            "KPI Impact": impact,
        })
    return tuple(rows)


@dataclass(frozen=True)
class TaskDetail:
    """One row for the details tab, including source and normalised statuses."""

    source_tool: str
    source_space: str | None
    task_id: str
    task_name: str | None
    unified_project: str | None
    original_status: str | None
    final_status: str
    assignee_group: str
    assignees: tuple[str, ...]
    priority: str
    created_date: date | None
    due_date: date | None
    actual_start_date: date | None
    final_completion_date: date | None
    workflow_events: tuple[str, ...]
    exception_events: tuple[str, ...]
    data_quality_flags: tuple[str, ...]
    parent_id: str | None
    parent_classification: str
    is_subtask: bool
    in_analysis_period: bool
    counted_in_kpis: bool
    exclusion_reason: str | None


@dataclass(frozen=True)
class CompanyDashboardModel:
    """A renderer-neutral contract with exactly five cards and three charts."""

    period_start: date
    period_end: date
    filters: DashboardFilters
    filter_options: FilterOptions
    kpis: CoreKPIs
    cards: tuple[KpiCard, ...]
    executive_charts: tuple[ExecutiveChart, ...]
    source_coverage: tuple[SourceCoverageView, ...]
    data_quality: tuple[DataQualityItem, ...]
    task_details: tuple[TaskDetail, ...]
    in_period_task_count: int


def available_filters(snapshots: Iterable[TaskPeriodSnapshot]) -> FilterOptions:
    """Return valid filter values for a dashboard session."""
    items = tuple(snapshots)
    return FilterOptions(
        source_tools=tuple(sorted({item.task.source_tool for item in items if item.task.source_tool})),
        unified_projects=tuple(sorted({item.task.unified_project or "Unmapped Project" for item in items})),
        statuses=tuple(status for status in _KNOWN_STATUS_ORDER if any(
            item.status_at_period_end is status for item in items
        )),
        assignee_groups=tuple(sorted({assignee_group(item.task) for item in items})),
        priorities=tuple(priority for priority in _PRIORITY_ORDER if any(
            normalize_priority(item.task.source_tool, item.task.priority) == priority for item in items
        )),
    )


def filter_snapshots(
    snapshots: Iterable[TaskPeriodSnapshot], filters: DashboardFilters | None = None
) -> tuple[TaskPeriodSnapshot, ...]:
    """Apply dashboard drill-down choices without altering snapshot calculations."""
    filters = filters or DashboardFilters()
    sources = {value.casefold() for value in filters.source_tools}
    projects = set(filters.unified_projects)
    statuses = set(filters.statuses)
    assignees = set(filters.assignee_groups)
    priorities = set(filters.priorities)
    selected: list[TaskPeriodSnapshot] = []
    for item in snapshots:
        task = item.task
        if sources and task.source_tool.casefold() not in sources:
            continue
        if projects and (task.unified_project or "Unmapped Project") not in projects:
            continue
        if statuses and item.status_at_period_end not in statuses:
            continue
        if assignees and assignee_group(task) not in assignees:
            continue
        if priorities and normalize_priority(task.source_tool, task.priority) not in priorities:
            continue
        selected.append(item)
    return tuple(selected)


def _format_rate(value: float | None) -> str:
    return "—" if value is None else f"{value:.1f}%"


def metric_help(
    *,
    formula: str,
    calculation: str,
    scope: str,
    period: str,
    exclusions: str = "None",
    validation: str = "PASS",
) -> str:
    """Render one consistent, human-readable KPI calculation tooltip."""
    return (
        f"**Formula:** {formula}\n\n"
        f"**Calculation:** {calculation}\n\n"
        f"**Scope:** {scope}\n\n"
        f"**Period:** {period}\n\n"
        f"**Exclusions:** {exclusions}\n\n"
        f"**Validation:** {validation}"
    )


def _quality_summary(items: Sequence[TaskPeriodSnapshot]) -> tuple[str, str]:
    flags = sorted({flag for item in items for flag in item.data_quality_flags if flag})
    unknown = sum(item.status_at_period_end is UnifiedStatus.UNKNOWN for item in items)
    if unknown:
        flags.append(f"{unknown} task(s) with unknown period-end status")
    if not flags:
        return "PASS", "No task-level quality flags in the selected scope."
    return "WARNING", "; ".join(dict.fromkeys(flags))


def build_kpi_help(
    snapshots: Sequence[TaskPeriodSnapshot],
    kpis: CoreKPIs,
    *,
    period_start: date,
    period_end: date,
    scope: str,
) -> dict[str, str]:
    """Build calculation explanations from the same snapshots used by KPIs."""
    items = [item for item in snapshots if item.counted_in_kpis]
    period = f"{period_start.isoformat()} to {period_end.isoformat()}"
    validation, quality_note = _quality_summary(items)
    completed = [item for item in items if item.status_at_period_end is UnifiedStatus.COMPLETED]
    completed_with_due = [
        item for item in completed
        if item.task.due_date is not None and item.final_completion_date is not None
    ]
    on_time = [
        item for item in completed_with_due
        if item.final_completion_date <= item.task.due_date
    ]
    known_status = [item for item in items if item.status_at_period_end is not UnifiedStatus.UNKNOWN]
    exclusions = quality_note
    return {
        "total-tasks": metric_help(
            formula="Count distinct Source Tool + Task ID",
            calculation=f"{len(items)} distinct task(s)",
            scope=scope,
            period=period,
            exclusions="Duplicate task keys are removed before KPI calculation.",
            validation=validation,
        ),
        "completion-rate": metric_help(
            formula="Completed tasks / Known-status KPI tasks × 100",
            calculation=f"{len(completed)} / {len(known_status)} × 100 = {_format_rate(kpis.completion_rate)}",
            scope=scope,
            period=period,
            exclusions=f"Unknown-status tasks are excluded from the denominator ({len(items) - len(known_status)} task(s)). {exclusions}",
            validation=validation,
        ),
        "current-wip": metric_help(
            formula="Count of tasks whose period-end status is In Execution or In Review",
            calculation=f"{kpis.current_wip} task(s)",
            scope=scope,
            period=period,
            exclusions="Unknown statuses are not classified as WIP.",
            validation=validation,
        ),
        "overdue-open": metric_help(
            formula="Count of open tasks with a valid due date before period end",
            calculation=f"{kpis.overdue_open_tasks} task(s)",
            scope=scope,
            period=period,
            exclusions="Tasks without a valid due date are excluded from overdue classification.",
            validation=validation,
        ),
        "on-time-rate": metric_help(
            formula="Completed on or before due date / Completed tasks with valid due date × 100",
            calculation=f"{len(on_time)} / {len(completed_with_due)} × 100 = {_format_rate(kpis.on_time_completion_rate)}",
            scope=scope,
            period=period,
            exclusions="Completed tasks missing either due date or completion date are excluded from this KPI.",
            validation="WARNING — " + quality_note if len(on_time) != len(completed_with_due) and quality_note == "No task-level quality flags in the selected scope." else validation,
        ),
    }


def _cards(
    kpis: CoreKPIs,
    snapshots: Sequence[TaskPeriodSnapshot],
    *,
    period_start: date,
    period_end: date,
    scope: str,
) -> tuple[KpiCard, ...]:
    """Build the company headline cards with project count first."""
    help_text = build_kpi_help(
        snapshots,
        kpis,
        period_start=period_start,
        period_end=period_end,
        scope=scope,
    )
    project_count = len({
        item.task.unified_project or "Unmapped Project"
        for item in snapshots
        if item.counted_in_kpis
    })
    return (
        KpiCard("total-projects", "Total Projects", str(project_count),
                "Distinct inferred project names in the selected scope."),
        KpiCard("total-tasks", "Total Tasks", str(kpis.total_tasks), help_text["total-tasks"]),
        KpiCard("completion-rate", "Completion Rate", _format_rate(kpis.completion_rate),
                f"{kpis.completed_tasks} completed at period end\n\n{help_text['completion-rate']}"),
        KpiCard("current-wip", "Current WIP", str(kpis.current_wip), help_text["current-wip"]),
        KpiCard("overdue-open", "Overdue Open Tasks", str(kpis.overdue_open_tasks),
                f"{kpis.high_priority_overdue_tasks} high-priority\n\n{help_text['overdue-open']}"),
        KpiCard("on-time-rate", "On-Time Completion Rate", _format_rate(kpis.on_time_completion_rate),
                help_text["on-time-rate"]),
    )



def _chart(key: str, title: str, points: Sequence[ChartPoint], note: str | None = None) -> ExecutiveChart:
    return ExecutiveChart(key=key, title=title, chart_type="bar", points=tuple(points), note=note)


def _executive_charts(snapshots: Sequence[TaskPeriodSnapshot]) -> tuple[ExecutiveChart, ...]:
    items = [item for item in snapshots if item.counted_in_kpis]
    status_counts = Counter(item.status_at_period_end for item in items)
    status_points = tuple(
        ChartPoint(status.value, status_counts.get(status, 0))
        for status in _KNOWN_STATUS_ORDER
        if status_counts.get(status, 0)
    )

    project_counts = Counter(item.task.unified_project or "Unmapped Project" for item in items)
    project_points = tuple(ChartPoint(label, project_counts[label]) for label in sorted(project_counts))

    overdue = [
        item for item in items
        if item.status_at_period_end.is_open
        and item.task.due_date is not None
        and item.task.due_date < item.period_end
    ]
    priority_counts = Counter(normalize_priority(item.task.source_tool, item.task.priority) for item in overdue)
    priority_points = tuple(
        ChartPoint(priority, priority_counts[priority])
        for priority in _PRIORITY_ORDER
        if priority_counts[priority]
    )
    return (
        _chart(
            "delivery-outcome",
            "Delivery Outcome",
            status_points,
            "How calculated: distinct selected tasks grouped by status at period end.",
        ),
        _chart(
            "workload-by-project",
            "Workload by Unified Project",
            project_points,
            "How calculated: distinct selected tasks grouped by their unified project label.",
        ),
        _chart(
            "overdue-by-priority",
            "Open Overdue Work by Priority",
            priority_points,
            "How calculated: open tasks with valid due dates before period end, grouped by normalized priority.",
        ),
    )


def _workflow_events(snapshot: TaskPeriodSnapshot) -> tuple[str, ...]:
    events: list[str] = []
    for transition in snapshot.task.workflow_history:
        changed = transition.changed_at.date()
        if snapshot.period_start <= changed <= snapshot.period_end:
            before = transition.from_status or "Unknown"
            after = transition.to_status or "Unknown"
            events.append(f"{changed.isoformat()}: {before} → {after}")
    return tuple(events)


def _task_details(snapshots: Sequence[TaskPeriodSnapshot]) -> tuple[TaskDetail, ...]:
    def exclusion_reason(item: TaskPeriodSnapshot) -> str | None:
        if not item.in_scope:
            return "Outside selected analysis period"
        if not item.task.counted_in_kpis:
            return f"Excluded from KPI calculation: {item.task.parent_classification.value}"
        return None

    return tuple(
        TaskDetail(
            source_tool=item.task.source_tool,
            source_space=item.task.source_space,
            task_id=item.task.task_id,
            task_name=item.task.task_name,
            unified_project=item.task.unified_project,
            original_status=item.task.raw_status,
            final_status=item.status_at_period_end.value,
            assignee_group=assignee_group(item.task),
            assignees=item.task.assignees,
            priority=normalize_priority(item.task.source_tool, item.task.priority),
            created_date=item.task.created_date,
            due_date=item.task.due_date,
            actual_start_date=item.actual_start_date,
            final_completion_date=item.final_completion_date,
            workflow_events=_workflow_events(item),
            exception_events=item.exception_events,
            data_quality_flags=tuple(sorted(item.data_quality_flags)),
            parent_id=item.task.parent_id,
            parent_classification=item.task.parent_classification.value,
            is_subtask=item.task.parent_classification.value == "Subtask",
            in_analysis_period=item.in_scope,
            counted_in_kpis=item.counted_in_kpis,
            exclusion_reason=exclusion_reason(item),
        )
        for item in snapshots
    )


def _quality_items(snapshots: Sequence[TaskPeriodSnapshot]) -> tuple[DataQualityItem, ...]:
    flags = Counter(flag for item in snapshots for flag in item.data_quality_flags)
    return tuple(DataQualityItem(flag, flags[flag]) for flag in sorted(flags))


def _coverage_views(
    coverages: Iterable[SourceCoverage] | None,
    snapshots: Iterable[TaskPeriodSnapshot] = (),
) -> tuple[SourceCoverageView, ...]:
    task_counts = Counter(
        (item.task.source_tool, item.task.source_space, item.task.unified_project)
        for item in snapshots
    )
    return tuple(
        SourceCoverageView(
            source_tool=item.source_tool,
            source_space=item.source_space,
            unified_project=item.unified_project,
            source_available=item.source_available,
            task_count=task_counts.get(
                (item.source_tool, item.source_space, item.unified_project), 0
            ),
            history_mode=item.history_mode,
            reason=item.reason,
            flags=item.flags,
        )
        for item in (coverages or ())
    )


def build_company_dashboard(
    snapshots: Iterable[TaskPeriodSnapshot],
    *,
    coverages: Iterable[SourceCoverage] | None = None,
    filters: DashboardFilters | None = None,
) -> CompanyDashboardModel:
    """Build a complete Company Performance dashboard view model.

    The six cards and three charts are fixed by design.  Details, coverage,
    and Data Quality remain available for executive drill-down rather than
    expanding the headline dashboard.
    """
    all_items = tuple(snapshots)
    if not all_items:
        raise ValueError("At least one task snapshot is required to build a dashboard")
    period_start = min(item.period_start for item in all_items)
    period_end = max(item.period_end for item in all_items)
    if any(item.period_start != period_start or item.period_end != period_end for item in all_items):
        raise ValueError("All task snapshots must use the same analysis period")
    active_filters = filters or DashboardFilters()
    selected = filter_snapshots(all_items, active_filters)
    # Collection may retrieve a wider snapshot so that the user can change
    # From/To without recollecting, but the analysis view itself is strictly
    # limited to tasks active during the selected period.
    selected_in_period = tuple(item for item in selected if item.in_scope)
    kpis = calculate_core_kpis(selected_in_period)
    model = CompanyDashboardModel(
        period_start=period_start,
        period_end=period_end,
        filters=active_filters,
        # Keep the available filter choices stable for the session; applying
        # a choice still cannot bring an out-of-period task into the view.
        filter_options=available_filters(all_items),
        kpis=kpis,
        cards=_cards(
            kpis,
            selected_in_period,
            period_start=period_start,
            period_end=period_end,
            scope="Company-wide scope",
        ),
        executive_charts=_executive_charts(selected_in_period),
        source_coverage=_coverage_views(coverages, selected_in_period),
        data_quality=_quality_items(selected_in_period),
        task_details=_task_details(selected_in_period),
        in_period_task_count=len(selected_in_period),
    )
    if len(model.cards) != 6 or len(model.executive_charts) != 3:
        raise AssertionError("Company dashboard contract requires six cards and three charts")
    return model


def render_company_performance_dashboard(st: Any, model: CompanyDashboardModel) -> None:
    """Minimal optional Streamlit renderer for a prepared dashboard model."""
    st.subheader("Company Performance — Company-Wide Scope")
    st.caption(f"Analysis period: {model.period_start.isoformat()} to {model.period_end.isoformat()}")
    columns = st.columns(6)
    for column, card in zip(columns, model.cards):
        column.metric(card.title, card.value, help=card.supporting_text)
    for chart in model.executive_charts:
        st.markdown(f"#### {chart.title}")
        st.bar_chart({point.label: point.value for point in chart.points})
        if chart.note:
            st.caption(chart.note)
    if model.source_coverage:
        st.markdown("#### Source Coverage")
        st.dataframe([item.__dict__ for item in model.source_coverage], hide_index=True)
    if model.data_quality:
        st.markdown("#### Data Quality")
        st.dataframe([item.__dict__ for item in model.data_quality], hide_index=True)
    st.markdown("#### Task Details")
    st.dataframe([item.__dict__ for item in model.task_details], hide_index=True)
