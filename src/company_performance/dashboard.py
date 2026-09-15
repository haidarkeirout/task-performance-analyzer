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
            formula="Completed tasks / Total tasks × 100",
            calculation=f"{len(completed)} / {len(items)} × 100 = {_format_rate(kpis.completion_rate)}",
            scope=scope,
            period=period,
            exclusions=exclusions,
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



    """The approved executive headline: exactly five cards."""
    return (
        KpiCard("total-tasks", "Total Tasks", str(kpis.total_tasks), "Selected projects and sources"),
        KpiCard("completion-rate", "Completion Rate", _format_rate(kpis.completion_rate),
                f"{kpis.completed_tasks} completed at period end"),
        KpiCard("current-wip", "Current WIP", str(kpis.current_wip),
                "In Execution and In Review"),
        KpiCard("overdue-open", "Overdue Open Tasks", str(kpis.overdue_open_tasks),
                f"{kpis.high_priority_overdue_tasks} high-priority"),
        KpiCard("on-time-rate", "On-Time Completion Rate", _format_rate(kpis.on_time_completion_rate),
                "Completed tasks with a due date"),
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
        )
        for item in snapshots
    )


def _quality_items(snapshots: Sequence[TaskPeriodSnapshot]) -> tuple[DataQualityItem, ...]:
    flags = Counter(flag for item in snapshots for flag in item.data_quality_flags)
    return tuple(DataQualityItem(flag, flags[flag]) for flag in sorted(flags))


def _coverage_views(coverages: Iterable[SourceCoverage] | None) -> tuple[SourceCoverageView, ...]:
    return tuple(
        SourceCoverageView(
            source_tool=item.source_tool,
            source_space=item.source_space,
            unified_project=item.unified_project,
            source_available=item.source_available,
            task_count=item.task_count,
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
    kpis = calculate_core_kpis(selected)
    model = CompanyDashboardModel(
        period_start=period_start,
        period_end=period_end,
        filters=active_filters,
        filter_options=available_filters(all_items),
        kpis=kpis,
        cards=_cards(
            kpis,
            selected,
            period_start=period_start,
            period_end=period_end,
            scope="Company-wide scope",
        ),
        executive_charts=_executive_charts(selected),
        source_coverage=_coverage_views(coverages),
        data_quality=_quality_items(selected),
        task_details=_task_details(selected),
    )
    if len(model.cards) != 6 or len(model.executive_charts) != 3:
        raise AssertionError("Company dashboard contract requires six cards and three charts")
    return model


def render_company_performance_dashboard(st: Any, model: CompanyDashboardModel) -> None:
    """Minimal optional Streamlit renderer for a prepared dashboard model."""
    st.subheader("Company Performance — Selected Projects")
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
