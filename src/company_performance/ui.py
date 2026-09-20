"""Additive Streamlit screens for Company Performance.

The functions receive Streamlit as an argument so the company domain remains
testable without a running Streamlit app.  They never invoke Jira or ClickUp
APIs; source collection remains in the established collector screens.
"""

from __future__ import annotations

import tempfile
import hashlib
from collections import defaultdict
from datetime import date, timedelta
from pathlib import Path
from typing import Any

import pandas as pd

from .application import (
    CompanyAnalysisResult,
    build_company_analysis,
    build_company_preview,
    filter_company_preview,
)
from .dashboard import DashboardFilters, build_company_dashboard, data_quality_task_rows
from .kpis import generate_recommendations, identify_bottleneck_candidates
from .models import UnifiedStatus
from .outputs import write_company_excel, write_company_raw_data, write_company_word_report
from kpi_transparency import population_from_snapshots, render_population_card


_MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
_MIME_DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"


def _registry_key(source: str) -> str:
    return "company_prepared_jira_spaces" if source == "Jira" else "company_prepared_clickup_spaces"


def _latest_key(source: str) -> str:
    return "company_prepared_jira" if source == "Jira" else "company_prepared_clickup"


def remember_prepared_source(session_state: Any, source: str, prepared: Any | None) -> None:
    """Keep prepared results available to Company Performance without recollecting.

    The legacy single-source keys remain populated for backwards compatibility.
    A small per-source registry additionally keeps previously collected spaces so
    Company Performance can select Jira, ClickUp, or both independently.
    """
    if prepared is None:
        return
    normalized_source = "Jira" if str(source).casefold() == "jira" else "ClickUp"
    session_state[_latest_key(normalized_source)] = prepared

    space_name = str(getattr(prepared, "space_name", None) or "Selected space")
    registry = dict(session_state.get(_registry_key(normalized_source), {}))
    registry[space_name] = prepared
    session_state[_registry_key(normalized_source)] = registry


def available_prepared_spaces(session_state: Any, source: str) -> dict[str, Any]:
    """Return the prepared spaces available for one Company Performance source."""
    normalized_source = "Jira" if str(source).casefold() == "jira" else "ClickUp"
    registry = dict(session_state.get(_registry_key(normalized_source), {}))
    latest = session_state.get(_latest_key(normalized_source))
    if latest is not None:
        space_name = str(getattr(latest, "space_name", None) or "Selected space")
        registry.setdefault(space_name, latest)
    return registry


def _period_defaults(prepared_items: tuple[Any, ...]) -> tuple[date, date]:
    dates: list[date] = []
    for item in prepared_items:
        value = getattr(item, "cutoff", None) or getattr(item, "collected_at", None)
        try:
            parsed = date.fromisoformat(str(value)[:10])
        except (TypeError, ValueError):
            continue
        dates.append(parsed)
    end = max(dates) if dates else date.today()
    return end.replace(day=1), end


def _preview_table_rows(rows: Any) -> list[dict[str, Any]]:
    return [
        {
            "Company": row.company_name or row.project_name,
            "Project": row.project_name,
            "Department": row.department,
            "Issue Type": row.issue_type,
            "Epic / Workstream": row.epic_name,
            "Parent / Epic": row.parent_id,
            "Hierarchy Role": row.parent_classification,
            "Source Tool": row.source_tool,
            "Space": row.space,
            "Task Name": row.task_name,
            "Original Status": row.original_status,
            "Final Status": row.final_status,
            "Assignee": row.assignee,
            "Priority": row.priority,
            "Due Date": row.due_date,
        }
        for row in rows
    ]


def _preview_filters(st: Any, rows: Any):
    if not rows:
        return rows

    first = st.columns(4)
    sources = first[0].multiselect(
        "Source Tool",
        sorted({row.source_tool for row in rows}),
        key="company_preview_source",
    )
    spaces = first[1].multiselect(
        "Space",
        sorted({row.space for row in rows if row.space}),
        key="company_preview_space",
    )
    task_name = first[2].text_input("Task Name", key="company_preview_task_name")
    original_statuses = first[3].multiselect(
        "Original Status",
        sorted({row.original_status for row in rows if row.original_status}),
        key="company_preview_original_status",
    )

    second = st.columns(4)
    final_statuses = second[0].multiselect(
        "Final Status",
        sorted({row.final_status for row in rows}),
        key="company_preview_final_status",
    )
    assignees = second[1].multiselect(
        "Assignee",
        sorted({row.assignee for row in rows}),
        key="company_preview_assignee",
    )
    priorities = second[2].multiselect(
        "Priority",
        sorted({row.priority for row in rows}),
        key="company_preview_priority",
    )

    due_dates = [row.due_date for row in rows if row.due_date]
    due_range = ()
    if due_dates:
        due_range = second[3].date_input(
            "Due Date",
            value=(),
            min_value=min(due_dates),
            max_value=max(due_dates),
            key="company_preview_due_date",
        )
    due_start = None
    due_end = None
    if isinstance(due_range, (tuple, list)) and len(due_range) == 2:
        due_start, due_end = due_range

    return filter_company_preview(
        rows,
        source_tools=sources,
        spaces=spaces,
        task_name=task_name,
        original_statuses=original_statuses,
        final_statuses=final_statuses,
        assignees=assignees,
        priorities=priorities,
        due_date_start=due_start,
        due_date_end=due_end,
    )


def render_company_launcher(st: Any) -> None:
    """Render the automatic Company preview and date-only run controls."""
    prepared_items = tuple(st.session_state.get("company_prepared_items") or ())
    has_analysis = st.session_state.get("company_analysis") is not None
    selected_company_name = st.session_state.get("company_selected_company_name", "Company")
    with st.expander(
        f"Company Performance — {selected_company_name}"
        if not has_analysis else "Collected Tasks & Analysis Period",
        expanded=bool(prepared_items) and not has_analysis,
    ):
        if not prepared_items:
            if st.session_state.get("company_collection_error"):
                st.info("Company collection could not prepare a usable preview. Retry above.")
            else:
                st.info("Collecting all Jira projects and ClickUp Spaces...")
            return

        st.caption(
            f"Scope: {selected_company_name}. The original source Space/Project remains "
            "visible in task details; Company and Project grouping are kept separately."
        )

        project_rows = []
        grouped: dict[str, list[str]] = defaultdict(list)
        for item in prepared_items:
            label = str(getattr(item, "project_name", "") or "Unmapped Project")
            source = str(getattr(item, "source_tool", "") or "")
            space = str(getattr(item, "source_space", "") or "")
            grouped[label].append(f"{source}: {space}")
        for project_name, source_spaces in sorted(grouped.items(), key=lambda pair: pair[0].casefold()):
            project_rows.append({
                "Project": project_name,
                "Source Spaces": " | ".join(sorted(source_spaces)),
            })

        st.subheader("Detected Company Sources")
        st.dataframe(project_rows, hide_index=True, use_container_width=True)

        jira_items = tuple(item for item in prepared_items if getattr(item, "source_tool", "") == "Jira")
        clickup_items = tuple(item for item in prepared_items if getattr(item, "source_tool", "") == "ClickUp")
        preview = build_company_preview(
            jira_prepared=jira_items,
            clickup_prepared=clickup_items,
        )
        st.subheader("Complete Task Preview")
        filtered_preview = _preview_filters(st, preview)
        st.caption(
            f"Showing {len(filtered_preview)} of {len(preview)} collected task(s) "
            f"across {len(prepared_items)} source Space(s). Preview filters do not change the analysis scope."
        )
        st.dataframe(
            _preview_table_rows(filtered_preview),
            hide_index=True,
            use_container_width=True,
        )

        default_start, default_end = _period_defaults(prepared_items)
        period_columns = st.columns(2)
        period_start = period_columns[0].date_input(
            "Analysis Period From",
            value=default_start,
            key="company_period_start",
        )
        period_end = period_columns[1].date_input(
            "Analysis Period To",
            value=default_end,
            key="company_period_end",
        )
        dates_ready = period_start is not None and period_end is not None
        if st.button(
            "Run Company Analysis" if not has_analysis else "Re-run Company Analysis",
            type="primary",
            key="run_company_performance" if not has_analysis else "rerun_company_performance",
            disabled=not dates_ready,
        ):
            try:
                st.session_state["company_analysis"] = build_company_analysis(
                    period_start=period_start,
                    period_end=period_end,
                    jira_prepared=jira_items,
                    clickup_prepared=clickup_items,
                )
                st.success("Company Performance analysis completed.")
            except (ValueError, TypeError) as exc:
                st.error(str(exc))


def _table_rows(items: Any) -> list[dict[str, Any]]:
    return [item.__dict__ for item in items]


def _task_detail_rows(items: Any) -> list[dict[str, Any]]:
    return [
        {
            "Source Tool": item.source_tool,
            "Space": item.source_space,
            "Task ID": item.task_id,
            "Task Name": item.task_name,
            "Unified Project": item.unified_project,
            "Original Status": item.original_status,
            "Final Status": item.final_status,
            "Assignee Group": item.assignee_group,
            "Assignees": item.assignees,
            "Priority": item.priority,
            "Created Date": item.created_date,
            "Due Date": item.due_date,
            "Actual Start Date": item.actual_start_date,
            "Final Completion Date": item.final_completion_date,
            "Workflow Events": item.workflow_events,
            "Exception Events": item.exception_events,
            "Data Quality Flags": item.data_quality_flags,
            "Parent ID": item.parent_id,
            "Parent Classification": item.parent_classification,
            "Is Subtask": item.is_subtask,
            "In Analysis Period": item.in_analysis_period,
            "Counted in KPIs": item.counted_in_kpis,
            "Exclusion Reason": item.exclusion_reason,
        }
        for item in items
    ]


def _filters(st: Any, result: CompanyAnalysisResult) -> DashboardFilters:
    options = build_company_dashboard(
        result.snapshots,
        coverages=result.collection.coverages,
    ).filter_options
    columns = st.columns(5)
    sources = columns[0].multiselect("Source", options.source_tools, key="company_filter_sources")
    projects = columns[1].multiselect("Unified Project", options.unified_projects, key="company_filter_projects")
    statuses = columns[2].multiselect(
        "Final Status", [status.value for status in options.statuses], key="company_filter_statuses"
    )
    assignees = columns[3].multiselect("Assignee", options.assignee_groups, key="company_filter_assignees")
    priorities = columns[4].multiselect("Priority", options.priorities, key="company_filter_priorities")
    lookup = {status.value: status for status in UnifiedStatus}
    return DashboardFilters(
        source_tools=tuple(sources),
        unified_projects=tuple(projects),
        statuses=tuple(lookup[value] for value in statuses),
        assignee_groups=tuple(assignees),
        priorities=tuple(priorities),
    )


def _output_bytes(result: CompanyAnalysisResult, model: Any) -> tuple[bytes, bytes, bytes]:
    with tempfile.TemporaryDirectory() as directory:
        root = Path(directory)
        selected_keys = {(detail.source_tool, detail.task_id) for detail in model.task_details}
        selected_snapshots = tuple(
            snapshot for snapshot in result.snapshots
            if snapshot.in_scope and (snapshot.task.source_tool, snapshot.task.task_id) in selected_keys
        )
        bottlenecks = identify_bottleneck_candidates(selected_snapshots)
        recommendations = generate_recommendations(selected_snapshots)
        excel = write_company_excel(
            model,
            root / "company_performance_analysis.xlsx",
            snapshots=selected_snapshots,
            bottlenecks=bottlenecks,
            recommendations=recommendations,
        ).read_bytes()
        raw = write_company_raw_data(
            selected_snapshots,
            root / "company_performance_raw_data.xlsx",
        ).read_bytes()
        word = write_company_word_report(
            model,
            root / "company_performance_report.docx",
            snapshots=selected_snapshots,
            bottlenecks=bottlenecks,
            recommendations=recommendations,
        ).read_bytes()
    return excel, raw, word


def _cached_output_bytes(
    st: Any,
    result: CompanyAnalysisResult,
    model: Any,
    *,
    cache_prefix: str = "company",
) -> tuple[bytes, bytes, bytes]:
    selected = sorted((detail.source_tool, detail.task_id) for detail in model.task_details)
    material = repr((result.snapshots[0].period_start if result.snapshots else None,
                     result.snapshots[0].period_end if result.snapshots else None, selected))
    key = hashlib.sha256(material.encode("utf-8")).hexdigest()
    cache_key = f"{cache_prefix}:{key}"
    if st.session_state.get("company_output_cache_key") != cache_key:
        st.session_state["company_output_cache"] = _output_bytes(result, model)
        st.session_state["company_output_cache_key"] = cache_key
    return st.session_state["company_output_cache"]


def _average_text(values: list[int | float]) -> str:
    usable = [float(value) for value in values if value is not None]
    return "Unavailable" if not usable else f"{sum(usable) / len(usable):.1f} days"


def _management_average_rows(snapshots: Any, period_end: date) -> list[dict[str, str]]:
    if not snapshots:
        return [
            {"Metric": label, "Value": "Unavailable"}
            for label in (
                "Avg Execution Time (Completed)",
                "Avg Lead Time (Completed)",
                "Avg Time to Start",
                "Avg Late Completion",
                "Avg Open Overdue",
                "Avg Due Variance (Completed)",
            )
        ]
    core = build_company_dashboard(snapshots).kpis
    completed = [
        item for item in snapshots
        if item.counted_in_kpis
        and item.status_at_period_end is UnifiedStatus.COMPLETED
    ]
    due_variance = [
        (item.final_completion_date - item.task.due_date).days
        for item in completed
        if item.final_completion_date is not None and item.task.due_date is not None
    ]
    late_completion = [value for value in due_variance if value > 0]
    open_overdue = [
        (period_end - item.task.due_date).days
        for item in snapshots
        if item.counted_in_kpis
        and item.status_at_period_end.is_open
        and item.task.due_date is not None
        and item.task.due_date < period_end
    ]
    return [
        {"Metric": "Avg Execution Time (Completed)", "Value": _average_text([core.average_execution_duration_days])},
        {"Metric": "Avg Lead Time (Completed)", "Value": _average_text([core.average_lead_time_days])},
        {"Metric": "Avg Time to Start", "Value": _average_text([core.average_time_to_start_days])},
        {"Metric": "Avg Late Completion", "Value": _average_text(late_completion)},
        {"Metric": "Avg Open Overdue", "Value": _average_text(open_overdue)},
        {"Metric": "Avg Due Variance (Completed)", "Value": _average_text(due_variance)},
    ]


def _project_summary_frame(snapshots: Any) -> pd.DataFrame:
    grouped: dict[str, list[Any]] = defaultdict(list)
    for item in snapshots:
        if item.counted_in_kpis:
            grouped[item.task.unified_project or "Unmapped Project"].append(item)

    rows = []
    for project, items in sorted(grouped.items(), key=lambda pair: pair[0].casefold()):
        known = [item for item in items if item.status_at_period_end is not UnifiedStatus.UNKNOWN]
        completed = [item for item in known if item.status_at_period_end is UnifiedStatus.COMPLETED]
        completed_with_due = [
            item for item in completed
            if item.task.due_date is not None and item.final_completion_date is not None
        ]
        on_time = [
            item for item in completed_with_due
            if item.final_completion_date <= item.task.due_date
        ]
        overdue = [
            item for item in known
            if item.status_at_period_end.is_open
            and item.task.due_date is not None
            and item.task.due_date < item.period_end
        ]
        total = len(known)
        rows.append({
            "Project": project,
            "Total Tasks": total,
            "Completed": len(completed),
            "Completion Rate": round(len(completed) / total * 100, 1) if total else None,
            "On-Time Rate": round(len(on_time) / len(completed_with_due) * 100, 1) if completed_with_due else None,
            "Open Overdue": len(overdue),
        })
    return pd.DataFrame(rows, columns=[
        "Project", "Total Tasks", "Completed", "Completion Rate", "On-Time Rate", "Open Overdue"
    ])


def _weekly_company_flow_frame(snapshots: Any) -> pd.DataFrame:
    values: dict[date, dict[str, int]] = defaultdict(lambda: {
        "Tasks Created": 0,
        "Tasks Completed": 0,
    })
    for item in snapshots:
        if not item.counted_in_kpis:
            continue
        if item.task.created_date:
            week = item.task.created_date - timedelta(days=item.task.created_date.weekday())
            values[week]["Tasks Created"] += 1
        if item.final_completion_date:
            week = item.final_completion_date - timedelta(days=item.final_completion_date.weekday())
            values[week]["Tasks Completed"] += 1
    rows = [{"Week Starting": week, **counts} for week, counts in sorted(values.items())]
    return pd.DataFrame(rows, columns=["Week Starting", "Tasks Created", "Tasks Completed"])


def render_company_result(
    st: Any,
    result: CompanyAnalysisResult,
    *,
    scope_title: str = "Company Performance Analysis",
    close_state_key: str = "company_analysis",
    download_stem: str = "company_performance",
) -> None:
    """Render a concise company dashboard plus detailed drill-down outputs."""
    st.divider()
    heading, action = st.columns([5, 1])
    heading.title(scope_title)
    if action.button("Close Analysis", key=f"close_{close_state_key}", use_container_width=True):
        st.session_state.pop(close_state_key, None)
        st.rerun()

    with st.expander("Optional dashboard filters", expanded=False):
        filters = _filters(st, result)
    model = build_company_dashboard(
        result.snapshots,
        coverages=result.collection.coverages,
        filters=filters,
    )
    selected_keys = {(detail.source_tool, detail.task_id) for detail in model.task_details}
    selected = [
        snapshot for snapshot in result.snapshots
        if (snapshot.task.source_tool, snapshot.task.task_id) in selected_keys
    ]
    bottlenecks = identify_bottleneck_candidates(selected)
    recommendations = generate_recommendations(selected)

    dashboard, workflow, details, quality = st.tabs([
        "Executive Dashboard", "Workflow & Recommendations", "Task Details", "Coverage & Data Quality"
    ])
    with dashboard:
        st.caption(
            f"Analysis period: {model.period_start.isoformat()} "
            f"to {model.period_end.isoformat()}"
        )
        st.caption(
            f"Showing {model.in_period_task_count} task(s) within the selected analysis period. "
            "Tasks outside From/To are excluded from the dashboard and analysis outputs."
        )
        kpis = model.kpis
        project_count = len({
            item.task.unified_project or "Unmapped Project"
            for item in selected
            if item.counted_in_kpis
        })
        card_values = {
            "total-projects": ("Total Projects", str(project_count), "Distinct inferred project names in the selected scope."),
            "total-tasks": ("Total Tasks", str(kpis.total_tasks), "Count of distinct tasks in the selected period."),
            "completed-tasks": ("Completed Tasks", str(kpis.completed_tasks), "Tasks whose status is Completed at period end."),
            "completion-rate": ("Completion Rate", "N/A" if kpis.completion_rate is None else f"{kpis.completion_rate:.1f}%", "Completed tasks / tasks with a verified status at period end × 100. Unknown statuses are excluded and shown in Data Quality."),
            "on-time-rate": ("On-Time Rate", "N/A" if kpis.on_time_completion_rate is None else f"{kpis.on_time_completion_rate:.1f}%", "Completed tasks on or before due date / completed tasks with valid dates × 100."),
            "overdue-open": ("Open Overdue", str(kpis.overdue_open_tasks), "Open tasks with a valid due date before period end."),
        }
        population = population_from_snapshots(selected)
        cards = st.columns(3)
        label, value, help_text = card_values["total-tasks"]
        cards[0].metric(label, value, help=help_text)
        render_population_card(cards[1], population)
        label, value, help_text = card_values["completion-rate"]
        cards[2].metric(label, value, help=help_text)
        cards = st.columns(4)
        for column, key in zip(
            cards,
            ("total-projects", "completed-tasks", "on-time-rate", "overdue-open"),
        ):
            label, value, help_text = card_values[key]
            column.metric(label, value, help=help_text)

        st.subheader("Management Averages")
        st.caption("Values are elapsed calendar days calculated from the same selected task snapshots.")
        average_cards = st.columns(6)
        for column, row in zip(average_cards, _management_average_rows(selected, model.period_end)):
            column.metric(row["Metric"], row["Value"], help="Hover for the metric definition and calculation basis.")

        st.subheader("Project Performance Comparison")
        summary = _project_summary_frame(selected)
        if summary.empty:
            st.info("No project-level data is available for the selected scope.")
        else:
            st.bar_chart(
                summary.set_index("Project")[["Completion Rate", "On-Time Rate"]].fillna(0),
                use_container_width=True,
            )

        st.subheader("Task Status by Project")
        status_rows: dict[tuple[str, str], int] = defaultdict(int)
        for item in selected:
            if item.counted_in_kpis:
                status_rows[(item.task.unified_project or "Unmapped Project", item.status_at_period_end.value)] += 1
        status_frame = pd.DataFrame([
            {"Project": project, "Status": status, "Tasks": count}
            for (project, status), count in sorted(status_rows.items())
        ])
        if status_frame.empty:
            st.info("No status data is available for the selected scope.")
        else:
            st.bar_chart(
                status_frame.pivot(index="Project", columns="Status", values="Tasks").fillna(0),
                use_container_width=True,
            )

        st.subheader("Company Weekly Task Flow")
        flow = _weekly_company_flow_frame(selected)
        if flow.empty:
            st.info("No created/completed dates are available for the selected scope.")
        else:
            st.line_chart(flow.set_index("Week Starting"), use_container_width=True)

        st.subheader("Project Summary")
        st.dataframe(summary, hide_index=True, use_container_width=True)

    with workflow:
        st.subheader("Bottleneck Candidates")
        st.caption("Candidates indicate follow-up evidence, not a confirmed root cause.")
        st.dataframe(_table_rows(bottlenecks), hide_index=True, use_container_width=True)
        st.subheader("Recommendations")
        st.dataframe(_table_rows(recommendations), hide_index=True, use_container_width=True)

    with details:
        st.dataframe(_task_detail_rows(model.task_details), hide_index=True, use_container_width=True)

    with quality:
        st.subheader("Source Coverage")
        st.dataframe(_table_rows(model.source_coverage), hide_index=True, use_container_width=True)
        st.subheader("Data Quality")
        st.dataframe(_table_rows(model.data_quality), hide_index=True, use_container_width=True)
        st.subheader("Task-level Data Quality Details")
        quality_rows = data_quality_task_rows(model.task_details)
        if quality_rows:
            st.dataframe(quality_rows, hide_index=True, use_container_width=True)
        else:
            st.success("No task-level Unknown status or quality reasons were recorded.")

    excel, raw, word = _cached_output_bytes(st, result, model, cache_prefix=download_stem)
    st.subheader("Downloads")
    left, middle, right = st.columns(3)
    left.download_button(
        "Download Excel Report",
        excel,
        f"{download_stem}_analysis.xlsx",
        _MIME_XLSX,
        use_container_width=True,
    )
    middle.download_button(
        "Download Raw Collected Data",
        raw,
        f"{download_stem}_raw_data.xlsx",
        _MIME_XLSX,
        use_container_width=True,
    )
    right.download_button(
        "Download Word Report",
        word,
        f"{download_stem}_report.docx",
        _MIME_DOCX,
        use_container_width=True,
    )
