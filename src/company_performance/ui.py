"""Additive Streamlit screens for Company Performance.

The functions receive Streamlit as an argument so the company domain remains
testable without a running Streamlit app.  They never invoke Jira or ClickUp
APIs; source collection remains in the established collector screens.
"""

from __future__ import annotations

import tempfile
from datetime import date
from pathlib import Path
from typing import Any

from .application import (
    CompanyAnalysisResult,
    build_company_analysis,
    build_company_preview,
    filter_company_preview,
)
from .dashboard import DashboardFilters, build_company_dashboard
from .kpis import generate_recommendations, identify_bottleneck_candidates
from .models import UnifiedStatus
from .outputs import write_company_excel, write_company_raw_data, write_company_word_report


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
    """Render independent source selectors, combined preview, and run controls."""
    jira_spaces = available_prepared_spaces(st.session_state, "Jira")
    clickup_spaces = available_prepared_spaces(st.session_state, "ClickUp")

    with st.expander("Company Performance — Selected Projects", expanded=bool(jira_spaces or clickup_spaces)):
        st.caption(
            "Select any collected Jira and/or ClickUp spaces. "
            "Existing source-specific analyses remain separate."
        )
        if not jira_spaces and not clickup_spaces:
            st.info("No source data is ready yet. Collect at least one Jira or ClickUp space first.")
            return

        selectors = st.columns(2)
        selected_jira_names = selectors[0].multiselect(
            "Select Jira Spaces",
            list(jira_spaces),
            default=list(jira_spaces),
            key="company_selected_jira_spaces",
        )
        selected_clickup_names = selectors[1].multiselect(
            "Select ClickUp Spaces",
            list(clickup_spaces),
            default=list(clickup_spaces),
            key="company_selected_clickup_spaces",
        )
        selected_jira = tuple(jira_spaces[name] for name in selected_jira_names)
        selected_clickup = tuple(clickup_spaces[name] for name in selected_clickup_names)

        if not (selected_jira or selected_clickup):
            st.info("Select at least one Jira or ClickUp space to preview tasks and run Company Performance.")
            return

        preview = build_company_preview(
            jira_prepared=selected_jira,
            clickup_prepared=selected_clickup,
        )
        st.subheader("Combined Task Preview")
        filtered_preview = _preview_filters(st, preview)
        st.dataframe(
            _preview_table_rows(filtered_preview),
            hide_index=True,
            use_container_width=True,
        )

        # The approved run order starts only after source selection and preview.
        unified_project = st.text_input("Unified Project Name", key="company_unified_project")
        default_start, default_end = _period_defaults((*selected_jira, *selected_clickup))
        period_columns = st.columns(2)
        period_start = period_columns[0].date_input(
            "Analysis Period Start",
            value=default_start,
            key="company_period_start",
        )
        period_end = period_columns[1].date_input(
            "Analysis Period End",
            value=default_end,
            key="company_period_end",
        )
        if st.button(
            "Run Analysis",
            type="primary",
            key="run_company_performance",
            disabled=not unified_project.strip(),
        ):
            try:
                st.session_state["company_analysis"] = build_company_analysis(
                    period_start=period_start,
                    period_end=period_end,
                    jira_prepared=selected_jira,
                    clickup_prepared=selected_clickup,
                    unified_project=unified_project,
                )
                st.success("Company Performance analysis completed.")
            except ValueError as exc:
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
        bottlenecks = identify_bottleneck_candidates(result.snapshots)
        recommendations = generate_recommendations(result.snapshots)
        excel = write_company_excel(
            model,
            root / "company_performance_analysis.xlsx",
            bottlenecks=bottlenecks,
            recommendations=recommendations,
        ).read_bytes()
        raw = write_company_raw_data(
            result.snapshots,
            root / "company_performance_raw_data.xlsx",
        ).read_bytes()
        word = write_company_word_report(
            model,
            root / "company_performance_report.docx",
            bottlenecks=bottlenecks,
            recommendations=recommendations,
        ).read_bytes()
    return excel, raw, word


def render_company_result(st: Any, result: CompanyAnalysisResult) -> None:
    """Render company results, drill-downs, and the separate audited outputs."""
    st.divider()
    heading, action = st.columns([5, 1])
    heading.title("Company Performance")
    if action.button("Close Company View", key="close_company_view", use_container_width=True):
        st.session_state.pop("company_analysis", None)
        st.rerun()

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
        st.caption(f"Analysis period: {model.period_start.isoformat()} to {model.period_end.isoformat()}")
        cards = st.columns(5)
        for column, card in zip(cards, model.cards):
            column.metric(card.title, card.value, help=card.supporting_text)
        for chart in model.executive_charts:
            st.subheader(chart.title)
            if chart.points:
                st.bar_chart({point.label: point.value for point in chart.points}, use_container_width=True)
            else:
                st.info("No eligible data for this chart.")
            if chart.note:
                st.caption(chart.note)

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

    excel, raw, word = _output_bytes(result, model)
    st.subheader("Downloads")
    left, middle, right = st.columns(3)
    left.download_button(
        "Download Company Excel",
        excel,
        "company_performance_analysis.xlsx",
        _MIME_XLSX,
        use_container_width=True,
    )
    middle.download_button(
        "Download Raw Collected Data",
        raw,
        "company_performance_raw_data.xlsx",
        _MIME_XLSX,
        use_container_width=True,
    )
    right.download_button(
        "Download Company Word Report",
        word,
        "company_performance_report.docx",
        _MIME_DOCX,
        use_container_width=True,
    )
