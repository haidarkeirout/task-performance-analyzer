"""Additive Streamlit screens for Company Performance.

The functions receive Streamlit as an argument so the company domain remains
testable without a running Streamlit app.  They never invoke Jira or ClickUp
APIs; source collection remains in the established collector screens.
"""

from __future__ import annotations

import tempfile
from datetime import date, timedelta
from pathlib import Path
from typing import Any

from .application import CompanyAnalysisResult, build_company_analysis
from .dashboard import DashboardFilters, build_company_dashboard
from .kpis import generate_recommendations, identify_bottleneck_candidates
from .models import UnifiedStatus
from .outputs import write_company_excel, write_company_raw_data, write_company_word_report


_MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
_MIME_DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"


def remember_prepared_source(session_state: Any, source: str, prepared: Any | None) -> None:
    """Keep one prepared result per source while the user collects the other.

    This cache stores the same in-memory collector result already used by the
    source-specific path.  It does not trigger collection or overwrite either
    Jira/ClickUp result.
    """
    if prepared is None:
        return
    key = "company_prepared_jira" if source == "Jira" else "company_prepared_clickup"
    session_state[key] = prepared


def _period_defaults(session_state: Any) -> tuple[date, date]:
    prepared = [session_state.get("company_prepared_jira"), session_state.get("company_prepared_clickup")]
    dates = []
    for item in prepared:
        value = getattr(item, "cutoff", None)
        try:
            parsed = date.fromisoformat(str(value)[:10])
        except (TypeError, ValueError):
            continue
        dates.append(parsed)
    end = max(dates) if dates else date.today()
    return end.replace(day=1), end


def render_company_launcher(st: Any) -> None:
    """Render source mapping and a Company Performance run button."""
    jira = st.session_state.get("company_prepared_jira")
    clickup = st.session_state.get("company_prepared_clickup")
    with st.expander("Company Performance — Selected Projects", expanded=bool(jira or clickup)):
        st.caption(
            "Collect the Jira and/or ClickUp spaces you want to include, then map each collected space "
            "to a Unified Project name. Existing source analyses remain separate."
        )
        if not jira and not clickup:
            st.info("No source data is ready yet. Select a source above, choose a space, and click Done first.")
            return

        use_jira = bool(jira)
        use_clickup = bool(clickup)
        jira_project = ""
        clickup_project = ""
        if jira:
            include_jira = st.checkbox(
                f"Include Jira: {getattr(jira, 'space_name', 'Selected space')}", value=True, key="company_include_jira"
            )
            use_jira = include_jira
            if include_jira:
                jira_project = st.text_input(
                    "Unified Project name for Jira", value=getattr(jira, "space_name", ""), key="company_jira_project"
                )
        if clickup:
            include_clickup = st.checkbox(
                f"Include ClickUp: {getattr(clickup, 'space_name', 'Selected space')}", value=True, key="company_include_clickup"
            )
            use_clickup = include_clickup
            if include_clickup:
                clickup_project = st.text_input(
                    "Unified Project name for ClickUp", value=getattr(clickup, "space_name", ""), key="company_clickup_project"
                )
        default_start, default_end = _period_defaults(st.session_state)
        period_columns = st.columns(2)
        period_start = period_columns[0].date_input("Analysis period start", value=default_start, key="company_period_start")
        period_end = period_columns[1].date_input("Analysis period end", value=default_end, key="company_period_end")
        if st.button("Run Company Performance", type="primary", key="run_company_performance", disabled=not (use_jira or use_clickup)):
            try:
                st.session_state["company_analysis"] = build_company_analysis(
                    period_start=period_start,
                    period_end=period_end,
                    jira_prepared=jira if use_jira else None,
                    jira_project=jira_project,
                    clickup_prepared=clickup if use_clickup else None,
                    clickup_project=clickup_project,
                )
                st.success("Company Performance analysis completed.")
            except ValueError as exc:
                st.error(str(exc))


def _table_rows(items: Any) -> list[dict[str, Any]]:
    return [item.__dict__ for item in items]


def _filters(st: Any, result: CompanyAnalysisResult) -> DashboardFilters:
    options = build_company_dashboard(result.snapshots, coverages=result.collection.coverages).filter_options
    columns = st.columns(5)
    sources = columns[0].multiselect("Source", options.source_tools, key="company_filter_sources")
    projects = columns[1].multiselect("Unified Project", options.unified_projects, key="company_filter_projects")
    statuses = columns[2].multiselect("Final Status", [status.value for status in options.statuses], key="company_filter_statuses")
    assignees = columns[3].multiselect("Assignee", options.assignee_groups, key="company_filter_assignees")
    priorities = columns[4].multiselect("Priority", options.priorities, key="company_filter_priorities")
    lookup = {status.value: status for status in UnifiedStatus}
    return DashboardFilters(
        source_tools=tuple(sources), unified_projects=tuple(projects),
        statuses=tuple(lookup[value] for value in statuses), assignee_groups=tuple(assignees), priorities=tuple(priorities),
    )


def _output_bytes(result: CompanyAnalysisResult, model: Any) -> tuple[bytes, bytes, bytes]:
    with tempfile.TemporaryDirectory() as directory:
        root = Path(directory)
        bottlenecks = identify_bottleneck_candidates(result.snapshots)
        recommendations = generate_recommendations(result.snapshots)
        excel = write_company_excel(model, root / "company_performance_analysis.xlsx", bottlenecks=bottlenecks, recommendations=recommendations).read_bytes()
        raw = write_company_raw_data(result.snapshots, root / "company_performance_raw_data.xlsx").read_bytes()
        word = write_company_word_report(model, root / "company_performance_report.docx", bottlenecks=bottlenecks, recommendations=recommendations).read_bytes()
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
    model = build_company_dashboard(result.snapshots, coverages=result.collection.coverages, filters=filters)
    selected_keys = {(detail.source_tool, detail.task_id) for detail in model.task_details}
    selected = [snapshot for snapshot in result.snapshots if (snapshot.task.source_tool, snapshot.task.task_id) in selected_keys]
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
        st.dataframe(_table_rows(model.task_details), hide_index=True, use_container_width=True)
    with quality:
        st.subheader("Source Coverage")
        st.dataframe(_table_rows(model.source_coverage), hide_index=True, use_container_width=True)
        st.subheader("Data Quality")
        st.dataframe(_table_rows(model.data_quality), hide_index=True, use_container_width=True)

    excel, raw, word = _output_bytes(result, model)
    st.subheader("Downloads")
    left, middle, right = st.columns(3)
    left.download_button("Download Company Excel", excel, "company_performance_analysis.xlsx", _MIME_XLSX, use_container_width=True)
    middle.download_button("Download Raw Collected Data", raw, "company_performance_raw_data.xlsx", _MIME_XLSX, use_container_width=True)
    right.download_button("Download Company Word Report", word, "company_performance_report.docx", _MIME_DOCX, use_container_width=True)
