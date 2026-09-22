from __future__ import annotations

import json
import hashlib
import os
import sys
import tempfile
from datetime import date, datetime
from io import BytesIO
from pathlib import Path
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st


ROOT_DIR = Path(__file__).resolve().parent
SRC_DIR = ROOT_DIR / "src"

if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from jira_processor import choose_sheet, normalize_tasks
from metrics_engine import (
    aggregate,
    analyze_tasks,
    load_calendar,
    normalize_labels,
    parse_timestamp,
)
from report_builder import build_report
from process_analysis import workbook_context, workbook_histories, process_tables, excel_bytes
from clickup_analysis import analyze_clickup, analysis_excel, recalculate_clickup_analysis
from employee_filters import (
    ALL as EMPLOYEE_FILTER_ALL,
    apply_employee_filters,
    employee_filter_options,
    with_employee_filter_dimensions,
)
from clickup_report import create_clickup_word_report
from company_performance.dashboard import metric_help
from kpi_transparency import build_population, render_population_card
from company_performance.collection import (
    ALL_COMPANIES_ID,
    company_option_label,
    render_company_source_mapping,
    discover_company_catalog,
)
from company_performance.ui import (
    remember_prepared_source,
    render_company_launcher,
    render_company_result,
)
from company_performance.application import build_company_analysis
from employee_ui import EmployeePreparedBundle
from project_ui import render_project_collection, render_project_result
from department_analysis import (
    build_department_result,
    build_jira_department_result,
    department_excel,
    department_word,
    filter_jira_department_period,
)


def _frame_signature(frame: pd.DataFrame) -> str:
    if frame is None or frame.empty:
        return "empty"
    hashed = pd.util.hash_pandas_object(frame.astype(str), index=True).values.tobytes()
    return hashlib.sha256(hashed).hexdigest()


CONFIG_DIR = ROOT_DIR / "configs"
INPUT_CONFIG_PATH = CONFIG_DIR / "jira_input_config.json"
METRICS_CONFIG_PATH = CONFIG_DIR / "jira_metrics_config.json"


APP_TITLE = "Performance Management"


st.set_page_config(
    page_title=APP_TITLE,
    page_icon="📊",
    layout="wide",
)


def load_configurations() -> tuple[dict, dict]:
    with INPUT_CONFIG_PATH.open(
        "r",
        encoding="utf-8",
    ) as file:
        input_config = json.load(file)

    with METRICS_CONFIG_PATH.open(
        "r",
        encoding="utf-8",
    ) as file:
        metrics_config = json.load(file)

    return input_config, metrics_config


def read_workbook(
    uploaded_file,
) -> dict[str, pd.DataFrame]:
    return pd.read_excel(
        BytesIO(uploaded_file.getvalue()),
        sheet_name=None,
        dtype=object,
    )


def read_history_file(
    uploaded_file,
) -> dict:
    return json.loads(
        uploaded_file.getvalue().decode("utf-8")
    )


def create_excel_download(task_metrics, overall, by_assignee, by_issue_type, tables=None):
    return excel_bytes(task_metrics, tables or process_tables(task_metrics))


def create_word_report_download(
    task_metrics: pd.DataFrame,
    cutoff: str,
    tables=None,
    title: str = "Employee Task Execution Performance Evaluation",
) -> bytes:
    temporary_file = tempfile.NamedTemporaryFile(
        suffix=".docx",
        delete=False,
    )

    temporary_path = Path(
        temporary_file.name
    )

    temporary_file.close()

    try:
        build_report(
            task_metrics,
            temporary_path,
            title=title,
            process_data=tables,
            evaluation_cutoff=cutoff,
            timezone_name="Asia/Damascus",
            work_days="Sunday-Thursday",
            work_window="09:00-17:00",
        )

        return temporary_path.read_bytes()

    finally:
        if temporary_path.exists():
            temporary_path.unlink()


def format_number(value) -> str:
    if value is None:
        return "Unavailable"

    try:
        if pd.isna(value):
            return "Unavailable"
    except (TypeError, ValueError):
        pass

    if isinstance(value, float):
        return f"{value:.2f}"

    return str(value)


def _streamlit_safe_frame(frame: pd.DataFrame) -> pd.DataFrame:
    """Make mixed object columns safe for Streamlit/Arrow display only."""
    if not isinstance(frame, pd.DataFrame):
        return frame
    safe = frame.copy()
    for column in safe.columns:
        if safe[column].dtype == "object":
            def display_value(value):
                if value is None:
                    return ""
                try:
                    if pd.isna(value):
                        return ""
                except (TypeError, ValueError):
                    pass
                return str(value)

            safe[column] = safe[column].map(display_value)
    return safe


def format_percentage(
    numerator: int,
    denominator: int,
) -> str:
    if denominator == 0:
        return "N/A"

    return f"{numerator / denominator * 100:.1f}%"


def _filename_component(value) -> str:
    """Return a non-empty, filesystem-safe component for report filenames."""
    text = "Unavailable" if value is None else str(value).strip()
    for character in '<>:"/\\|?*':
        text = text.replace(character, "-")
    text = " ".join(text.split()).strip(" .-")
    return text or "Unavailable"


def calculate_dashboard_values(task_metrics):
    row = aggregate(task_metrics, []).iloc[0]
    return {"total": row["total_tasks"], "completed": row["completed_tasks"],
            "rejected": row["rejected_tasks"], "open": row["open_tasks"], "wip": row["wip_tasks"],
            "completion_rate": format_percentage(int(row["completed_tasks"]), int(row["known_status_tasks"])),
            "on_time_rate": format_percentage(int(row["on_time_tasks"]), int(row["on_time_valid_tasks"])),
            "on_time_tasks": row["on_time_tasks"],
            "on_time_valid_tasks": row["on_time_valid_tasks"],
            "known_status_tasks": row["known_status_tasks"],
            "unknown_status_tasks": row["unknown_status_tasks"],
            "overdue": row["overdue_open_tasks"], "rework": row["tasks_with_rework"]}


def _employee_kpi_help(task_metrics: pd.DataFrame, values: dict) -> dict[str, str]:
    """Build the same calculation tooltip format used by Project Analysis."""
    total = int(values["total"])
    completed = int(values["completed"])
    known = int(values.get("known_status_tasks", total))
    on_time = int(values["on_time_tasks"])
    on_time_valid = int(values["on_time_valid_tasks"])
    unknown = int(
        task_metrics["status_known"].eq(False).sum()
    ) if "status_known" in task_metrics.columns else 0
    validation = (
        f"WARNING — {unknown} task(s) have unavailable period-end status"
        if unknown else "PASS"
    )
    scope = "Selected Employee scope"
    period = str(task_metrics.attrs.get("cutoff", "Selected analysis period"))
    return {
        "total": metric_help(
            formula="Count distinct Source Tool + Task ID",
            calculation=f"{total} distinct task(s)",
            scope=scope,
            period=period,
            exclusions="Duplicate task keys are removed before KPI calculation.",
            validation=validation,
        ),
        "completed": metric_help(
            formula="Count of tasks completed by the evaluation cutoff",
            calculation=f"{completed} task(s)",
            scope=scope,
            period=period,
            exclusions="Tasks without a verified completed status remain open or unavailable.",
            validation=validation,
        ),
        "completion_rate": metric_help(
            formula="Completed tasks / Known-status KPI tasks × 100",
            calculation=f"{completed} / {known} × 100 = {values['completion_rate']}",
            scope=scope,
            period=period,
            exclusions=f"Unknown-status tasks are excluded from the denominator ({values.get('unknown_status_tasks', 0)} task(s)).",
            validation=validation,
        ),
        "on_time_rate": metric_help(
            formula="Completed on or before due date / Completed tasks with valid due date × 100",
            calculation=f"{on_time} / {on_time_valid} × 100 = {values['on_time_rate']}",
            scope=scope,
            period=period,
            exclusions="Completed tasks missing a valid due date are excluded.",
            validation=validation,
        ),
        "overdue": metric_help(
            formula="Count of open tasks with a valid due date before the evaluation cutoff",
            calculation=f"{int(values['overdue'])} task(s)",
            scope=scope,
            period=period,
            exclusions="Tasks without a valid due date are excluded from overdue classification.",
            validation=validation,
        ),
        "wip": metric_help(
            formula="Count of tasks currently in a work-in-progress status",
            calculation=f"{int(values['wip'])} task(s)",
            scope=scope,
            period=period,
            exclusions="Tasks with unavailable status are not classified as WIP.",
            validation=validation,
        ),
    }


def _mean_for_mask(frame: pd.DataFrame, column: str, mask: pd.Series):
    if column not in frame.columns:
        return None
    values = pd.to_numeric(frame.loc[mask, column], errors="coerce").dropna()
    return None if values.empty else float(values.mean())


def _format_average(value, suffix: str) -> str:
    if value is None or pd.isna(value):
        return "Unavailable"
    return f"{value:.1f} {suffix}"


def show_management_averages(task_metrics: pd.DataFrame) -> None:
    """Show leadership-level averages with explicit valid-population scopes."""
    completed = task_metrics["is_completed"].eq(True)
    completed_late = completed & task_metrics["schedule_variance_days"].gt(0)
    open_overdue = task_metrics["is_open"].eq(True) & task_metrics["overdue_days"].gt(0)
    started = task_metrics["time_to_start_business_hours"].notna()
    start_variance = task_metrics["start_schedule_variance_days"].notna()

    st.subheader("Management Averages")
    st.caption(
        "Execution, lead-time, and start-wait averages use configured business hours. "
        "Delay and schedule-variance averages use calendar days. Unavailable means no qualifying tasks exist."
    )

    completed_cards = st.columns(3)
    completed_cards[0].metric(
        "Avg Execution Time (Completed)",
        _format_average(_mean_for_mask(task_metrics, "execution_business_hours", completed), "h"),
    )
    completed_cards[1].metric(
        "Avg Lead Time (Completed)",
        _format_average(_mean_for_mask(task_metrics, "lead_time_business_hours", completed), "h"),
    )
    completed_cards[2].metric(
        "Avg Time to Start",
        _format_average(_mean_for_mask(task_metrics, "time_to_start_business_hours", started), "h"),
    )

    delay_cards = st.columns(3)
    delay_cards[0].metric(
        "Avg Late Completion",
        _format_average(_mean_for_mask(task_metrics, "schedule_variance_days", completed_late), "days"),
    )
    delay_cards[1].metric(
        "Avg Open Overdue",
        _format_average(_mean_for_mask(task_metrics, "overdue_days", open_overdue), "days"),
    )
    delay_cards[2].metric(
        "Avg Start Variance",
        _format_average(_mean_for_mask(task_metrics, "start_schedule_variance_days", start_variance), "days"),
    )


def show_weekly_task_flow(process_data: dict | None = None) -> None:
    """Show weekly task intake versus completion without a chart dependency."""
    weekly = process_data.get("weekly_flow") if process_data else None
    st.subheader("Weekly Task Flow")
    st.caption(
        "Tasks opened are grouped by Created Date; tasks completed are grouped by the date they reached Done. "
        "Each point represents a Monday-starting week."
    )
    if weekly is None or weekly.empty:
        st.info("No created or completed dates are available for a weekly trend.")
        return

    range_options = {
        "All available weeks": None,
        "Last 4 weeks": 4,
        "Last 12 weeks": 12,
        "Last 24 weeks": 24,
        "Last 36 weeks": 36,
        "Last 48 weeks": 48,
    }
    selected_range = st.selectbox(
        "Trend period",
        list(range_options),
        key="weekly_flow_period",
    )
    weeks_to_show = range_options[selected_range]
    visible = weekly.tail(weeks_to_show).copy() if weeks_to_show else weekly.copy()
    chart = visible.set_index("week_start")[["tasks_opened", "tasks_completed"]].rename(
        columns={
            "tasks_opened": "Tasks Opened",
            "tasks_completed": "Tasks Completed",
        }
    )
    st.line_chart(chart, use_container_width=True)
    st.dataframe(
        visible.rename(columns={
            "week_start": "Week Starting",
            "tasks_opened": "Tasks Opened",
            "tasks_completed": "Tasks Completed",
            "net_flow": "Net Flow",
            "cumulative_net_flow": "Cumulative Net Flow",
        }),
        use_container_width=True,
        hide_index=True,
    )


def run_analysis(
    uploaded_excel,
    uploaded_history,
    base_url: str,
    email: str,
    token: str,
    cutoff_text: str,
    source_timezone: str,
    history_source="Jira API",
    coverage_through=None,
    coverage_confirmed=False,
    context_overrides=None,
    period_start=None,
):
    input_config, metrics_config = load_configurations()

    cutoff = parse_timestamp(
        cutoff_text,
        source_timezone,
    )

    if cutoff is None:
        raise ValueError(
            "The evaluation cutoff is invalid."
        )

    calendar = load_calendar(
        metrics_config,
        extra_holidays=[],
    )

    workbook = read_workbook(
        uploaded_excel,
    )

    sheet_name, source_frame = choose_sheet(
        workbook,
        input_config,
        requested_sheet=None,
    )

    tasks_frame, validation_log = normalize_tasks(
        source_frame,
        input_config=input_config,
        source_timezone=ZoneInfo(source_timezone),
        labels_delimiter=None,
        explicit_date_format=None,
    )

    if history_source == "Workbook transitions":
        histories = workbook_histories(workbook, tasks_frame, source_timezone,
                                       coverage_through or cutoff_text, coverage_confirmed)
    elif history_source == "History JSON":
        if uploaded_history is None:
            raise ValueError("Upload the history JSON file.")
        histories = read_history_file(
            uploaded_history,
        )
    else:
        if not base_url or not email or not token:
            raise ValueError(
                "Enter Jira URL, email, and API token, "
                "or upload a history JSON file."
            )

        from jira_client import JiraClient
        client = JiraClient(
            base_url=base_url.strip(),
            email=email.strip(),
            api_token=token.strip(),
        )

        client.verify_connection()

        issue_keys = (
            tasks_frame["issue_key"]
            .dropna()
            .astype(str)
            .tolist()
        )

        histories = client.fetch_many_issue_histories(
            issue_keys,
        )

    task_metrics = analyze_tasks(
        tasks_frame,
        histories,
        cutoff=cutoff,
        calendar=calendar,
    )

    if period_start is not None:
        task_metrics = filter_jira_department_period(
            task_metrics,
            histories,
            period_start,
            cutoff,
            source_timezone,
        )

    if task_metrics.empty:
        raise ValueError("No tasks were active during the selected analysis period.")
    context = workbook_context(workbook)
    context.update(context_overrides or {})
    context["History source"] = history_source
    task_metrics.attrs["process_context"] = context
    task_metrics.attrs["selected_sheet"] = sheet_name
    task_metrics.attrs["validation_log"] = validation_log
    task_metrics.attrs["cutoff"] = cutoff.isoformat()
    task_metrics.attrs["histories"] = histories

    tables = process_tables(task_metrics, histories, context, validation_log)
    return task_metrics, validation_log, tables


def _reset_employee_project_filter() -> None:
    st.session_state["employee_filter_projects"] = []


def _clear_employee_post_analysis_filters() -> None:
    st.session_state["employee_filter_company"] = EMPLOYEE_FILTER_ALL
    st.session_state["employee_filter_projects"] = []
    for key in (
        "employee_filter_status",
        "employee_filter_task_type",
        "employee_filter_priority",
        "employee_filter_due_state",
    ):
        st.session_state[key] = []


def render_employee_post_analysis_filters(frame: pd.DataFrame, source: str) -> pd.DataFrame:
    """Render Employee-only filters and return the locally filtered snapshot."""
    dimensions = with_employee_filter_dimensions(frame, source)
    options = employee_filter_options(dimensions)
    st.subheader("Employee analysis filters")
    st.caption(
        "These filters use the collected snapshot only. Changing them does not recollect data from Jira or ClickUp."
    )
    top = st.columns(2)
    company = top[0].selectbox(
        "Company",
        options["Company"],
        key="employee_filter_company",
        on_change=_reset_employee_project_filter,
    )
    company_options = employee_filter_options(dimensions, company)
    projects = top[1].multiselect(
        "Projects / Spaces",
        company_options["Project / Space"],
        key="employee_filter_projects",
        help="Only Projects / Spaces belonging to the selected Company are shown.",
    )
    bottom = st.columns(4)
    status = bottom[0].multiselect("Status", options["Status"], key="employee_filter_status")
    task_type_value = bottom[1].multiselect("Task Type", options["Task Type"], key="employee_filter_task_type")
    priority = bottom[2].multiselect("Priority", options["Priority"], key="employee_filter_priority")
    due_state = bottom[3].multiselect("Due-Date State", options["Due-Date State"], key="employee_filter_due_state")
    st.button("Clear Employee Filters", on_click=_clear_employee_post_analysis_filters, key="clear_employee_post_analysis_filters")

    filtered = apply_employee_filters(
        dimensions,
        {
            "Company": company,
            "Project / Space": projects,
            "Status": status,
            "Task Type": task_type_value,
            "Priority": priority,
            "Due-Date State": due_state,
        },
    )
    st.caption(f"Showing {len(filtered)} of {len(frame)} analysed task(s).")
    return filtered


def show_executive_dashboard(
    task_metrics: pd.DataFrame,
    process_data: dict | None = None,
) -> None:
    values = calculate_dashboard_values(
        task_metrics,
    )

    help_text = _employee_kpi_help(task_metrics, values)
    population = build_population(
        values["total"],
        values.get("known_status_tasks", values["total"]),
        {"Unknown status/history": values.get("unknown_status_tasks", 0)},
    )
    columns = st.columns(3)

    columns[0].metric(
        "Total Tasks",
        values["total"],
        help=help_text["total"],
    )
    render_population_card(columns[1], population)
    columns[2].metric(
        "Completion Rate",
        values["completion_rate"],
        help=help_text["completion_rate"],
    )

    columns = st.columns(4)
    columns[0].metric(
        "Completed",
        values["completed"],
        help=help_text["completed"],
    )

    columns[1].metric(
        "On-Time Rate",
        values["on_time_rate"],
        help=help_text["on_time_rate"],
    )
    columns[2].metric(
        "Open Overdue",
        values["overdue"],
        help=help_text["overdue"],
    )
    columns[3].metric(
        "WIP Tasks",
        values["wip"],
        help=help_text["wip"],
    )

    st.divider()

    show_management_averages(task_metrics)

    st.divider()

    show_weekly_task_flow(process_data)

    st.divider()

    left, right = st.columns(2)

    with left:
        status_counts = (
            task_metrics["status_at_cutoff"]
            .fillna("Unavailable")
            .value_counts()
            .rename_axis("Status")
            .reset_index(name="Tasks")
        )

        if not status_counts.empty:
            st.subheader("Task Distribution by Status")
            st.bar_chart(
                status_counts.set_index("Status")["Tasks"],
                use_container_width=True,
            )

        st.subheader("Open Tasks by Due Status")
        st.caption(
            "Open tasks with verified status, grouped by the supplied due date. "
            "Unknown-status tasks remain separate in Data Quality."
        )
        deadline_table = process_data.get("deadline_summary") if process_data else None
        if deadline_table is not None and not deadline_table.empty:
            st.bar_chart(
                deadline_table.set_index("due_status")["task_count"],
                use_container_width=True,
            )
            st.dataframe(
                deadline_table.rename(columns={
                    "due_status": "Due status",
                    "task_count": "Tasks",
                    "share_of_open_known_tasks": "Share of open known tasks (%)",
                }),
                use_container_width=True,
                hide_index=True,
            )
        unknown_status = int(task_metrics["status_known"].eq(False).sum()) if "status_known" in task_metrics.columns else 0
        late_completed = int((
            task_metrics["is_completed"].eq(True)
            & task_metrics["schedule_variance_days"].notna()
            & task_metrics["schedule_variance_days"].gt(0)
        ).sum()) if "schedule_variance_days" in task_metrics.columns else 0
        card_a, card_b = st.columns(2)
        card_a.metric("Completed late", late_completed)
        card_b.metric("Status unavailable", unknown_status)
        if process_data and not process_data.get("late_completed_tasks", pd.DataFrame()).empty:
            st.caption("Completed late tasks")
            st.dataframe(
                process_data["late_completed_tasks"],
                use_container_width=True,
                hide_index=True,
            )

    with right:
        if "assignee_name" in task_metrics.columns:
            assignee_frame = (
                task_metrics.groupby(
                    "assignee_name",
                    dropna=False,
                )
                .agg(
                    Completed=(
                        "is_completed",
                        "sum",
                    ),
                    Open=(
                        "is_open",
                        "sum",
                    ),
                    Rejected=("is_rejected", "sum"),
                )
                .reset_index()
                .rename(
                    columns={
                        "assignee_name": "Assignee",
                    }
                )
            )

            if not assignee_frame.empty:
                st.subheader("Work Distribution by Assignee")
                st.bar_chart(
                    assignee_frame.set_index("Assignee")[["Completed", "Open", "Rejected"]],
                    use_container_width=True,
                )

    st.caption("Work distribution uses the uploaded assignee snapshot and does not establish individual contribution.")


def show_process_analysis(frame, tables):
    st.subheader("Process scope and definitions")
    st.dataframe(_streamlit_safe_frame(tables["process_context"]), hide_index=True, use_container_width=True)
    row = tables["overall_summary"].iloc[0]
    st.caption(f"Complete histories: {int(row['history_complete_tasks'])}/{int(row['total_tasks'])}; "
               f"excluded from history-dependent metrics: {int(row['history_excluded_tasks'])}; "
               f"reviewed tasks eligible for exception rates: {int(row['reviewed_valid_tasks'])}.")
    counts = st.columns(4)
    for col, label, count_name, rate_name in zip(counts,
            ["Rework", "Replanning", "Re-evaluation", "Any review exception"],
            ["tasks_with_rework", "tasks_with_replanning", "tasks_with_re_evaluation", "review_exception_tasks"],
            ["rework_rate", "replanning_rate", "re_evaluation_rate", "review_exception_rate"]):
        rate = row[rate_name]
        col.metric(label, "Unavailable" if pd.isna(rate) else f"{rate:.2f}%")
        col.caption(f"{int(row[count_name])} affected / {int(row['reviewed_valid_tasks'])} reviewed tasks")
    event_counts = {
        kind: row["total_" + kind + "_count"]
        for kind in ["rework", "replanning", "re_evaluation"]
    }
    event_labels = {
        "rework": "Rework",
        "replanning": "Replanning",
        "re_evaluation": "Re-evaluation",
    }
    event_totals = pd.DataFrame([
        {
            "Event Type": event_labels[kind],
            "Count": int(count),
        }
        for kind, count in event_counts.items()
    ])
    st.subheader("Event totals")
    st.caption("A task may have multiple events.")
    st.dataframe(event_totals, hide_index=True, use_container_width=True)
    durations = []
    for field in ["execution", "lead_time", "time_to_start"]:
        durations.append({"Measure": field.replace("_", " ").title(),
                          "Mean elapsed hours": row["mean_" + field + "_elapsed_hours"],
                          "Mean business hours": row["mean_" + field + "_business_hours"],
                          "Valid tasks": row[field + "_elapsed_hours_valid_tasks"]})
    st.dataframe(pd.DataFrame(durations), hide_index=True, use_container_width=True)
    st.subheader("Stage residence")
    st.caption("Repeated visits are summed per task. Unfinished visits are included through cutoff; terminal states are excluded. These are process durations, not labor hours.")
    st.dataframe(tables["stage_summary"], hide_index=True, use_container_width=True)
    st.subheader("Open tasks and current waiting")
    st.dataframe(tables["open_tasks"], hide_index=True, use_container_width=True)
    st.subheader("Open overdue tasks")
    if tables["overdue_tasks"].empty:
        st.success("No open overdue tasks were identified in the current scope.")
    else:
        st.dataframe(tables["overdue_tasks"], hide_index=True, use_container_width=True)
    st.subheader("Evidence and follow-up")
    st.dataframe(tables["process_findings"], hide_index=True, use_container_width=True)
    with st.expander("Transition audit trail"):
        st.dataframe(tables["workflow_events"], hide_index=True, use_container_width=True)
    with st.expander("Metric definitions"):
        st.dataframe(tables["metric_definitions"], hide_index=True, use_container_width=True)


def show_task_detail(
    task_metrics: pd.DataFrame,
) -> None:
    st.subheader("Task-Level Evaluation")

    filtered = task_metrics.copy()

    if "assignee_name" in filtered.columns:
        assignees = [
            "All",
            *sorted(
                filtered["assignee_name"]
                .fillna("Assignee unavailable")
                .astype(str)
                .unique()
                .tolist()
            ),
        ]

        selected_assignee = st.selectbox(
            "Filter by assignee",
            assignees,
            key="task_assignee_filter",
        )

        if selected_assignee != "All":
            filtered = filtered[
                filtered["assignee_name"]
                .fillna("Assignee unavailable")
                .astype(str)
                == selected_assignee
            ]

    if "issue_type" in filtered.columns:
        issue_types = [
            "All",
            *sorted(
                filtered["issue_type"]
                .fillna("Issue type unavailable")
                .astype(str)
                .unique()
                .tolist()
            ),
        ]

        selected_type = st.selectbox(
            "Filter by issue type",
            issue_types,
            key="task_type_filter",
        )

        if selected_type != "All":
            filtered = filtered[
                filtered["issue_type"]
                .fillna("Issue type unavailable")
                .astype(str)
                == selected_type
            ]

    label_filter = st.text_input(
        "Optional label filter",
        key="task_label_filter",
    ).strip()

    if label_filter and "labels" in filtered.columns:
        filtered = filtered[
            filtered["labels"].apply(
                lambda value: label_filter in normalize_labels(value)
            )
        ]

    display_columns = [
        "issue_key",
        "task_name",
        "issue_type",
        "labels",
        "assignee_name",
        "priority",
        "status_at_cutoff",
        "created_at",
        "planned_start_date",
        "actual_start_at",
        "start_schedule_variance_days",
        "completed_at",
        "due_date",
        "execution_elapsed_hours",
        "execution_business_hours",
        "lead_time_business_hours",
        "on_time_completion",
        "overdue_days",
        "rework_count",
        "replanning_count",
        "re_evaluation_count",
    ]

    display_columns = [
        column
        for column in display_columns
        if column in filtered.columns
    ]

    st.dataframe(
        filtered[display_columns],
        use_container_width=True,
        hide_index=True,
    )


def show_assignment_summary(task_metrics: pd.DataFrame) -> None:
    st.subheader("Individual Achievements and Assignment Summary")
    st.caption(
        "Assignment uses the assignee recorded in the Jira snapshot. "
        "Unassigned is a separate group; these values do not establish individual contribution."
    )
    st.dataframe(
        aggregate(task_metrics, ["assignee_name"]),
        use_container_width=True,
        hide_index=True,
    )

    assignees = sorted(task_metrics["assignee_name"].fillna("Unassigned").astype(str).unique())
    selected = st.selectbox(
        "Show tasks for",
        ["All", *assignees],
        key="assignment_task_filter",
    )
    detail = task_metrics.copy()
    if selected != "All":
        detail = detail[detail["assignee_name"].fillna("Unassigned").astype(str).eq(selected)]

    columns = [
        "issue_key", "task_name", "priority", "status_at_cutoff",
        "created_at", "planned_start_date", "actual_start_at",
        "start_schedule_variance_days", "due_date", "completed_at",
        "execution_elapsed_hours", "execution_business_hours",
        "current_status_age_elapsed_hours", "current_status_age_business_hours",
        "overdue_days", "rework_count",
    ]
    detail = detail[[column for column in columns if column in detail.columns]].rename(columns={
        "issue_key": "Issue Key", "task_name": "Task Name", "priority": "Priority",
        "status_at_cutoff": "Status", "created_at": "Created Date",
        "planned_start_date": "Planned Start Date", "actual_start_at": "Actual Start Date",
        "start_schedule_variance_days": "Start Variance (Days)", "due_date": "Due Date",
        "completed_at": "Completed Date", "execution_elapsed_hours": "Execution Elapsed Hours",
        "execution_business_hours": "Execution Business Hours",
        "current_status_age_elapsed_hours": "Current Status Age (Elapsed Hours)",
        "current_status_age_business_hours": "Current Status Age (Business Hours)",
        "overdue_days": "Overdue Days", "rework_count": "Rework Count",
    })
    st.subheader("Task details")
    st.dataframe(detail, use_container_width=True, hide_index=True)


def show_data_quality(
    task_metrics: pd.DataFrame,
    validation_log: list[dict],
) -> None:
    st.subheader("Data Quality and Coverage")

    if validation_log:
        log_frame = pd.DataFrame(
            validation_log
        )

        st.dataframe(
            log_frame,
            use_container_width=True,
            hide_index=True,
        )
    else:
        st.success(
            "No validation warnings or errors were returned."
        )

    if "history_complete" in task_metrics.columns:
        complete_count = int(
            task_metrics["history_complete"].sum()
        )

        st.write(
            (
                f"Tasks with complete usable status history: "
                f"{complete_count} of {len(task_metrics)}."
            )
        )

    st.write(
        (
            "Unavailable metrics remain unavailable when their source "
            "data is missing. They are not replaced with zero."
        )
    )


def show_clickup_analysis(result) -> None:
    """Render a ClickUp-only analytics experience that mirrors the Jira layout.

    It intentionally consumes the independent ClickUp canonical data model, so
    no Jira history, Jira files, or Jira calculations are used here.
    """
    def metric_value(name, default=None):
        matches = result["overall"].loc[result["overall"]["Metric"].eq(name), "Value"]
        return default if matches.empty else matches.iloc[0]

    def count(name):
        value = metric_value(name, 0)
        return 0 if pd.isna(value) else int(value)

    def percent(name):
        value = metric_value(name)
        return "N/A" if value is None or pd.isna(value) else f"{float(value):.1f}%"

    def average(name, suffix):
        value = metric_value(name)
        return "Unavailable" if value is None or pd.isna(value) else f"{float(value):.1f} {suffix}"

    tasks = result["tasks"].copy()
    clickup_total = len(tasks)
    clickup_known = int(tasks["Status Known?"].eq(True).sum()) if clickup_total else 0
    clickup_completed = int(tasks["Completed?"].sum()) if clickup_total else 0
    clickup_on_time = int(tasks.loc[tasks["Completed?"] & tasks["Due Variance (days)"].notna(), "On Time?"].sum()) if clickup_total else 0
    clickup_on_time_valid = int((tasks["Completed?"] & tasks["Due Variance (days)"].notna()).sum()) if clickup_total else 0
    clickup_validation = "PASS"
    if "Timing Data Status" in tasks.columns and tasks["Timing Data Status"].ne("Available").any():
        clickup_validation = "WARNING — some timing fields are unavailable"
    clickup_scope = f"ClickUp Space: {result['space_name']}"
    clickup_period = str(result["cutoff"])
    clickup_help = {
        "total": metric_help(
            formula="Count distinct ClickUp Task ID",
            calculation=f"{clickup_total} distinct task(s)",
            scope=clickup_scope, period=clickup_period,
            exclusions="Duplicate task IDs are removed before KPI calculation.",
            validation=clickup_validation,
        ),
        "completed": metric_help(
            formula="Count of tasks with Completed? = True",
            calculation=f"{clickup_completed} task(s)",
            scope=clickup_scope, period=clickup_period,
            exclusions="Tasks without a completed status are not counted as completed.",
            validation=clickup_validation,
        ),
        "completion": metric_help(
            formula="Completed tasks / Known-status tasks × 100",
            calculation=f"{clickup_completed} / {clickup_known} × 100 = {percent('Completion rate (%)')}",
            scope=clickup_scope, period=clickup_period,
            exclusions=f"Unknown-status tasks are excluded from the denominator ({clickup_total - clickup_known} task(s)).",
            validation=clickup_validation,
        ),
        "on_time": metric_help(
            formula="On-time completed tasks / Completed tasks with valid due variance × 100",
            calculation=f"{clickup_on_time} / {clickup_on_time_valid} × 100 = {percent('On-time completion rate (%)')}",
            scope=clickup_scope, period=clickup_period,
            exclusions="Completed tasks without valid due/completion dates are excluded.",
            validation=clickup_validation,
        ),
        "overdue": metric_help(
            formula="Count of open tasks with positive Due Variance",
            calculation=f"{len(result['overdue_tasks'])} task(s)",
            scope=clickup_scope, period=clickup_period,
            exclusions="Tasks without a valid due date are excluded.",
            validation=clickup_validation,
        ),
        "wip": metric_help(
            formula="Count of tasks with WIP? = True",
            calculation=f"{count('WIP tasks')} task(s)",
            scope=clickup_scope, period=clickup_period,
            exclusions="Tasks with unavailable current status are not classified as WIP.",
            validation=clickup_validation,
        ),
    }
    st.success("ClickUp analysis completed.")
    st.caption(
        f"Space: {result['space_name']} · Cutoff: {result['cutoff']} · "
        f"Scope: {result.get('filter_summary', 'All tasks in the selected ClickUp Space')}"
    )

    tab_dashboard, tab_process, tab_individual, tab_tasks, tab_quality = st.tabs([
        "Executive Dashboard",
        "Process Analysis",
        "Individual Achievements",
        "Task Detail",
        "Data Quality",
    ])

    with tab_dashboard:
        st.subheader("Executive Overview")
        population = build_population(
            clickup_total,
            clickup_known,
            {"Unknown status/history": clickup_total - clickup_known},
        )
        cards = st.columns(3)
        cards[0].metric("Total Tasks", count("Total tasks"), help=clickup_help["total"])
        render_population_card(cards[1], population)
        cards[2].metric("Completion Rate", percent("Completion rate (%)"), help=clickup_help["completion"])
        cards = st.columns(4)
        cards[0].metric("Completed", count("Completed tasks"), help=clickup_help["completed"])
        cards[1].metric("On-Time Rate", percent("On-time completion rate (%)"), help=clickup_help["on_time"])
        cards[2].metric("Open Overdue", count("Open overdue tasks"), help=clickup_help["overdue"])
        cards[3].metric("WIP Tasks", count("WIP tasks"), help=clickup_help["wip"])

        st.divider()
        st.subheader("Management Averages")
        st.caption(
            "Execution, lead-time, and time-to-start averages are measured in elapsed hours. "
            "Timing values stay unavailable when ClickUp dates conflict. Due Variance and overdue measures use calendar days."
        )


        completed_cards = st.columns(3)
        completed_cards[0].metric("Avg Execution Time (Completed)", average("Average execution hours", "h"))
        completed_cards[1].metric("Avg Lead Time (Completed)", average("Average lead time hours", "h"))
        completed_cards[2].metric("Avg Time to Start", average("Average time to start hours", "h"))
        variance_cards = st.columns(3)
        late_tasks = result["late_completed_tasks"]
        late_mean = None if late_tasks.empty else pd.to_numeric(late_tasks["Due Variance (days)"], errors="coerce").mean()
        overdue_tasks = result["overdue_tasks"]
        overdue_mean = None if overdue_tasks.empty else pd.to_numeric(overdue_tasks["Overdue Days"], errors="coerce").mean()
        variance_cards[0].metric("Avg Late Completion", _format_average(late_mean, "days"))
        variance_cards[1].metric("Avg Open Overdue", _format_average(overdue_mean, "days"))
        variance_cards[2].metric("Avg Due Variance (Completed)", average("Average due variance (days)", "days"))

        st.divider()
        st.subheader("Weekly Task Flow")
        st.caption("Tasks created and tasks completed are grouped by Monday-starting weeks.")
        weekly = result["weekly_flow"]
        if weekly.empty:
            st.info("No created or completed dates are available for a weekly trend.")
        else:
            range_options = {
                "All available weeks": None,
                "Last 4 weeks": 4,
                "Last 12 weeks": 12,
                "Last 24 weeks": 24,
                "Last 36 weeks": 36,
                "Last 48 weeks": 48,
            }
            selected_range = st.selectbox("Trend period", list(range_options), key="clickup_weekly_flow_period")
            visible_weekly = weekly.tail(range_options[selected_range]).copy() if range_options[selected_range] else weekly.copy()
            st.line_chart(
                visible_weekly.set_index("Week Starting")[["Tasks Created", "Tasks Completed"]],
                use_container_width=True,
            )
            st.dataframe(visible_weekly, hide_index=True, use_container_width=True)

        st.divider()
        left, right = st.columns(2)
        with left:
            st.subheader("Task Distribution by Status")
            status_counts = result["status_counts"]
            if status_counts.empty:
                st.info("No status values are available.")
            else:
                st.bar_chart(status_counts.set_index("Status")["Tasks"], use_container_width=True)

            st.subheader("Open Tasks by Due Status")
            due_status = result["due_status_summary"]
            if due_status.empty:
                st.info("No open tasks are available for due-date grouping.")
            else:
                st.bar_chart(due_status.set_index("Due Status")["Tasks"], use_container_width=True)
                st.dataframe(due_status, hide_index=True, use_container_width=True)
            issue_cards = st.columns(2)
            issue_cards[0].metric("Completed Late", count("Completed late tasks"))
            issue_cards[1].metric("Tasks Missing Due Date", len(result["missing_due_tasks"]))

        with right:
            st.subheader("Work Distribution by Assignee")
            assignee_summary = result["assignee_summary"]
            if assignee_summary.empty:
                st.info("No assignee information is available.")
            else:
                st.bar_chart(
                    assignee_summary.set_index("Assignee")[["Completed", "Open", "WIP"]],
                    use_container_width=True,
                )

            st.subheader("Due Variance Distribution")
            st.caption("Positive values mean late completion or an overdue open task; negative values mean early completion or time remaining.")
            due_variance = result["due_variance_summary"]
            if due_variance.empty:
                st.info("No tasks with a due date are available for Due Variance analysis.")
            else:
                st.bar_chart(due_variance.set_index("Due Variance Category")[["Tasks"]], use_container_width=True)
                st.dataframe(due_variance, hide_index=True, use_container_width=True)

        st.subheader("Total Time in Status")
        status_summary = result["status_summary"]
        if status_summary.empty:
            st.info("Total time in Status was not available for this selected ClickUp scope. Missing values are not treated as zero.")
        else:
            st.bar_chart(status_summary.set_index("Status")["Total Hours"], use_container_width=True)
            st.dataframe(status_summary, hide_index=True, use_container_width=True)

    with tab_process:
        st.subheader("Process Scope and Definitions")
        st.dataframe(_streamlit_safe_frame(result["analysis_context"]), hide_index=True, use_container_width=True)
        st.subheader("Workflow and Status Analysis")
        if result["status_summary"].empty:
            st.info("No status-duration values were returned by ClickUp for this selected scope.")
        else:
            st.dataframe(result["status_summary"], hide_index=True, use_container_width=True)
            with st.expander("Status-duration detail"):
                st.dataframe(result["status_detail"], hide_index=True, use_container_width=True)
        st.subheader("Open Overdue Tasks")
        if result["overdue_tasks"].empty:
            st.success("No open overdue tasks were identified in the current scope.")
        else:
            st.dataframe(result["overdue_tasks"], hide_index=True, use_container_width=True)
        st.subheader("Completed Late Tasks")
        if result["late_completed_tasks"].empty:
            st.success("No completed-late tasks were identified in the current scope.")
        else:
            st.dataframe(result["late_completed_tasks"], hide_index=True, use_container_width=True)
        st.subheader("Evidence and Follow-up")
        st.dataframe(result["findings"], hide_index=True, use_container_width=True)
        with st.expander("Metric definitions"):
            st.dataframe(result["metric_definitions"], hide_index=True, use_container_width=True)
        st.info("Activity-history collection is disabled. Rework, replanning, and transition-event metrics are not inferred from the ClickUp task snapshot.")

    with tab_individual:
        st.subheader("Individual Achievements and Assignment Summary")
        st.caption("Assignment uses the assignee snapshot returned by ClickUp. These values support workload visibility and do not establish individual contribution.")
        st.dataframe(result["assignee_summary"], hide_index=True, use_container_width=True)
        assignee_options = ["All", *sorted(tasks["Assignee"].fillna("Unassigned").astype(str).unique())] if not tasks.empty else ["All"]
        selected_assignee = st.selectbox("Show tasks for", assignee_options, key="clickup_individual_assignee")
        assignee_tasks = tasks if selected_assignee == "All" else tasks[tasks["Assignee"].fillna("Unassigned").astype(str).eq(selected_assignee)]
        columns = [
            "Task ID", "Task Name", "Priority", "Current Status", "Created", "Start Date", "Due Date", "Completed",
            "Due Variance (days)", "On Time?", "Execution Hours", "Lead Time Hours", "Total Time in Status (min)",
        ]
        st.subheader("Task details")
        st.dataframe(assignee_tasks[[column for column in columns if column in assignee_tasks]], hide_index=True, use_container_width=True)

    with tab_tasks:
        st.subheader("Task-Level Evaluation")
        filtered = tasks.copy()
        if not filtered.empty:
            filter_left, filter_middle, filter_right = st.columns(3)
            with filter_left:
                assignee_options = ["All", *sorted(filtered["Assignee"].fillna("Unassigned").astype(str).unique())]
                task_assignee = st.selectbox("Filter by assignee", assignee_options, key="clickup_task_assignee")
                if task_assignee != "All":
                    filtered = filtered[filtered["Assignee"].fillna("Unassigned").astype(str).eq(task_assignee)]
            with filter_middle:
                status_options = ["All", *sorted(filtered["Current Status"].fillna("Unavailable").astype(str).unique())]
                task_status = st.selectbox("Filter by status", status_options, key="clickup_task_status")
                if task_status != "All":
                    filtered = filtered[filtered["Current Status"].fillna("Unavailable").astype(str).eq(task_status)]
            with filter_right:
                priority_options = ["All", *sorted(filtered["Priority"].fillna("No priority").astype(str).unique())]
                task_priority = st.selectbox("Filter by priority", priority_options, key="clickup_task_priority")
                if task_priority != "All":
                    filtered = filtered[filtered["Priority"].fillna("No priority").astype(str).eq(task_priority)]
        detail_columns = [
            "Task ID", "Task Name", "Assignee", "Created By", "Priority", "Task Type", "Tags", "Location/List",
            "Current Status", "Created", "Updated", "Start Date", "Due Date", "Completed", "Completed?", "Open?",
            "On Time?", "Due Variance (days)", "Due Variance Basis", "Due Variance Category", "Overdue Days",
            "Execution Hours", "Lead Time Hours", "Time to Start Hours", "Timing Data Status", "Time Estimate Hours", "Time Tracked Hours",
            "Current Status Time (min)", "Total Time in Status (min)",
        ]
        st.dataframe(filtered[[column for column in detail_columns if column in filtered]], hide_index=True, use_container_width=True)

    with tab_quality:
        st.subheader("Data Quality and Coverage")
        st.dataframe(_streamlit_safe_frame(result["quality"]), hide_index=True, use_container_width=True)
        st.info("Unavailable ClickUp fields remain unavailable; they are not replaced with zero. Jira was not used or changed by this analysis.")
        with st.expander("Metric definitions"):
            st.dataframe(result["metric_definitions"], hide_index=True, use_container_width=True)

    st.divider()
    download_left, download_right = st.columns(2)
    with download_left:
        clickup_report_key = f"{result['space_name']}:{result['cutoff']}:{_frame_signature(tasks)}"
        if st.session_state.get("clickup_excel_report_key") != clickup_report_key:
            st.session_state["clickup_excel_report"] = analysis_excel(result)
            st.session_state["clickup_excel_report_key"] = clickup_report_key
        st.download_button(
            "Download ClickUp Analysis Excel",
            data=st.session_state["clickup_excel_report"],
            file_name="clickup_performance_analysis.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with download_right:
        report_key = clickup_report_key
        if st.session_state.get("clickup_word_report_key") != report_key:
            st.session_state["clickup_word_report"] = create_clickup_word_report(result)
            st.session_state["clickup_word_report_key"] = report_key
        st.download_button(
            "Download ClickUp Analysis Word Report",
            data=st.session_state["clickup_word_report"],
            file_name="clickup_performance_analysis.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )


def show_department_analysis(result: dict) -> None:
    """Render the isolated department dashboard and approved exports."""
    def value(name, default=None):
        matches = result["kpis"].loc[result["kpis"]["KPI"].eq(name), "Value"]
        return default if matches.empty else matches.iloc[0]

    def percent(name):
        current = value(name)
        return "N/A" if current is None or pd.isna(current) else f"{float(current):.1f}%"

    st.success("Department performance analysis completed.")
    st.caption(
        f"Department: {result['department_name']} · Source Spaces: "
        f"{', '.join(map(str, result.get('space_names', [result['space_name']])))} · "
        f"Cutoff: {result['cutoff']}"
    )
    cards = st.columns(3)
    cards[0].metric("Total Tasks", int(value("Total Tasks", 0)))
    population = result.get("completion_rate_population")
    if population is None:
        denominator = result["kpis"].loc[
            result["kpis"]["KPI"].eq("Task Completion Rate (%)"), "Denominator"
        ]
        population = build_population(
            int(value("Total Tasks", 0)),
            int(denominator.iloc[0]) if not denominator.empty and pd.notna(denominator.iloc[0]) else 0,
        )
    render_population_card(cards[1], population)
    cards[2].metric("Task Completion Rate", percent("Task Completion Rate (%)"))
    cards = st.columns(4)
    cards[0].metric("On-Time Completion Rate", percent("On-Time Completion Rate (%)"))
    cards[1].metric("Open Overdue Tasks", int(value("Open Overdue Tasks", 0)))
    lead = value("Average Lead Time (hours)")
    cards[2].metric("Average Lead Time", "Unavailable" if lead is None or pd.isna(lead) else f"{float(lead):.1f} h")
    cards[3].metric("Workflow Exception Rate", percent("Workflow Exception Rate (%)"))
    st.caption(
        f"Cancelled/Rejected: {int(value('Cancelled/Rejected Tasks', 0))} · "
        f"Rate: {percent('Cancellation/Rejected Rate (%)')} · "
        "Workflow exception rate remains unavailable when activity history is not collected."
    )
    chart_left, chart_right = st.columns(2)
    with chart_left:
        st.subheader("Task Status Distribution")
        if result["status_counts"].empty:
            st.info("No status data is available.")
        else:
            st.bar_chart(result["status_counts"].set_index("Status")["Tasks"], use_container_width=True)
        st.subheader("Workload by Employee")
        employees = result["employee_breakdown"]
        employee_columns = [c for c in ["Open", "WIP", "Overdue"] if c in employees]
        if employees.empty:
            st.info("No assignee data is available.")
        else:
            st.bar_chart(employees.set_index("Assignee")[employee_columns], use_container_width=True)
    with chart_right:
        st.subheader("Delivery Performance")
        if result["due_variance_summary"].empty:
            st.info("No due-date delivery data is available.")
        else:
            st.bar_chart(result["due_variance_summary"].set_index("Due Variance Category")["Tasks"], use_container_width=True)
        st.subheader("Department Performance Trend")
        if result["weekly_flow"].empty:
            st.info("No dated tasks are available for a trend.")
        else:
            st.line_chart(result["weekly_flow"].set_index("Week Starting")[["Tasks Created", "Tasks Completed"]], use_container_width=True)
    summary_tab, attention_tab, quality_tab = st.tabs(
        ["Department Employee Summary", "Tasks Requiring Attention", "Bottlenecks & Data Quality"]
    )
    with summary_tab:
        st.caption("This table explains department workload; it is not an individual employee performance score.")
        st.dataframe(result["employee_breakdown"], hide_index=True, use_container_width=True)
    with attention_tab:
        if result["attention"].empty:
            st.success("No tasks requiring attention were identified.")
        else:
            st.dataframe(result["attention"], hide_index=True, use_container_width=True)
    with quality_tab:
        st.subheader("Bottleneck Candidates")
        if result["bottlenecks"].empty:
            st.info("Status-duration evidence is unavailable.")
        else:
            st.dataframe(result["bottlenecks"], hide_index=True, use_container_width=True)
        st.subheader("Data Quality")
        if result["department_quality"].empty:
            st.success("No material data-quality issues were identified.")
        else:
            st.dataframe(result["department_quality"], hide_index=True, use_container_width=True)
    st.divider()
    excel_name = f"Department_Performance_{_filename_component(result['department_name'])}.xlsx"
    word_name = f"Department_Performance_{_filename_component(result['department_name'])}.docx"
    department_report_key = (
        f"{result.get('department_id')}:{result.get('data_source')}:"
        f"{result.get('cutoff')}:{_frame_signature(result.get('tasks'))}"
    )
    if st.session_state.get("department_report_key") != department_report_key:
        st.session_state["department_excel_report"] = department_excel(result)
        st.session_state["department_word_report"] = department_word(result)
        st.session_state["department_report_key"] = department_report_key
    left, right = st.columns(2)
    left.download_button("Download Excel Report", st.session_state["department_excel_report"], excel_name,
                         mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                         use_container_width=True)
    right.download_button("Download Word Report", st.session_state["department_word_report"], word_name,
                          mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                          use_container_width=True)


from automation_ui import invalidate_selection, require_sign_in, render_collection

settings = require_sign_in()

LOGO_PATH = ROOT_DIR / "assets" / "bidayah_logo_white.png"
HEADER_LOGO_PATH = ROOT_DIR / "assets" / "bidayah_logo_header.png"
display_logo_path = HEADER_LOGO_PATH if HEADER_LOGO_PATH.exists() else LOGO_PATH
if display_logo_path.exists():
    st.image(str(display_logo_path), width=190)

st.title(APP_TITLE)
analysis_mode_label = st.radio(
    "Analysis type",
    ["Company Performance", "Department Performance", "Employee Performance", "Project Performance"],
    horizontal=True,
    key="analysis_type_selector",
    help="The existing employee/project analysis and department analysis use separate results and exports.",
)
analysis_mode = {
    "Employee Performance": "employee",
    "Project Performance": "project",
    "Department Performance": "department",
    "Company Performance": "company",
}[analysis_mode_label]
if st.session_state.get("active_analysis_mode") != analysis_mode:
    # A prepared department List must never be reused as a whole-space Company
    # source (or vice versa). Changing the analysis type starts a clean run.
    invalidate_selection()
    for key in (
        "clickup_analysis", "department_analysis", "company_analysis",
        "task_metrics", "process_data", "validation_log",
        "clickup_prepared_data", "clickup_run_analysis",
        "department_run_requested", "jira_department_run_requested",
        "department_selected_name", "department_period_start", "department_period_end",
        "department_search", "department_status", "department_assignee", "department_priority",
        "department_match", "department_tasks_cache_key", "department_tasks_cache",
        "department_list_checkpoints",
        "jira_department_period_start", "jira_department_period_end",
        "jira_department_prepared_data", "jira_department_run_analysis",
        "jira_department_preview_cache",
        "employee_snapshot", "employee_snapshot_key", "employee_selected_spaces",
        "employee_prepared", "employee_fingerprint", "employee_clickup_checkpoints",
        "employee_company_analysis",
        "company_prepared_items", "company_collection_attempted",
        "company_collection_error", "company_collection_errors",
        "company_partial_prepared_items",
        "company_prepared_jira", "company_prepared_clickup",
        "company_prepared_jira_spaces", "company_prepared_clickup_spaces",
        "company_selected_jira_spaces", "company_selected_clickup_spaces",
        "company_selected_company", "company_selected_company_name", "company_active_selection",
        "company_catalog", "company_catalog_revision", "company_collection_selection",
        "company_collection_company_name",
        "company_unified_project", "company_period_start", "company_period_end",
        "project_catalog_revision", "project_selected_jira_space", "project_selected_clickup_space",
        "project_selection_fingerprint", "project_jira_prepared", "project_clickup_prepared",
        "project_preview", "project_analysis", "project_scope_label", "project_scope_slug",
        "project_period_start", "project_period_end", "project_report_key",
        "project_clickup_checkpoints",
        "employee_filter_company", "employee_filter_projects", "employee_filter_status",
        "employee_filter_task_type", "employee_filter_priority", "employee_filter_due_state",
        "employee_filter_companies", "employee_filter_spaces", "employee_filter_statuses",
        "employee_filter_task_types", "employee_filter_priorities", "employee_filter_due_states",
    ):
        st.session_state.pop(key, None)
    st.session_state["active_analysis_mode"] = analysis_mode
st.caption("Choose an analysis, apply its filters, and run the available process analysis.")
if analysis_mode == "company":
    # Company mode discovers the connected source catalog before collecting any
    # tasks.  The selected ID is stable for the session; display names remain
    # presentation-only.
    company_catalog = discover_company_catalog(settings)
    if not company_catalog:
        st.selectbox(
            "Choose Company",
            ["Choose a company"],
            index=0,
            key="company_selected_company",
        )
        st.info(
            "Choose a company once a connected source is available. "
            "No companies were discovered from the connected Jira or ClickUp sources yet."
        )
        st.stop()
    render_company_source_mapping(company_catalog)
    company_options = [ALL_COMPANIES_ID, *(entry.company_id for entry in company_catalog)]
    company_labels = {
        ALL_COMPANIES_ID: f"All Companies ({len(company_catalog)} available)",
        **{entry.company_id: company_option_label(entry) for entry in company_catalog},
    }
    selected_company = st.selectbox(
        "Choose Company",
        ["Choose a company", *company_options],
        index=0,
        format_func=lambda value: "Choose a company" if value == "Choose a company" else company_labels[value],
        key="company_selected_company",
    )
    if selected_company == "Choose a company":
        st.info("Choose a company or All Companies to load its Jira Projects and ClickUp Spaces.")
        st.stop()
    selected_entry = next(
        (entry for entry in company_catalog if entry.company_id == selected_company),
        None,
    )
    selected_company_name = (
        "All Companies" if selected_company == ALL_COMPANIES_ID
        else selected_entry.company_name if selected_entry else selected_company
    )
    if st.session_state.get("company_active_selection") != selected_company:
        for key in (
            "company_prepared_items",
            "company_collection_attempted",
            "company_collection_error",
            "company_collection_errors",
            "company_partial_prepared_items",
            "company_analysis",
            "company_report_key",
        ):
            st.session_state.pop(key, None)
        st.session_state["company_active_selection"] = selected_company
    st.session_state["company_selected_company_name"] = selected_company_name

prepared_data, run_button = render_collection(
    settings,
    analysis_mode=analysis_mode,
    company_selection=selected_company if analysis_mode == "company" else None,
)
cutoff_text = prepared_data.cutoff if prepared_data else ""

if analysis_mode == "employee":
    run_button = st.button(
        "Run Analysis",
        type="primary",
        disabled=prepared_data is None,
        key="employee_run",
    )

if analysis_mode == "project":
    if st.session_state.get("project_analysis") is not None:
        render_project_result(st, st.session_state["project_analysis"])
    else:
        st.info("Select a Jira Space, a ClickUp Space, or one from each, then choose the analysis period.")
    st.stop()

# Company Performance collects the complete Jira + ClickUp scope
# automatically; the source-specific flows below remain unchanged.
if analysis_mode == "company":
    render_company_launcher(st)
    if st.session_state.get("company_analysis") is not None:
        render_company_result(st, st.session_state["company_analysis"])
        st.stop()

# An employee can have a Jira account and a ClickUp member account.  The
# collection UI prepares both snapshots, then the source-neutral Company
# domain combines them without changing the existing single-source analyses.
if analysis_mode == "employee" and run_button and isinstance(prepared_data, EmployeePreparedBundle):
    try:
        with st.spinner("Calculating the combined employee analysis..."):
            st.session_state["employee_company_analysis"] = build_company_analysis(
                period_start=None,
                period_end=None,
                jira_prepared=prepared_data.jira,
                clickup_prepared=prepared_data.clickup,
                unified_project=f"Employee: {prepared_data.employee_name}",
            )
        st.success("Combined Jira + ClickUp employee analysis completed successfully.")
    except (ValueError, TypeError) as exc:
        st.session_state.pop("employee_company_analysis", None)
        st.error(f"Combined Jira + ClickUp employee analysis could not be completed: {exc}")

if analysis_mode == "employee" and st.session_state.get("employee_company_analysis") is not None:
    render_company_result(
        st,
        st.session_state["employee_company_analysis"],
        scope_title=f"Employee Performance — {getattr(prepared_data, 'employee_name', 'Selected Employee')}",
        close_state_key="employee_company_analysis",
        download_stem="employee_performance",
        scope_key="employee",
        scope_label="Employee",
    )
    st.stop()

if analysis_mode != "company" and run_button and prepared_data is not None:
    if st.session_state.get("data_source") == "ClickUp":
        try:
            with st.spinner("Calculating ClickUp performance analysis..."):
                clickup_result = analyze_clickup(prepared_data)
                if analysis_mode == "department":
                    st.session_state["department_analysis"] = build_department_result(clickup_result, prepared_data)
                    st.session_state.pop("clickup_analysis", None)
                else:
                    st.session_state["clickup_analysis"] = clickup_result
                    st.session_state.pop("department_analysis", None)
            st.success("ClickUp analysis completed successfully.")
        except Exception as exc:
            st.session_state.pop("clickup_analysis", None)
            st.error(f"ClickUp analysis could not be completed: {exc}")
    else:
        try:
            with st.spinner("Reading data and calculating Jira metrics..."):
                task_metrics, validation_log, process_data = run_analysis(
                    BytesIO(prepared_data.xlsx),
                    BytesIO(prepared_data.history_json),
                    "", "", "",
                    prepared_data.cutoff,
                    prepared_data.source_timezone,
                    history_source="History JSON",
                    context_overrides={
                        "Process Name": prepared_data.space_name,
                        "Evaluation Scope": "Selected Jira work items",
                        "Dataset Type": "Jira API collection",
                    },
                    period_start=(
                        getattr(prepared_data, "period_start", None)
                        if analysis_mode == "department" else None
                    ),
                )
            if task_metrics.empty:
                st.session_state.pop("task_metrics", None)
                st.session_state.pop("process_data", None)
                st.info("No selected work items existed at the evaluation cutoff. Click Done to collect a newer snapshot.")
            else:
                if analysis_mode == "department":
                    jira_department_result = build_jira_department_result(
                        task_metrics, process_data, prepared_data.space_name, prepared_data.cutoff,
                    )
                    jira_department_result["space_names"] = getattr(
                        prepared_data, "space_names", [prepared_data.space_name]
                    )
                    st.session_state["department_analysis"] = jira_department_result
                    for key in ("task_metrics", "process_data", "validation_log", "clickup_analysis"):
                        st.session_state.pop(key, None)
                else:
                    st.session_state["process_data"] = process_data
                    st.session_state["task_metrics"] = task_metrics
                    st.session_state["validation_log"] = validation_log
                    st.session_state["cutoff_text"] = prepared_data.cutoff
                    st.session_state.pop("department_analysis", None)
                st.success("Analysis completed successfully.")
        except Exception:
            st.session_state.pop("task_metrics", None)
            st.session_state.pop("process_data", None)
            st.error("Analysis could not be completed for this dataset. Check the source data or contact the administrator.")

if "department_analysis" in st.session_state:
    show_department_analysis(st.session_state["department_analysis"])
    st.stop()

if "clickup_analysis" in st.session_state:
    clickup_result = st.session_state["clickup_analysis"]
    if analysis_mode == "employee":
        filtered_clickup_tasks = render_employee_post_analysis_filters(
            clickup_result["tasks"],
            "ClickUp",
        )
        clickup_result = recalculate_clickup_analysis(
            clickup_result,
            filtered_clickup_tasks,
        )
    show_clickup_analysis(clickup_result)
    st.stop()

if "task_metrics" in st.session_state and "status_known" not in st.session_state["task_metrics"].columns:
    st.session_state.pop("task_metrics", None)
    st.session_state.pop("process_data", None)
    st.info("The analysis has been updated. Run Analysis to calculate the new process metrics.")

if "task_metrics" in st.session_state:
    task_metrics = st.session_state["task_metrics"]
    if analysis_mode == "employee":
        task_metrics = render_employee_post_analysis_filters(task_metrics, "Jira")
        if task_metrics.empty:
            st.warning("No analysed tasks match the selected Employee filters.")
            st.stop()
    validation_log = st.session_state.get(
        "validation_log",
        [],
    )

    if analysis_mode == "employee":
        process_data = process_tables(
            task_metrics,
            histories=task_metrics.attrs.get("histories", {}),
            context=task_metrics.attrs.get("process_context", {}),
            validation_log=validation_log,
        )
    else:
        process_data = st.session_state.get("process_data") or process_tables(task_metrics)
    st.caption("Results calculated through: " + str(task_metrics.iloc[0]["evaluation_cutoff"]))
    if not task_metrics["history_complete"].all():
        st.warning("Some task histories cannot be verified. Affected statuses and metrics are unavailable; see Data Quality.")
    tab_dashboard, tab_process, tab_individual, tab_tasks, tab_quality = st.tabs(
        [
            "Executive Dashboard",
            "Process Analysis",
            "Individual Achievements",
            "Task Detail",
            "Data Quality",
        ]
    )

    with tab_dashboard:
        show_executive_dashboard(
            task_metrics,
            process_data,
        )

    with tab_process:
        show_process_analysis(task_metrics, process_data)

    with tab_individual:
        show_assignment_summary(task_metrics)

    with tab_tasks:
        show_task_detail(
            task_metrics,
        )

    with tab_quality:
        st.dataframe(process_data["data_quality"], hide_index=True, use_container_width=True)
        show_data_quality(
            task_metrics,
            validation_log,
        )

    st.divider()

    st.subheader("Downloads")

    overall = aggregate(
        task_metrics,
        [],
    )

    by_assignee = aggregate(
        task_metrics,
        ["assignee_name"],
    )

    by_issue_type = aggregate(
        task_metrics,
        ["issue_type"],
    )

    download_columns = st.columns(2)

    with download_columns[0]:
        st.download_button(
            "Download Excel Analysis",
            data=create_excel_download(
                task_metrics,
                overall,
                by_assignee,
                by_issue_type,
                tables=process_data,
            ),
            file_name="jira_performance_analysis.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            use_container_width=True,
        )

    with download_columns[1]:
        try:
            report_title = (
                "Employee Task Execution Performance Evaluation"
                if analysis_mode == "employee"
                else "Jira Process Performance Evaluation Report"
            )
            report_bytes = create_word_report_download(
                task_metrics,
                st.session_state.get(
                    "cutoff_text",
                    cutoff_text,
                ),
                tables=process_data,
                title=report_title,
            )

            st.download_button(
                "Download Word Report",
                data=report_bytes,
                file_name="jira_performance_report.docx",
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "wordprocessingml.document"
                ),
                use_container_width=True,
            )

        except Exception as exc:
            st.warning(
                f"Word report is unavailable: {exc}"
            )

else:
    if analysis_mode == "employee":
        st.info("Choose an employee, load the assigned tasks, and click Done before running the analysis.")
    elif st.session_state.get("data_source") == "ClickUp":
        st.info(
            "Select a ClickUp Space, click Done, then download the ClickUp source Excel."
        )
    else:
        st.info(
            (
                "Select a Jira space, choose filters, and click Done. "
                "When your data is ready, click Run Analysis."
            )
        )
