from __future__ import annotations

import json
import os
import sys
import tempfile
from datetime import datetime
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


CONFIG_DIR = ROOT_DIR / "configs"
INPUT_CONFIG_PATH = CONFIG_DIR / "jira_input_config.json"
METRICS_CONFIG_PATH = CONFIG_DIR / "jira_metrics_config.json"


st.set_page_config(
    page_title="Jira Process Performance Dashboard",
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
            title="Jira Process Performance Evaluation Report",
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


def format_percentage(
    numerator: int,
    denominator: int,
) -> str:
    if denominator == 0:
        return "Unavailable"

    return f"{numerator / denominator * 100:.1f}%"


def calculate_dashboard_values(task_metrics):
    row = aggregate(task_metrics, []).iloc[0]
    return {"total": row["total_tasks"], "completed": row["completed_tasks"],
            "rejected": row["rejected_tasks"], "open": row["open_tasks"], "wip": row["wip_tasks"],
            "completion_rate": format_percentage(int(row["completed_tasks"]), int(row["total_tasks"])),
            "on_time_rate": format_percentage(int(row["on_time_tasks"]), int(row["on_time_valid_tasks"])),
            "overdue": row["overdue_open_tasks"], "rework": row["tasks_with_rework"]}


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
        "Last 12 weeks (quarter)": 12,
        "Last 52 weeks (year)": 52,
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

    if task_metrics.empty:
        raise ValueError("No tasks have a valid creation date on or before this cutoff.")
    context = workbook_context(workbook)
    context.update(context_overrides or {})
    context["History source"] = history_source
    task_metrics.attrs["process_context"] = context
    task_metrics.attrs["selected_sheet"] = sheet_name
    task_metrics.attrs["validation_log"] = validation_log
    task_metrics.attrs["cutoff"] = cutoff.isoformat()

    tables = process_tables(task_metrics, histories, context, validation_log)
    return task_metrics, validation_log, tables


def show_executive_dashboard(
    task_metrics: pd.DataFrame,
    process_data: dict | None = None,
) -> None:
    values = calculate_dashboard_values(
        task_metrics,
    )

    columns = st.columns(6)

    columns[0].metric(
        "Total Tasks",
        values["total"],
    )

    columns[1].metric(
        "Completed",
        values["completed"],
    )

    columns[2].metric(
        "Completion Rate",
        values["completion_rate"],
    )

    columns[3].metric(
        "On-Time Rate",
        values["on_time_rate"],
    )

    columns[4].metric(
        "Open Overdue",
        values["overdue"],
    )

    columns[5].metric(
        "WIP Tasks",
        values["wip"],
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
    st.dataframe(tables["process_context"], hide_index=True, use_container_width=True)
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
    event_counts = {kind: row["total_" + kind + "_count"] for kind in ["rework", "replanning", "re_evaluation"]}
    st.write("Event totals (a task may have multiple events):", event_counts)
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


from automation_ui import require_sign_in, render_collection

settings = require_sign_in()

st.title("Jira Process Performance Dashboard")
st.caption("Data collection v3.1 — select your Jira work items and run the existing process analysis")

prepared_data, run_button = render_collection(settings)
cutoff_text = prepared_data.cutoff if prepared_data else ""

if run_button and prepared_data is not None:
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
            )
        if task_metrics.empty:
            st.session_state.pop("task_metrics", None)
            st.session_state.pop("process_data", None)
            st.info("No selected work items existed at the evaluation cutoff. Click Done to collect a newer snapshot.")
        else:
            st.session_state["process_data"] = process_data
            st.session_state["task_metrics"] = task_metrics
            st.session_state["validation_log"] = validation_log
            st.session_state["cutoff_text"] = prepared_data.cutoff
            st.success("Analysis completed successfully.")
    except Exception:
        st.session_state.pop("task_metrics", None)
        st.session_state.pop("process_data", None)
        st.error("Analysis could not be completed for this dataset. Check the source data or contact the administrator.")

if "task_metrics" in st.session_state and "status_known" not in st.session_state["task_metrics"].columns:
    st.session_state.pop("task_metrics", None)
    st.session_state.pop("process_data", None)
    st.info("The analysis has been updated. Run Analysis to calculate the new process metrics.")

if "task_metrics" in st.session_state:
    task_metrics = st.session_state["task_metrics"]
    validation_log = st.session_state.get(
        "validation_log",
        [],
    )

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
            report_bytes = create_word_report_download(
                task_metrics,
                st.session_state.get(
                    "cutoff_text",
                    cutoff_text,
                ),
                tables=process_data,
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
    st.info(
        (
            "Select a Jira space, choose filters, and click Done. "
            "When your data is ready, click Run Analysis."
        )
    )
