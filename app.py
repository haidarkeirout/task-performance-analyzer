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

from jira_client import JiraClient
from jira_processor import choose_sheet, normalize_tasks
from metrics_engine import (
    aggregate,
    analyze_tasks,
    load_calendar,
    normalize_labels,
    parse_timestamp,
)
from report_builder import build_report


CONFIG_DIR = ROOT_DIR / "configs"
INPUT_CONFIG_PATH = CONFIG_DIR / "jira_input_config.json"
METRICS_CONFIG_PATH = CONFIG_DIR / "jira_metrics_config.json"


st.set_page_config(
    page_title="Jira Executive Performance Dashboard",
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


def create_excel_download(
    task_metrics: pd.DataFrame,
    overall: pd.DataFrame,
    by_assignee: pd.DataFrame,
    by_issue_type: pd.DataFrame,
) -> bytes:
    buffer = BytesIO()

    export_tasks = task_metrics.copy()

    for column in ["labels", "time_in_status"]:
        if column in export_tasks.columns:
            export_tasks[column] = export_tasks[column].apply(
                lambda value: json.dumps(
                    value,
                    ensure_ascii=False,
                )
                if isinstance(value, (list, dict))
                else value
            )

    with pd.ExcelWriter(
        buffer,
        engine="openpyxl",
    ) as writer:
        export_tasks.to_excel(
            writer,
            index=False,
            sheet_name="task_metrics",
        )

        overall.to_excel(
            writer,
            index=False,
            sheet_name="overall_summary",
        )

        by_assignee.to_excel(
            writer,
            index=False,
            sheet_name="by_assignee",
        )

        by_issue_type.to_excel(
            writer,
            index=False,
            sheet_name="by_issue_type",
        )

    return buffer.getvalue()


def create_word_report_download(
    task_metrics: pd.DataFrame,
    cutoff: str,
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
            title="Jira Task Performance Evaluation Report",
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


def calculate_dashboard_values(
    task_metrics: pd.DataFrame,
) -> dict[str, object]:
    total = int(
        task_metrics["issue_key"].nunique()
    ) if "issue_key" in task_metrics.columns else len(task_metrics)

    completed = int(
        task_metrics["is_completed"].sum()
    ) if "is_completed" in task_metrics.columns else 0

    rejected = int(
        task_metrics["is_rejected"].sum()
    ) if "is_rejected" in task_metrics.columns else 0

    open_tasks = int(
        task_metrics["is_open"].sum()
    ) if "is_open" in task_metrics.columns else 0

    wip = int(
        task_metrics["is_wip"].sum()
    ) if "is_wip" in task_metrics.columns else 0

    valid_on_time = task_metrics[
        task_metrics["on_time_completion"].notna()
    ] if "on_time_completion" in task_metrics.columns else pd.DataFrame()

    on_time = int(
        valid_on_time["on_time_completion"].sum()
    ) if not valid_on_time.empty else 0

    overdue = int(
        (
            task_metrics["overdue_days"].fillna(0) > 0
        ).sum()
    ) if "overdue_days" in task_metrics.columns else 0

    rework = int(
        (
            task_metrics["rework_count"].fillna(0) > 0
        ).sum()
    ) if "rework_count" in task_metrics.columns else 0

    return {
        "total": total,
        "completed": completed,
        "rejected": rejected,
        "open": open_tasks,
        "wip": wip,
        "completion_rate": format_percentage(
            completed,
            total,
        ),
        "on_time_rate": format_percentage(
            on_time,
            len(valid_on_time),
        ),
        "overdue": overdue,
        "rework": rework,
    }


def run_analysis(
    uploaded_excel,
    uploaded_history,
    base_url: str,
    email: str,
    token: str,
    cutoff_text: str,
    source_timezone: str,
) -> tuple[pd.DataFrame, list[dict]]:
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

    if uploaded_history is not None:
        histories = read_history_file(
            uploaded_history,
        )
    else:
        if not base_url or not email or not token:
            raise ValueError(
                "Enter Jira URL, email, and API token, "
                "or upload a history JSON file."
            )

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

    task_metrics.attrs["selected_sheet"] = sheet_name
    task_metrics.attrs["validation_log"] = validation_log
    task_metrics.attrs["cutoff"] = cutoff.isoformat()

    return task_metrics, validation_log


def show_executive_dashboard(
    task_metrics: pd.DataFrame,
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
                    assignee_frame.set_index("Assignee")[["Completed", "Open"]],
                    use_container_width=True,
                )

    st.subheader(
        "Executive Exceptions Requiring Review"
    )

    exception_columns = [
        "issue_key",
        "task_name",
        "issue_type",
        "assignee_name",
        "status_at_cutoff",
        "overdue_days",
        "rework_count",
        "replanning_count",
    ]

    available_columns = [
        column
        for column in exception_columns
        if column in task_metrics.columns
    ]

    exceptions = task_metrics.copy()

    if "overdue_days" in exceptions.columns:
        exceptions = exceptions[
            exceptions["overdue_days"].fillna(0) > 0
        ]

    if exceptions.empty:
        st.success(
            "No overdue tasks were identified in the current scope."
        )
    else:
        st.dataframe(
            exceptions[available_columns],
            use_container_width=True,
            hide_index=True,
        )


def show_individual_achievements(
    task_metrics: pd.DataFrame,
) -> None:
    st.subheader(
        "Individual Achievement Profiles"
    )

    if "assignee_name" not in task_metrics.columns:
        st.warning(
            "Assignee information is unavailable."
        )
        return

    assignees = sorted(
        task_metrics["assignee_name"]
        .fillna("Assignee unavailable")
        .astype(str)
        .unique()
        .tolist()
    )

    selected_assignee = st.selectbox(
        "Select an assignee",
        assignees,
    )

    selected = task_metrics[
        task_metrics["assignee_name"]
        .fillna("Assignee unavailable")
        .astype(str)
        == selected_assignee
    ]

    completed = selected[
        selected["is_completed"] == True
    ]

    valid_on_time = completed[
        completed["on_time_completion"].notna()
    ] if "on_time_completion" in completed.columns else pd.DataFrame()

    on_time_count = int(
        valid_on_time["on_time_completion"].sum()
    ) if not valid_on_time.empty else 0

    columns = st.columns(5)

    columns[0].metric(
        "Assigned Tasks",
        len(selected),
    )

    columns[1].metric(
        "Completed",
        len(completed),
    )

    columns[2].metric(
        "Completion Rate",
        format_percentage(
            len(completed),
            len(selected),
        ),
    )

    columns[3].metric(
        "On-Time Rate",
        format_percentage(
            on_time_count,
            len(valid_on_time),
        ),
    )

    columns[4].metric(
        "Open Tasks",
        int(
            selected["is_open"].sum()
        ) if "is_open" in selected.columns else 0,
    )

    st.write(
        (
            f"Evidence-based achievement view for "
            f"**{selected_assignee}**. "
            "Durations describe process timing and do not "
            "automatically prove effort or quality."
        )
    )

    completed_columns = [
        "issue_key",
        "task_name",
        "issue_type",
        "labels",
        "completed_at",
        "due_date",
        "on_time_completion",
        "execution_business_hours",
        "rework_count",
    ]

    completed_columns = [
        column
        for column in completed_columns
        if column in completed.columns
    ]

    if completed.empty:
        st.info(
            "No completed tasks were verified for this assignee."
        )
    else:
        st.dataframe(
            completed[completed_columns],
            use_container_width=True,
            hide_index=True,
        )


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
        "status_at_cutoff",
        "created_at",
        "actual_start_at",
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


st.title(
    "Jira Executive Performance Dashboard"
)

st.caption(
    "Executive delivery oversight and individual achievement analysis"
)

with st.sidebar:
    st.header("Analysis Setup")

    uploaded_excel = st.file_uploader(
        "Upload Jira Excel export",
        type=["xlsx", "xls"],
    )

    uploaded_history = st.file_uploader(
        "Optional Jira history JSON",
        type=["json"],
        help=(
            "Use this for offline testing. If omitted, the app "
            "will retrieve history from Jira API."
        ),
    )

    st.subheader("Jira Connection")

    base_url = st.text_input(
        "Jira site URL",
        placeholder="https://your-domain.atlassian.net",
    )

    email = st.text_input(
        "Jira email",
        placeholder="your-email@example.com",
    )

    token = st.text_input(
        "Jira API token",
        type="password",
    )

    st.subheader("Evaluation Context")

    source_timezone = st.text_input(
        "Source timezone",
        value="Asia/Damascus",
    )

    default_cutoff = datetime.now(
        ZoneInfo("Asia/Damascus")
    ).replace(
        second=0,
        microsecond=0,
    ).isoformat()

    cutoff_text = st.text_input(
        "Evaluation cutoff",
        value=default_cutoff,
        help=(
            "Example: 2026-09-06T17:00:00+03:00"
        ),
    )

    run_button = st.button(
        "Run Analysis",
        type="primary",
        use_container_width=True,
    )

if run_button:
    if uploaded_excel is None:
        st.error(
            "Upload the Jira Excel export first."
        )
    else:
        try:
            with st.spinner(
                "Reading data and calculating Jira metrics..."
            ):
                task_metrics, validation_log = run_analysis(
                    uploaded_excel,
                    uploaded_history,
                    base_url,
                    email,
                    token,
                    cutoff_text,
                    source_timezone,
                )

            st.session_state["task_metrics"] = task_metrics
            st.session_state["validation_log"] = validation_log
            st.session_state["cutoff_text"] = cutoff_text

            st.success(
                "Analysis completed successfully."
            )

        except Exception as exc:
            st.error(
                f"Analysis failed: {exc}"
            )

if "task_metrics" in st.session_state:
    task_metrics = st.session_state["task_metrics"]
    validation_log = st.session_state.get(
        "validation_log",
        [],
    )

    tab_dashboard, tab_individual, tab_tasks, tab_quality = st.tabs(
        [
            "Executive Dashboard",
            "Individual Achievements",
            "Task Detail",
            "Data Quality",
        ]
    )

    with tab_dashboard:
        show_executive_dashboard(
            task_metrics,
        )

    with tab_individual:
        show_individual_achievements(
            task_metrics,
        )

    with tab_tasks:
        show_task_detail(
            task_metrics,
        )

    with tab_quality:
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
            "Upload the Jira Excel export, enter Jira credentials, "
            "or provide a previously retrieved history JSON file, "
            "then click Run Analysis."
        )
    )
