from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any

import pandas as pd
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.shared import Inches, Pt


def is_true(value: Any) -> bool:
    return value is True or str(value).strip().lower() == "true"


def format_value(value: Any) -> str:
    if value is None:
        return "Unavailable"

    if isinstance(value, float):
        if pd.isna(value):
            return "Unavailable"

        return f"{value:.2f}"

    if isinstance(value, bool):
        return "Yes" if value else "No"

    text = str(value).strip()

    if not text or text.lower() in {"nan", "none", "null"}:
        return "Unavailable"

    return text


def format_percent(
    numerator: int,
    denominator: int,
) -> str:
    if denominator == 0:
        return "Unavailable"

    return f"{numerator / denominator * 100:.1f}%"


def parse_labels(value: Any) -> list[str]:
    if value is None:
        return []

    if isinstance(value, list):
        return [str(item) for item in value]

    text = str(value).strip()

    if not text or text.lower() in {
        "nan",
        "none",
        "null",
        "[]",
    }:
        return []

    try:
        parsed = json.loads(text)

        if isinstance(parsed, list):
            return [str(item) for item in parsed]
    except json.JSONDecodeError:
        pass

    return [text]


def set_document_style(document: Document) -> None:
    styles = document.styles

    normal = styles["Normal"]
    normal.font.name = "Arial"
    normal.font.size = Pt(10)

    for section in document.sections:
        section.top_margin = Inches(0.65)
        section.bottom_margin = Inches(0.65)
        section.left_margin = Inches(0.7)
        section.right_margin = Inches(0.7)


def add_title(
    document: Document,
    title: str,
) -> None:
    paragraph = document.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

    run = paragraph.add_run(title)
    run.bold = True
    run.font.size = Pt(20)
    run.font.name = "Arial"


def add_heading(
    document: Document,
    text: str,
    level: int = 1,
) -> None:
    heading = document.add_heading(
        text,
        level=level,
    )
    heading.style.font.name = "Arial"


def add_paragraph(
    document: Document,
    text: str,
    bold_prefix: str | None = None,
) -> None:
    paragraph = document.add_paragraph()

    if bold_prefix and text.startswith(bold_prefix):
        prefix_run = paragraph.add_run(bold_prefix)
        prefix_run.bold = True
        paragraph.add_run(text[len(bold_prefix):])
    else:
        paragraph.add_run(text)


def add_key_value_table(
    document: Document,
    values: dict[str, Any],
) -> None:
    table = document.add_table(
        rows=0,
        cols=2,
    )

    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"

    for key, value in values.items():
        cells = table.add_row().cells
        cells[0].text = str(key)
        cells[1].text = format_value(value)

        cells[0].paragraphs[0].runs[0].bold = True


def add_task_table(
    document: Document,
    frame: pd.DataFrame,
) -> None:
    columns = [
        "issue_key",
        "task_name",
        "issue_type",
        "status_at_cutoff",
        "completed_at",
        "due_date",
        "on_time_completion",
        "execution_business_hours",
        "rework_count",
    ]

    available_columns = [
        column
        for column in columns
        if column in frame.columns
    ]

    if not available_columns:
        add_paragraph(
            document,
            "No task-level data is available.",
        )
        return

    table = document.add_table(
        rows=1,
        cols=len(available_columns),
    )

    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"

    headers = table.rows[0].cells

    for index, column in enumerate(available_columns):
        headers[index].text = column.replace(
            "_",
            " ",
        ).title()

        headers[index].paragraphs[0].runs[0].bold = True

    for _, row in frame.iterrows():
        cells = table.add_row().cells

        for index, column in enumerate(available_columns):
            cells[index].text = format_value(
                row.get(column)
            )


def calculate_overall_metrics(
    frame: pd.DataFrame,
) -> dict[str, Any]:
    total = int(
        frame["issue_key"].nunique()
    ) if "issue_key" in frame.columns else len(frame)

    completed = int(
        frame["is_completed"].sum()
    ) if "is_completed" in frame.columns else 0

    rejected = int(
        frame["is_rejected"].sum()
    ) if "is_rejected" in frame.columns else 0

    open_tasks = int(
        frame["is_open"].sum()
    ) if "is_open" in frame.columns else 0

    wip = int(
        frame["is_wip"].sum()
    ) if "is_wip" in frame.columns else 0

    valid_on_time = frame[
        frame["on_time_completion"].notna()
    ] if "on_time_completion" in frame.columns else pd.DataFrame()

    on_time = int(
        valid_on_time["on_time_completion"].sum()
    ) if not valid_on_time.empty else 0

    valid_overdue = frame[
        frame["overdue_days"].notna()
    ] if "overdue_days" in frame.columns else pd.DataFrame()

    overdue = int(
        (valid_overdue["overdue_days"] > 0).sum()
    ) if not valid_overdue.empty else 0

    return {
        "Total tasks": total,
        "Completed tasks": completed,
        "Rejected tasks": rejected,
        "Open tasks": open_tasks,
        "Work in progress tasks": wip,
        "Completion rate": format_percent(
            completed,
            total,
        ),
        "On-time completion rate": format_percent(
            on_time,
            len(valid_on_time),
        ),
        "Open overdue rate": format_percent(
            overdue,
            len(valid_overdue),
        ),
        "Mean execution business hours": (
            frame["execution_business_hours"].mean()
            if "execution_business_hours" in frame.columns
            else None
        ),
        "Median execution business hours": (
            frame["execution_business_hours"].median()
            if "execution_business_hours" in frame.columns
            else None
        ),
    }


def add_individual_profile(
    document: Document,
    assignee: str,
    frame: pd.DataFrame,
) -> None:
    add_heading(
        document,
        f"Individual Achievement Profile: {assignee}",
        level=2,
    )

    completed = frame[
        frame["is_completed"] == True
    ] if "is_completed" in frame.columns else pd.DataFrame()

    valid_on_time = completed[
        completed["on_time_completion"].notna()
    ] if not completed.empty and "on_time_completion" in completed.columns else pd.DataFrame()

    on_time_count = int(
        valid_on_time["on_time_completion"].sum()
    ) if not valid_on_time.empty else 0

    rework_valid = frame[
        frame["rework_count"].notna()
    ] if "rework_count" in frame.columns else pd.DataFrame()

    rework_total = int(
        rework_valid["rework_count"].sum()
    ) if not rework_valid.empty else 0

    profile_values = {
        "Tasks assigned at evaluation cutoff": len(frame),
        "Completed tasks": len(completed),
        "Completion rate": format_percent(
            len(completed),
            len(frame),
        ),
        "On-time completed tasks": on_time_count,
        "On-time completion rate": format_percent(
            on_time_count,
            len(valid_on_time),
        ),
        "Open tasks": int(
            frame["is_open"].sum()
        ) if "is_open" in frame.columns else 0,
        "Open overdue tasks": int(
            (
                frame["overdue_days"] > 0
            ).sum()
        ) if "overdue_days" in frame.columns else 0,
        "Mean execution business hours": (
            completed["execution_business_hours"].mean()
            if not completed.empty
            and "execution_business_hours" in completed.columns
            else None
        ),
        "Total recorded rework transitions": rework_total,
    }

    add_key_value_table(
        document,
        profile_values,
    )

    if completed.empty:
        add_paragraph(
            document,
            "No completed tasks were verified for this assignee "
            "within the selected evaluation scope.",
        )
        return

    issue_types = []

    if "issue_type" in completed.columns:
        issue_types = sorted(
            completed["issue_type"]
            .dropna()
            .astype(str)
            .unique()
            .tolist()
        )

    if issue_types:
        add_paragraph(
            document,
            "Completed issue types: "
            + ", ".join(issue_types),
        )

    labels: set[str] = set()

    if "labels" in completed.columns:
        for value in completed["labels"]:
            labels.update(parse_labels(value))

    if labels:
        add_paragraph(
            document,
            "Labels represented in completed work: "
            + ", ".join(sorted(labels)),
        )

    add_paragraph(
        document,
        (
            f"{assignee} has {len(completed)} verified completed task(s) "
            f"within the evaluation scope. "
            f"The profile describes recorded task outcomes and timing; "
            f"it does not infer effort, quality, or business impact "
            f"from duration alone."
        ),
    )

    add_task_table(
        document,
        completed,
    )


def add_bottlenecks_section(
    document: Document,
    frame: pd.DataFrame,
) -> None:
    add_heading(
        document,
        "Bottlenecks and Exceptions",
        level=1,
    )

    if "status_at_cutoff" in frame.columns:
        status_counts = (
            frame["status_at_cutoff"]
            .fillna("Unavailable")
            .value_counts()
        )

        if not status_counts.empty:
            add_paragraph(
                document,
                "Task distribution by current status:",
            )

            for status, count in status_counts.items():
                add_paragraph(
                    document,
                    f"{status}: {count} task(s)",
                )

    if "rework_count" in frame.columns:
        rework_tasks = frame[
            frame["rework_count"].fillna(0) > 0
        ]

        add_paragraph(
            document,
            (
                f"Tasks with at least one verified rework transition: "
                f"{len(rework_tasks)}."
            ),
        )

    if "overdue_days" in frame.columns:
        overdue_tasks = frame[
            frame["overdue_days"].fillna(0) > 0
        ]

        add_paragraph(
            document,
            (
                f"Tasks with overdue days greater than zero: "
                f"{len(overdue_tasks)}."
            ),
        )

    add_paragraph(
        document,
        (
            "These observations identify review candidates. "
            "They should be interpreted with task complexity, dependencies, "
            "assignment history, and data completeness."
        ),
    )


def add_recommendations(
    document: Document,
    frame: pd.DataFrame,
) -> None:
    add_heading(
        document,
        "Conclusions and Recommendations",
        level=1,
    )

    recommendations = [
        (
            "Review open tasks with overdue days greater than zero "
            "and confirm whether the due date, dependency, or scope changed."
        ),
        (
            "Review repeated In Review to In Progress transitions "
            "to identify requirements, testing, or approval issues."
        ),
        (
            "Use individual achievement profiles for evidence-based "
            "discussion of completed work and delivery reliability."
        ),
        (
            "Keep elapsed duration and business-hours duration visible "
            "together so management can distinguish calendar delay "
            "from time inside the standard work calendar."
        ),
        (
            "Collect complete Jira worklogs when actual recorded effort "
            "and work outside standard hours must be evaluated."
        ),
    ]

    for recommendation in recommendations:
        paragraph = document.add_paragraph(
            style="List Bullet"
        )
        paragraph.add_run(recommendation)


def build_report(
    task_metrics: pd.DataFrame,
    output_path: Path,
    *,
    title: str,
    evaluation_cutoff: str,
    timezone_name: str,
    work_days: str,
    work_window: str,
) -> None:
    document = Document()
    set_document_style(document)

    add_title(
        document,
        title,
    )

    metadata = {
        "Evaluation cutoff": evaluation_cutoff,
        "Timezone": timezone_name,
        "Working days": work_days,
        "Standard work window": work_window,
        "Report language": "English",
        "Primary report focus": "Individual achievements",
    }

    add_key_value_table(
        document,
        metadata,
    )

    add_heading(
        document,
        "Executive Summary",
        level=1,
    )

    overall = calculate_overall_metrics(
        task_metrics,
    )

    add_key_value_table(
        document,
        overall,
    )

    add_paragraph(
        document,
        (
            "This report evaluates recorded Jira task outcomes within "
            "the selected scope. The executive view summarizes delivery "
            "status, timeliness, work in progress, and exceptions. "
            "The individual sections focus on verified achievements "
            "associated with each assignee."
        ),
    )

    add_heading(
        document,
        "Methodology and Work Calendar",
        level=1,
    )

    add_paragraph(
        document,
        (
            "Elapsed durations measure the complete time between verified "
            "timestamps. Business-hours durations count only time within "
            "the configured Syria work calendar: Sunday through Thursday, "
            "09:00 to 17:00, Asia/Damascus. Friday and Saturday are excluded."
        ),
    )

    add_paragraph(
        document,
        (
            "Business-hours duration describes process time inside the "
            "configured calendar. It is not a direct measure of employee "
            "labor. Actual work outside standard hours requires verified "
            "Jira worklog or time-tracking records."
        ),
    )

    add_heading(
        document,
        "Overall Performance Indicators",
        level=1,
    )

    add_key_value_table(
        document,
        overall,
    )

    add_heading(
        document,
        "Individual Achievement Profiles",
        level=1,
    )

    if "assignee_name" in task_metrics.columns:
        assignees = sorted(
            task_metrics["assignee_name"]
            .fillna("Assignee unavailable")
            .astype(str)
            .unique()
            .tolist()
        )

        for assignee in assignees:
            assignee_frame = task_metrics[
                task_metrics["assignee_name"].fillna(
                    "Assignee unavailable"
                ).astype(str) == assignee
            ]

            add_individual_profile(
                document,
                assignee,
                assignee_frame,
            )
    else:
        add_paragraph(
            document,
            "Assignee information was unavailable.",
        )

    add_heading(
        document,
        "Per-Task Evaluation",
        level=1,
    )

    add_task_table(
        document,
        task_metrics,
    )

    add_bottlenecks_section(
        document,
        task_metrics,
    )

    add_recommendations(
        document,
        task_metrics,
    )

    add_heading(
        document,
        "Data Quality and Limitations",
        level=1,
    )

    limitations = [
        "Missing dates make affected duration metrics unavailable.",
        "Incomplete Jira history disables metrics that require verified transitions.",
        "A task outcome is not automatically proof of one person's total contribution.",
        "Task duration alone does not establish effort, quality, or impact.",
        "Labels can overlap; label-group counts should not be added to calculate overall task counts.",
        "Outside-hours activity is reported only when supported by verified work records.",
    ]

    for limitation in limitations:
        paragraph = document.add_paragraph(
            style="List Bullet"
        )
        paragraph.add_run(limitation)

    output_path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    document.save(output_path)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Build an individual-achievement Jira Word report."
    )

    parser.add_argument(
        "--metrics-workbook",
        required=True,
        help="Workbook containing the task_metrics sheet.",
    )

    parser.add_argument(
        "--output",
        required=True,
        help="Output DOCX path.",
    )

    parser.add_argument(
        "--title",
        default="Jira Task Performance Evaluation Report",
        help="Report title.",
    )

    parser.add_argument(
        "--evaluation-cutoff",
        required=True,
        help="Evaluation cutoff shown in the report.",
    )

    parser.add_argument(
        "--timezone",
        default="Asia/Damascus",
        help="Report timezone.",
    )

    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()

    task_metrics = pd.read_excel(
        args.metrics_workbook,
        sheet_name="task_metrics",
        dtype=object,
    )

    cutoff = args.evaluation_cutoff

    build_report(
        task_metrics,
        Path(args.output),
        title=args.title,
        evaluation_cutoff=cutoff,
        timezone_name=args.timezone,
        work_days="Sunday-Thursday",
        work_window="09:00-17:00",
    )

    print(
        f"Word report written to: {args.output}"
    )

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
