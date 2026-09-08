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


def build_report(
    task_metrics, output_path, *, title, evaluation_cutoff,
    timezone_name, work_days, work_window, process_data=None,
):
    from process_analysis import process_tables
    from metrics_engine import aggregate
    from docx.oxml import OxmlElement
    tables = process_data if process_data is not None else process_tables(task_metrics)
    document = Document()
    set_document_style(document)
    document.styles["Normal"].paragraph_format.space_after = Pt(6)
    add_title(document, title)
    add_heading(document, "Process Context and Scope")
    add_key_value_table(document, dict(tables["process_context"].itertuples(index=False, name=None)))
    add_heading(document, "Overall Performance Indicators")
    row = tables["overall_summary"].iloc[0]
    names = {"Total tasks": "total_tasks", "Completed tasks": "completed_tasks",
             "Rejected tasks": "rejected_tasks", "Open tasks": "open_tasks", "WIP tasks": "wip_tasks",
             "Status unavailable": "unknown_status_tasks", "Complete histories": "history_complete_tasks",
             "Histories excluded": "history_excluded_tasks", "Reviewed eligible tasks": "reviewed_valid_tasks"}
    add_key_value_table(document, {name: int(row[key]) for name, key in names.items()})
    for label, count, denom in [
        ("Completion", "completed_tasks", "total_tasks"),
        ("On-time completion", "on_time_tasks", "on_time_valid_tasks"),
        ("Open overdue", "overdue_open_tasks", "overdue_valid_tasks"),
        ("Rework", "tasks_with_rework", "reviewed_valid_tasks"),
        ("Replanning", "tasks_with_replanning", "reviewed_valid_tasks"),
        ("Re-evaluation", "tasks_with_re_evaluation", "reviewed_valid_tasks"),
        ("Any review exception", "review_exception_tasks", "reviewed_valid_tasks")]:
        add_paragraph(document, f"{label}: {format_percent(int(row[count]), int(row[denom]))} ({int(row[count])}/{int(row[denom])} tasks).")
    add_key_value_table(document, {kind.replace("_", " ").title() + " events": row["total_" + kind + "_count"]
                                  for kind in ["rework", "replanning", "re_evaluation"]})
    add_heading(document, "Process Timing")
    for metric in ["execution", "lead_time", "time_to_start"]:
        add_paragraph(document, f"{metric.replace('_', ' ').title()}: mean {format_value(row['mean_' + metric + '_elapsed_hours'])} elapsed hours / {format_value(row['mean_' + metric + '_business_hours'])} business hours; valid tasks: {int(row[metric + '_elapsed_hours_valid_tasks'])}.")
    add_paragraph(document, "Business hours are process residence within the configured calendar, not recorded employee effort. Durations include waiting and repeated visits. Date adherence uses uploaded schedule values, not a reconstructed historical baseline.")
    add_heading(document, "Stage Residence and Open Work")
    for stage in tables["stage_summary"].to_dict("records"):
        add_heading(document, str(stage["status"]), 2)
        add_key_value_table(document, {"Tasks visited": stage["tasks_visited"],
            "Mean elapsed / business hours": f"{stage['elapsed_mean_hours']:.2f} / {stage['business_mean_hours']:.2f}",
            "Median elapsed / business hours": f"{stage['elapsed_median_hours']:.2f} / {stage['business_median_hours']:.2f}",
            "Total elapsed / business hours": f"{stage['elapsed_total_hours']:.2f} / {stage['business_total_hours']:.2f}",
            "Open tasks currently here": stage["open_tasks_currently_here"]})
    add_paragraph(document, "Stage statistics sum all visits per task through cutoff, including unfinished visits. Done and Rejected residence is excluded. High residence identifies a review candidate; it does not establish a cause.")
    add_heading(document, "Bottlenecks, Exceptions and Follow-up")
    findings = tables["process_findings"]
    if findings.empty:
        add_paragraph(document, "No supported exception findings are available in this scope.")
    for finding in findings.to_dict("records"):
        add_paragraph(document, f"{finding['issue_key']}: {finding['observation']} {finding['follow_up']}")
    add_heading(document, "Assignment Distribution")
    add_paragraph(document, "Groups use the uploaded assignee snapshot. Unassigned is a separate task group, not an individual. These counts do not verify contribution or ownership at completion.")
    for group in aggregate(task_metrics, ["assignee_name"]).to_dict("records"):
        add_paragraph(document, f"{group['assignee_name']}: {group['total_tasks']} tasks; {group['completed_tasks']} completed, {group['open_tasks']} open, {group['rejected_tasks']} rejected, {group['unknown_status_tasks']} status unavailable.")
    add_heading(document, "Per-Task Evaluation")
    for task in task_metrics.to_dict("records"):
        add_heading(document, f"{task['issue_key']} — {task['task_name']}", 2)
        add_key_value_table(document, {
            "Status at cutoff": task["status_at_cutoff"],
            "Actual start / completion (UTC)": f"{format_value(task['actual_start_at'])} / {format_value(task['completed_at'])}",
            "Due date (uploaded snapshot)": task["due_date"],
            "On-time completion": task["on_time_completion"],
            "Execution elapsed / business hours": f"{format_value(task['execution_elapsed_hours'])} / {format_value(task['execution_business_hours'])}",
            "Open task age elapsed / business hours": f"{format_value(task['task_age_elapsed_hours'])} / {format_value(task['task_age_business_hours'])}",
            "Current status age elapsed / business hours": f"{format_value(task['current_status_age_elapsed_hours'])} / {format_value(task['current_status_age_business_hours'])}",
            "Overdue days": task["overdue_days"],
            "Rework / replanning / re-evaluation events": " / ".join(format_value(task[k]) for k in ["rework_count", "replanning_count", "re_evaluation_count"]),
            "History complete": task["history_complete"],
        })
    add_heading(document, "Data Quality")
    if tables["data_quality"].empty:
        add_paragraph(document, "No data-quality findings were recorded by these validation checks.")
    for item in tables["data_quality"].to_dict("records"):
        add_paragraph(document, f"{item['issue_key']}: {item['finding']}")
    add_heading(document, "Metric Definitions and Limitations")
    for metric, definition in tables["metric_definitions"].itertuples(index=False, name=None):
        add_paragraph(document, f"{metric}: {definition}")
    add_paragraph(document, "All percentages use a 0–100 scale. A zero denominator is unavailable. The Excel export includes the transition audit trail and detailed stage statistics. Simulation data cannot establish long-term employee performance.")
    # Keep rows intact; narrow two-column tables avoid the old nine-column overflow.
    for table in document.tables:
        for row in table.rows:
            prop = row._tr.get_or_add_trPr()
            prop.append(OxmlElement("w:cantSplit"))
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    document.save(output_path)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Build a Jira process-performance Word report."
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

    workbook = pd.read_excel(args.metrics_workbook, sheet_name=None)
    task_metrics = workbook["task_metrics"]
    required = {"process_context", "overall_summary", "stage_summary", "process_findings", "data_quality", "metric_definitions"}
    process_data = {name: workbook[name] for name in required} if required.issubset(workbook) else None

    cutoff = args.evaluation_cutoff

    build_report(
        task_metrics,
        Path(args.output),
        title=args.title,
        evaluation_cutoff=cutoff,
        timezone_name=args.timezone,
        work_days="Sunday-Thursday",
        work_window="09:00-17:00",
        process_data=process_data,
    )

    print(
        f"Word report written to: {args.output}"
    )

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
