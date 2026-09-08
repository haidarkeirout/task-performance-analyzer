"""Shared process views for the dashboard and both downloads.

No transition log is declared complete merely because rows are present.
Workbook coverage must be explicitly confirmed; API coverage is recorded by the client.
"""
from __future__ import annotations
import json
from datetime import datetime, time
from io import BytesIO
import pandas as pd
from metrics_engine import aggregate, get_status_events, parse_timestamp

EVENT_COLUMNS = ["issue_key", "task_name", "from_status", "to_status", "event_type", "changed_at", "author_name", "source", "included_in_metrics"]
STAGE_COLUMNS = ["status", "tasks_visited", "elapsed_total_hours", "elapsed_mean_hours", "elapsed_median_hours", "business_total_hours", "business_mean_hours", "business_median_hours", "open_tasks_currently_here"]
RATE_DEFINITIONS = {
    "Completion rate": "Completed tasks / all tasks created by cutoff. Unknown status tasks remain in total and are disclosed.",
    "On-time completion rate": "On-time completed tasks / completed tasks with known completion and due date. Due dates are the uploaded schedule snapshot, not a reconstructed historical baseline.",
    "Open overdue rate": "Overdue open tasks / open tasks with a known due date. Completed and rejected tasks are excluded.",
    "Rework rate": "Distinct reviewed tasks with In Review -> In Progress / distinct reviewed tasks with complete history.",
    "Replanning rate": "Distinct reviewed tasks with In Review -> To Do / distinct reviewed tasks with complete history.",
    "Re-evaluation rate": "Distinct reviewed tasks with In Review -> In Triage / distinct reviewed tasks with complete history.",
    "Review exception rate": "Union of tasks with any of the three review-return types / distinct reviewed tasks with complete history; each task counted once.",
    "Stage durations": "Visits to each status are summed per task through cutoff; mean/median are over tasks visiting that status. Open visits are included and unfinished. Done/Rejected residence is excluded from stage bottleneck summaries.",
    "Execution duration": "First verified entry into In Progress to the latest Done entry when Done at cutoff, including waiting and repeat cycles; not labor hours.",
    "Coverage": "Missing or inconsistent history disables transition-derived metrics. Zero events is valid only with confirmed complete coverage and known initial status.",
    "Time scope": "Cumulative history from task creation through the explicit cutoff; not a period-only event filter. Tasks created later are excluded.",
}

DEADLINE_COLUMNS = [
    "due_status",
    "task_count",
    "share_of_open_known_tasks",
]

WEEKLY_FLOW_COLUMNS = [
    "week_start",
    "tasks_opened",
    "tasks_completed",
    "net_flow",
    "cumulative_net_flow",
]


def workbook_context(workbook):
    frame = workbook.get("Process_Context")
    if frame is None or frame.shape[1] < 2:
        return {}
    return {str(row.iloc[0]).strip(): str(row.iloc[1]) for _, row in frame.iterrows()
            if pd.notna(row.iloc[0]) and pd.notna(row.iloc[1])}


def workbook_histories(workbook, tasks, source_timezone, through_text, confirmed):
    frame = workbook.get("Workflow_Events")
    if frame is None:
        raise ValueError("This workbook has no Workflow_Events sheet. Select Jira API or History JSON.")
    needed = {"Work Item ID", "From Status", "To Status", "Transition Date", "Transition Time"}
    if not needed.issubset(frame.columns):
        raise ValueError("Workflow_Events is missing: " + ", ".join(sorted(needed - set(frame.columns))))
    through = parse_timestamp(through_text, source_timezone)
    if through is None:
        raise ValueError("Enter a valid history coverage end with time and timezone.")
    result = {}
    for row in tasks.to_dict("records"):
        result[str(row["issue_key"])] = {
            "snapshot": {"current_status": row.get("current_status")},
            "initial_status": row.get("current_status"), "status_events": [],
            "history_complete": bool(confirmed), "history_through": through.isoformat(),
            "source": "workbook_declared_history", "validation_notes": [],
        }
    seen = set()
    for index, row in frame.iterrows():
        if pd.isna(row["Work Item ID"]):
            continue
        key = str(row["Work Item ID"]).strip()
        if key not in result:
            continue
        h = result[key]
        # The source workbook lists never-transitioned tasks as blank event rows.
        # This is a task placeholder, not a malformed transition.
        if all(pd.isna(row[name]) for name in ["From Status", "To Status", "Transition Date", "Transition Time"]):
            continue
        try:
            day = pd.Timestamp(row["Transition Date"])
            clock = row["Transition Time"]
            if not isinstance(clock, time):
                clock = time.fromisoformat(str(clock).strip())
            stamp = parse_timestamp(datetime.combine(day.date(), clock), source_timezone)
            if stamp is None or pd.isna(row["From Status"]) or pd.isna(row["To Status"]):
                raise ValueError("Missing event fields")
            event = {"issue_key": key, "from_status": str(row["From Status"]).strip(),
                     "to_status": str(row["To Status"]).strip(), "changed_at": stamp.isoformat(),
                     "author_name": None if pd.isna(row.get("Performed By")) else str(row.get("Performed By")),
                     "source": "Workflow_Events", "source_row": int(index) + 2}
            identity = (key, event["from_status"], event["to_status"], event["changed_at"])
            if identity in seen:
                h["history_complete"] = False
                h["validation_notes"].append(f"Duplicate transition row {index + 2}; confirm source log.")
            seen.add(identity)
            h["status_events"].append(event)
        except (ValueError, TypeError):
            h["history_complete"] = False
            h["validation_notes"].append(f"Invalid transition row {index + 2}.")
    return result


def classify_event(before, after):
    return {("In Review", "In Progress"): "Rework", ("In Review", "To Do"): "Replanning",
            ("In Review", "In Triage"): "Re-evaluation"}.get((before, after),
            "Completion" if after == "Done" else "Rejection" if after == "Rejected" else
            "Reopen" if before == "Done" else "Normal")


def stage_summary(frame):
    rows = []
    for task in frame.to_dict("records"):
        if not task["history_complete"]:
            continue
        times = task["time_in_status"]
        if isinstance(times, str):
            times = json.loads(times)
        for status, values in times.items():
            if status in {"Done", "Rejected"}:
                continue
            rows.append({"issue_key": task["issue_key"], "status": status,
                         "elapsed": values["elapsed_hours"], "business": values["business_hours"],
                         "current": int(task["is_open"] and task["status_at_cutoff"] == status)})
    if not rows:
        return pd.DataFrame(columns=STAGE_COLUMNS)
    return pd.DataFrame(rows).groupby("status").agg(
        tasks_visited=("issue_key", "nunique"), elapsed_total_hours=("elapsed", "sum"),
        elapsed_mean_hours=("elapsed", "mean"), elapsed_median_hours=("elapsed", "median"),
        business_total_hours=("business", "sum"), business_mean_hours=("business", "mean"),
        business_median_hours=("business", "median"), open_tasks_currently_here=("current", "sum")
    ).reset_index()


def deadline_summary(frame):
    """Classify only open tasks whose status is verified.

    Unknown-status tasks are kept separate so a missing history cannot silently
    become an on-time or overdue task.
    """
    open_known = frame[frame["is_open"].eq(True)].copy()
    categories = [
        "Open overdue",
        "Open within due date",
        "Open without due date",
    ]
    if open_known.empty:
        return pd.DataFrame(
            [{"due_status": name, "task_count": 0,
              "share_of_open_known_tasks": None} for name in categories],
            columns=DEADLINE_COLUMNS,
        )
    overdue = open_known["overdue_days"].notna() & open_known["overdue_days"].gt(0)
    missing_due = open_known["due_date"].isna() | open_known["due_date"].eq("")
    counts = {
        "Open overdue": int(overdue.sum()),
        "Open within due date": int((~overdue & ~missing_due).sum()),
        "Open without due date": int(missing_due.sum()),
    }
    denominator = len(open_known)
    return pd.DataFrame(
        [{"due_status": name, "task_count": count,
          "share_of_open_known_tasks": count / denominator * 100}
         for name, count in counts.items()],
        columns=DEADLINE_COLUMNS,
    )


def weekly_flow_summary(frame):
    """Count task creation and completion events by Monday-starting week.

    The chart is intentionally based on task-level created/completed dates,
    not on status-transition rows. Each task can therefore contribute at most
    once to each weekly series. Counts are limited to the evaluation cutoff.
    """
    if frame is None or frame.empty:
        return pd.DataFrame(columns=WEEKLY_FLOW_COLUMNS)

    timezone_name = str(frame.iloc[0].get("work_calendar_timezone") or "UTC")
    try:
        created = pd.to_datetime(frame.get("created_at"), utc=True, errors="coerce")
        completed = pd.to_datetime(frame.get("completed_at"), utc=True, errors="coerce")
        created = created.dt.tz_convert(timezone_name)
        completed = completed.dt.tz_convert(timezone_name)
    except (TypeError, ValueError):
        return pd.DataFrame(columns=WEEKLY_FLOW_COLUMNS)

    cutoff = parse_timestamp(frame.iloc[0].get("evaluation_cutoff"), "UTC")
    if cutoff is not None:
        cutoff_local = cutoff.tz_convert(timezone_name)
        created = created.where(created.le(cutoff_local))
        completed = completed.where(completed.le(cutoff_local))

    def week_starts(values):
        values = values.dropna()
        if values.empty:
            return pd.Series(dtype="datetime64[ns]")
        normalized = values.dt.normalize()
        starts = normalized - pd.to_timedelta(normalized.dt.weekday, unit="D")
        # Excel cannot store timezone-aware datetimes. Keep the local calendar
        # date as a naive timestamp for charts and exports.
        return starts.dt.tz_localize(None)

    created_weeks = week_starts(created)
    completed_weeks = week_starts(completed)
    available = pd.concat([created_weeks, completed_weeks], ignore_index=True).dropna()
    if available.empty:
        return pd.DataFrame(columns=WEEKLY_FLOW_COLUMNS)

    first_week = available.min()
    last_week = available.max()
    weeks = pd.date_range(first_week, last_week, freq="7D")
    result = pd.DataFrame({"week_start": weeks})
    result["tasks_opened"] = result["week_start"].map(created_weeks.value_counts()).fillna(0).astype(int)
    result["tasks_completed"] = result["week_start"].map(completed_weeks.value_counts()).fillna(0).astype(int)
    result["net_flow"] = result["tasks_opened"] - result["tasks_completed"]
    result["cumulative_net_flow"] = result["net_flow"].cumsum()
    return result[WEEKLY_FLOW_COLUMNS]


def process_tables(frame, histories=None, context=None, validation_log=None):
    context = dict(context or frame.attrs.get("process_context", {}))
    histories = histories or {}
    summary = aggregate(frame, [])
    cutoff = parse_timestamp(frame.iloc[0]["evaluation_cutoff"], "UTC")
    events = []
    for task in frame.to_dict("records"):
        for event in get_status_events(histories.get(task["issue_key"]), cutoff, "UTC"):
            events.append({"issue_key": task["issue_key"], "task_name": task["task_name"],
                "from_status": event.get("from_status"), "to_status": event.get("to_status"),
                "event_type": classify_event(event.get("from_status"), event.get("to_status")),
                "changed_at": event["_changed_at"].isoformat(), "author_name": event.get("author_name"),
                "source": event.get("source", "History JSON"), "included_in_metrics": task["history_complete"]})
    stages = stage_summary(frame)
    deadlines = deadline_summary(frame)
    weekly_flow = weekly_flow_summary(frame)
    quality = [{"issue_key": t["issue_key"], "finding": t["history_note"]}
               for t in frame.to_dict("records") if not t["history_complete"]]
    for t in frame.to_dict("records"):
        if not t.get("due_date"):
            quality.append({"issue_key": t["issue_key"], "finding": "Due date missing; schedule metrics unavailable."})
    for item in frame.attrs.get("excluded_tasks", []):
        quality.append({"issue_key": item["issue_key"], "finding": item["reason"]})
    for item in validation_log or []:
        quality.append({"issue_key": item.get("issue_key", "Workbook"), "finding": item.get("message", str(item))})
    for key, history in histories.items():
        for note in history.get("validation_notes", []):
            quality.append({"issue_key": key, "finding": note})
    findings = []
    for t in frame.to_dict("records"):
        for col, label in [("rework_count", "rework"), ("replanning_count", "replanning"), ("re_evaluation_count", "re-evaluation")]:
            if pd.notna(t[col]) and t[col] > 0:
                findings.append({"issue_key": t["issue_key"], "observation": f"{int(t[col])} {label} return(s) before cutoff.",
                                 "follow_up": "Review the transition comments and requirements with the process owner; the cause is not established by the count."})
        if t["is_open"] and pd.notna(t["overdue_days"]) and t["overdue_days"] > 0:
            findings.append({"issue_key": t["issue_key"], "observation": f"Open and overdue by {int(t['overdue_days'])} calendar day(s) against the supplied due date.",
                             "follow_up": "Confirm the due date, dependencies, and next action with the process owner."})
    for stage in stages.to_dict("records"):
        findings.append({"issue_key": "Stage: " + stage["status"],
                         "observation": f"{stage['tasks_visited']} task(s) visited; mean residence {stage['elapsed_mean_hours']:.2f} elapsed hours / {stage['business_mean_hours']:.2f} business hours; {stage['open_tasks_currently_here']} currently here.",
                         "follow_up": "Inspect the task-level stage times; residence is a review candidate, not proof of a bottleneck."})
    for q in quality:
        if q["issue_key"] != "Workbook":
            findings.append({"issue_key": q["issue_key"], "observation": q["finding"], "follow_up": "Complete or correct the source record before interpreting affected metrics."})
    context.update({"Evaluation cutoff used": cutoff.isoformat(),
                    "Metric scope": "Cumulative from task creation through cutoff",
                    "Schedule and assignee basis": "Uploaded snapshot values; historical due dates and ownership are not reconstructed.",
                    "Work calendar": f"{frame.iloc[0]['work_calendar_timezone']}; {frame.iloc[0]['work_calendar_days']}; {frame.iloc[0]['work_calendar_window']}",
                    "Interpretation": "Process residence is not recorded labor, productivity, or individual contribution."})
    open_columns = ["issue_key", "task_name", "assignee_name", "priority", "status_at_cutoff", "created_at", "planned_start_date", "actual_start_at", "start_schedule_variance_days", "due_date", "completed_at", "schedule_variance_days", "overdue_days", "task_age_elapsed_hours", "task_age_business_hours", "current_status_age_elapsed_hours", "current_status_age_business_hours"]
    return {
        "overall_summary": summary,
        "process_context": pd.DataFrame(context.items(), columns=["Field", "Value"]),
        "workflow_events": pd.DataFrame(events, columns=EVENT_COLUMNS),
        "stage_summary": stages,
        "deadline_summary": deadlines,
        "weekly_flow": weekly_flow,
        "overdue_tasks": frame.loc[
            frame["is_open"].eq(True)
            & frame["overdue_days"].notna()
            & frame["overdue_days"].gt(0),
            open_columns,
        ].copy(),
        "late_completed_tasks": frame.loc[
            frame["is_completed"].eq(True)
            & frame["schedule_variance_days"].notna()
            & frame["schedule_variance_days"].gt(0),
            open_columns,
        ].copy(),
        "open_tasks": frame.loc[frame["is_open"].eq(True), open_columns].copy(),
        "data_quality": pd.DataFrame(quality, columns=["issue_key", "finding"]),
        "process_findings": pd.DataFrame(findings, columns=["issue_key", "observation", "follow_up"]),
        "metric_definitions": pd.DataFrame(RATE_DEFINITIONS.items(), columns=["Metric", "Definition"]),
    }


def excel_bytes(frame, tables):
    """Runtime exporter for this existing pandas application (not an authoring tool)."""
    result = frame.copy()
    for col in ["labels", "time_in_status"]:
        result[col] = result[col].map(lambda x: json.dumps(x, ensure_ascii=False) if isinstance(x, (list, dict)) else x)
    output = BytesIO()
    sheets = {"task_metrics": result, **tables,
              "by_assignee": aggregate(frame, ["assignee_name"]), "by_issue_type": aggregate(frame, ["issue_type"])}
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for name, data in sheets.items():
            data.to_excel(writer, sheet_name=name, index=False)
            sheet = writer.sheets[name]
            sheet.freeze_panes = "A2"
            sheet.auto_filter.ref = sheet.dimensions
            # Prevent source text beginning '=' from becoming executable Excel formulas.
            for row in sheet:
                for cell in row:
                    if cell.data_type == "f":
                        cell.data_type = "s"
    return output.getvalue()
