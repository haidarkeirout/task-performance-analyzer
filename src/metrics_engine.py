"""
Jira metrics engine.

Calculates:
- Elapsed durations
- Syria business-hours durations
- Lead time and time to start
- Time in each Jira status
- Completion and overdue metrics
- Rework, replanning, and re-evaluation
- Aggregations by assignee, issue type, and labels
"""

from __future__ import annotations

import argparse
import json
import logging
from dataclasses import dataclass
from datetime import date, datetime, time, timedelta
from pathlib import Path
from typing import Any

import pandas as pd
from zoneinfo import ZoneInfo


LOGGER = logging.getLogger(__name__)


@dataclass(frozen=True)
class WorkCalendar:
    timezone_name: str = "Asia/Damascus"
    working_days: frozenset[int] = frozenset({0, 1, 2, 3, 6})
    start_hour: int = 9
    end_hour: int = 17
    holidays: frozenset[date] = frozenset()

    @property
    def timezone(self) -> ZoneInfo:
        return ZoneInfo(self.timezone_name)

    def is_working_day(self, current_date: date) -> bool:
        return (
            current_date.weekday() in self.working_days
            and current_date not in self.holidays
        )


def parse_timestamp(
    value: Any,
    *,
    timezone_name: str,
) -> pd.Timestamp | None:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None

    if isinstance(value, str) and not value.strip():
        return None

    try:
        timestamp = pd.Timestamp(value)

        if timestamp.tzinfo is None:
            timestamp = timestamp.tz_localize(ZoneInfo(timezone_name))
        else:
            timestamp = timestamp.tz_convert("UTC")

        return timestamp.tz_convert("UTC")

    except Exception:
        return None


def parse_date_value(value: Any) -> date | None:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None

    if isinstance(value, date) and not isinstance(value, datetime):
        return value

    try:
        return pd.Timestamp(value).date()
    except Exception:
        return None


def elapsed_hours(
    start: pd.Timestamp | None,
    end: pd.Timestamp | None,
) -> float | None:
    if start is None or end is None or end < start:
        return None

    return (end - start).total_seconds() / 3600.0


def localize_window(
    current_date: date,
    current_time: time,
    calendar: WorkCalendar,
) -> pd.Timestamp:
    timestamp = pd.Timestamp(
        datetime.combine(current_date, current_time)
    )
    return timestamp.tz_localize(calendar.timezone).tz_convert("UTC")


def business_hours_between(
    start: pd.Timestamp | None,
    end: pd.Timestamp | None,
    calendar: WorkCalendar,
) -> float | None:
    """
    Return hours that intersect the configured business calendar.

    Calendar:
        Sunday through Thursday
        09:00 through 17:00
        Asia/Damascus
        Friday and Saturday excluded
    """
    if start is None or end is None or end < start:
        return None

    if end == start:
        return 0.0

    start = start.tz_convert("UTC")
    end = end.tz_convert("UTC")

    start_local = start.tz_convert(calendar.timezone)
    end_local = end.tz_convert(calendar.timezone)

    current_date = start_local.date()
    last_date = end_local.date()
    total_seconds = 0.0

    while current_date <= last_date:
        if calendar.is_working_day(current_date):
            window_start = localize_window(
                current_date,
                time(calendar.start_hour, 0),
                calendar,
            )
            window_end = localize_window(
                current_date,
                time(calendar.end_hour, 0),
                calendar,
            )

            overlap_start = max(start, window_start)
            overlap_end = min(end, window_end)

            if overlap_end > overlap_start:
                total_seconds += (
                    overlap_end - overlap_start
                ).total_seconds()

        current_date += timedelta(days=1)

    return total_seconds / 3600.0


def load_json(path: Path) -> Any:
    with path.open("r", encoding="utf-8") as file:
        return json.load(file)


def load_calendar(
    metrics_config: dict[str, Any],
    *,
    holiday_values: list[str] | None = None,
) -> WorkCalendar:
    config = metrics_config.get("work_calendar", {})

    configured_holidays = config.get("holiday_dates", [])
    supplied_holidays = holiday_values or []

    holidays: set[date] = set()

    for value in [*configured_holidays, *supplied_holidays]:
        parsed = parse_date_value(value)
        if parsed:
            holidays.add(parsed)

    working_days_by_name = {
        "Monday": 0,
        "Tuesday": 1,
        "Wednesday": 2,
        "Thursday": 3,
        "Friday": 4,
        "Saturday": 5,
        "Sunday": 6,
    }

    working_days = frozenset(
        working_days_by_name[name]
        for name in config.get(
            "working_days",
            [
                "Sunday",
                "Monday",
                "Tuesday",
                "Wednesday",
                "Thursday",
            ],
        )
        if name in working_days_by_name
    )

    windows = config.get(
        "daily_windows",
        [{"start": "09:00", "end": "17:00"}],
    )

    if not windows:
        raise ValueError("No work-calendar window was configured.")

    start_text = windows[0]["start"]
    end_text = windows[0]["end"]

    start_hour, start_minute = map(int, start_text.split(":"))
    end_hour, end_minute = map(int, end_text.split(":"))

    if start_minute != 0 or end_minute != 0:
        raise ValueError(
            "This version expects work-calendar windows on full hours."
        )

    return WorkCalendar(
        timezone_name=config.get(
            "timezone",
            "Asia/Damascus",
        ),
        working_days=working_days,
        start_hour=start_hour,
        end_hour=end_hour,
        holidays=frozenset(holidays),
    )


def normalize_labels(value: Any) -> list[str]:
    if value is None or (isinstance(value, float) and pd.isna(value):
        return []

    if isinstance(value, list):
        values = value
    else:
        text = str(value).strip()

        if not text:
            return []

        try:
            parsed = json.loads(text)
            values = parsed if isinstance(parsed, list) else [text]
        except json.JSONDecodeError:
            values = [text]

    result: list[str] = []

    for item in values:
        text = str(item).strip()
        if text and text not in result:
            result.append(text)

    return result


def get_status_events(
    history: dict[str, Any] | None,
    *,
    cutoff: pd.Timestamp,
    timezone_name: str,
) -> list[dict[str, Any]]:
    if not history:
        return []

    events: list[dict[str, Any]] = []

    for event in history.get("status_events", []):
        changed_at = parse_timestamp(
            event.get("changed_at"),
            timezone_name=timezone_name,
        )

        if changed_at is None or changed_at > cutoff:
            continue

        events.append(
            {
                **event,
                "_changed_at": changed_at,
            }
        )

    events.sort(key=lambda item: item["_changed_at"])
    return events


def build_status_intervals(
    *,
    created_at: pd.Timestamp | None,
    events: list[dict[str, Any]],
    cutoff: pd.Timestamp,
) -> tuple[list[dict[str, Any]], bool]:
    """
    Build status intervals from created_at through the evaluation cutoff.

    The first interval is accepted only when the first transition supplies
    a verified from_status. This avoids inventing an initial Jira status.
    """
    if created_at is None or not events:
        return [], False

    first_event = events[0]
    current_status = first_event.get("from_status")

    if not current_status:
        return [], False

    cursor = created_at
    intervals: list[dict[str, Any]] = []

    for event in events:
        changed_at = event["_changed_at"]
        from_status = event.get("from_status")
        to_status = event.get("to_status")

        if changed_at < cursor:
            return [], False

        if from_status and current_status != from_status:
            return [], False

        if not to_status:
            return [], False

        intervals.append(
            {
                "status": current_status,
                "start": cursor,
                "end": changed_at,
            }
        )

        current_status = to_status
        cursor = changed_at

    if cutoff >= cursor:
        intervals.append(
            {
                "status": current_status,
                "start": cursor,
                "end": cutoff,
            }
        )

    return intervals, True


def summarize_status_time(
    intervals: list[dict[str, Any]],
    calendar: WorkCalendar,
) -> dict[str, dict[str, float]]:
    summary: dict[str, dict[str, float]] = {}

    for interval in intervals:
        status = str(interval["status"])
        start = interval["start"]
        end = interval["end"]

        elapsed = elapsed_hours(start, end)
        business = business_hours_between(start, end, calendar)

        if elapsed is None or business is None:
            continue

        if status not in summary:
            summary[status] = {
                "elapsed_hours": 0.0,
                "business_hours": 0.0,
            }

        summary[status]["elapsed_hours"] += elapsed
        summary[status]["business_hours"] += business

    return summary


def first_status_entry(
    events: list[dict[str, Any]],
    target_status: str,
) -> pd.Timestamp | None:
    for event in events:
        if event.get("to_status") == target_status:
            return event["_changed_at"]

    return None


def latest_status_entry(
    events: list[dict[str, Any]],
    target_status: str,
) -> pd.Timestamp | None:
    matching = [
        event["_changed_at"]
        for event in events
        if event.get("to_status") == target_status
    ]

    return max(matching) if matching else None


def count_transition(
    events: list[dict[str, Any]],
    *,
    from_status: str,
    to_status: str,
) -> int:
    return sum(
        1
        for event in events
        if event.get("from_status") == from_status
        and event.get("to_status") == to_status
    )


def calculate_task_metrics(
    task: dict[str, Any],
    history: dict[str, Any] | None,
    *,
    cutoff: pd.Timestamp,
    calendar: WorkCalendar,
) -> dict[str, Any]:
    issue_key = str(task.get("issue_key") or "").strip()

    created_at = parse_timestamp(
        task.get("created_at"),
        timezone_name=calendar.timezone_name,
    )

    due_date = parse_date_value(task.get("due_date"))
    planned_start_date = parse_date_value(
        task.get("planned_start_date")
    )

    current_status = task.get("current_status")
    current_status = (
        str(current_status).strip()
        if current_status is not None
        else None
    )

    labels = normalize_labels(task.get("labels"))

    events = get_status_events(
        history,
        cutoff=cutoff,
        timezone_name=calendar.timezone_name,
    )

    intervals, history_complete = build_status_intervals(
        created_at=created_at,
        events=events,
        cutoff=cutoff,
    )

    status_times = summarize_status_time(intervals, calendar)

    actual_start_at = first_status_entry(
        events,
        "In Progress",
    )

    latest_done_at = latest_status_entry(
        events,
        "Done",
    )

    final_status = (
        intervals[-1]["status"]
        if intervals
        else current_status
    )

    completed_at = (
        latest_done_at
        if final_status == "Done"
        else None
    )

    is_completed = final_status == "Done"
    is_rejected = final_status == "Rejected"
    is_open = not is_completed and not is_rejected
    is_wip = final_status in {"In Progress", "In Review"}

    execution_elapsed = elapsed_hours(
        actual_start_at,
        completed_at,
    )

    execution_business = business_hours_between(
        actual_start_at,
        completed_at,
        calendar,
    )

    lead_elapsed = elapsed_hours(
        created_at,
        completed_at,
    )

    lead_business = business_hours_between(
        created_at,
        completed_at,
        calendar,
    )

    start_elapsed = elapsed_hours(
        created_at,
        actual_start_at,
    )

    start_business = business_hours_between(
        created_at,
        actual_start_at,
        calendar,
    )

    age_elapsed = elapsed_hours(
        created_at,
        cutoff,
    )

    age_business = business_hours_between(
        created_at,
        cutoff,
        calendar,
    )

    status_start = None

    if intervals:
        status_start = intervals[-1]["start"]

    current_status_age_elapsed = elapsed_hours(
        status_start,
        cutoff,
    )

    current_status_age_business = business_hours_between(
        status_start,
        cutoff,
        calendar,
    )

    on_time: bool | None = None
    schedule_variance: int | None = None
    overdue_days: int | None = None

    if due_date is not None:
        cutoff_local_date = cutoff.tz_convert(
            calendar.timezone
        ).date()

        if completed_at is not None:
            completed_local_date = completed_at.tz_convert(
                calendar.timezone
            ).date()

            schedule_variance = (
                completed_local_date - due_date
            ).days

            on_time = schedule_variance <= 0
            overdue_days = max(0, schedule_variance)

        elif is_open:
            overdue_days = max(
                0,
                (cutoff_local_date - due_date).days,
            )

    start_schedule_variance: int | None = None

    if planned_start_date is not None and actual_start_at is not None:
        actual_start_local_date = actual_start_at.tz_convert(
            calendar.timezone
        ).date()

        start_schedule_variance = (
            actual_start_local_date - planned_start_date
        ).days

    rework_count = count_transition(
        events,
        from_status="In Review",
        to_status="In Progress",
    )

    replanning_count = count_transition(
        events,
        from_status="In Review",
        to_status="To Do",
    )

    reevaluation_count = count_transition(
        events,
        from_status="In Review",
        to_status="In Triage",
    )

    reopen_count = sum(
        1
        for event in events
        if event.get("from_status") == "Done"
        and event.get("to_status")
        not in {"Done", "Rejected"}
    )

    if not history_complete:
        history_note = "Status history incomplete or inconsistent."
    elif not events:
        history_note = "No usable status history was supplied."
    else:
        history_note = None

    result: dict[str, Any] = {
        "issue_key": issue_key,
        "task_name": task.get("task_name"),
        "issue_type": task.get("issue_type")
        or "Issue type unavailable",
        "labels": labels,
        "assignee_id": task.get("assignee_id"),
        "assignee_name": task.get("assignee_name")
        or "Unassigned",
        "priority": task.get("priority"),
        "status_at_cutoff": final_status,
        "created_at": created_at.isoformat()
        if created_at is not None
        else None,
        "actual_start_at": actual_start_at.isoformat()
        if actual_start_at is not None
        else None,
        "completed_at": completed_at.isoformat()
        if completed_at is not None
        else None,
        "due_date": due_date.isoformat()
        if due_date is not None
        else None,
        "planned_start_date": planned_start_date.isoformat()
        if planned_start_date is not None
        else None,
        "is_completed": is_completed,
        "is_rejected": is_rejected,
        "is_open": is_open,
        "is_wip": is_wip,
        "history_complete": history_complete,
        "history_note": history_note,
        "execution_elapsed_hours": execution_elapsed,
        "execution_business_hours": execution_business,
        "lead_time_elapsed_hours": lead_elapsed,
        "lead_time_business_hours": lead_business,
        "time_to_start_elapsed_hours": start_elapsed,
        "time_to_start_business_hours": start_business,
        "task_age_elapsed_hours": age_elapsed
        if is_open
        else None,
        "task_age_business_hours": age_business
        if is_open
        else None,
        "current_status_age_elapsed_hours": (
            current_status_age_elapsed
            if is_open
            else None
        ),
        "current_status_age_business_hours": (
            current_status_age_business
            if is_open
            else None
        ),
        "on_time_completion": on_time,
        "schedule_variance_days": schedule_variance,
        "overdue_days": overdue_days,
        "start_schedule_variance_days": start_schedule_variance,
        "rework_count": rework_count
        if history_complete
        else None,
        "replanning_count": replanning_count
        if history_complete
        else None,
        "re_evaluation_count": reevaluation_count
        if history_complete
        else None,
        "reopen_count": reopen_count
        if history_complete
        else None,
        "time_in_status": status_times,
        "evaluation_cutoff": cutoff.isoformat(),
        "work_calendar_timezone": calendar.timezone_name,
        "work_calendar_days": (
            "Sunday-Thursday"
        ),
        "work_calendar_window": (
            f"{calendar.start_hour:02d}:00-"
            f"{calendar.end_hour:02d}:00"
        ),
    }

    return result


def analyze_tasks(
    tasks_frame: pd.DataFrame,
    histories: dict[str, dict[str, Any]],
    *,
    cutoff: pd.Timestamp,
    calendar: WorkCalendar,
) -> pd.DataFrame:
    records: list[dict[str, Any]] = []

    for _, row in tasks_frame.iterrows():
        task = row.to_dict()
        issue_key = str(task.get("issue_key") or "").strip()

        history = histories.get(issue_key)

        records.append(
            calculate_task_metrics(
                task,
                history,
                cutoff=cutoff,
                calendar=calendar,
            )
        )

    return pd.DataFrame(records)


def safe_mean(series: pd.Series) -> float | None:
    values = pd.to_numeric(series, errors="coerce").dropna()

    if values.empty:
        return None

    return float(values.mean())


def safe_median(series: pd.Series) -> float | None:
    values = pd.to_numeric(series, errors="coerce").dropna()

    if values.empty:
        return None

    return float(values.median())


def aggregate_frame(
    frame: pd.DataFrame,
    *,
    group_columns: list[str],
) -> pd.DataFrame:
    if frame.empty:
        return pd.DataFrame()

    working = frame.copy()

    rows: list[dict[str, Any]] = []

    grouped = working.groupby(
        group_columns,
        dropna=False,
    )

    for group_key, group in grouped:
        if not isinstance(group_key, tuple):
            group_key = (group_key,)

        record = {
            column: value
            for column, value in zip(group_columns, group_key)
        }

        total_tasks = int(group["issue_key"].nunique())
        completed_tasks = int(group["is_completed"].sum())
        rejected_tasks = int(group["is_rejected"].sum())
        open_tasks = int(group["is_open"].sum())
        wip_tasks = int(group["is_wip"].sum())

        valid_on_time = group[
            group["on_time_completion"].notna()
        ]

        valid_overdue = group[
            group["overdue_days"].notna()
        ]

        valid_rework = group[
            group["rework_count"].notna()
        ]

        on_time_count = int(
            valid_on_time["on_time_completion"].sum()
        )

        overdue_count = int(
            (valid_overdue["overdue_days"] > 0).sum()
        )

        rework_task_count = int(
            (valid_rework["rework_count"] > 0).sum()
        )

        record.update(
            {
                "total_tasks": total_tasks,
                "completed_tasks": completed_tasks,
                "rejected_tasks": rejected_tasks,
                "open_tasks": open_tasks,
                "wip_tasks": wip_tasks,
                "completion_rate": (
                    completed_tasks / total_tasks * 100
                    if total_tasks
                    else None
                ),
                "rejection_rate": (
                    rejected_tasks / total_tasks * 100
                    if total_tasks
                    else None
                ),
                "on_time_tasks": on_time_count,
                "on_time_valid_tasks": len(valid_on_time),
                "on_time_completion_rate": (
                    on_time_count / len(valid_on_time) * 100
                    if len(valid_on_time)
                    else None
                ),
                "overdue_open_tasks": overdue_count,
                "overdue_valid_tasks": len(valid_overdue),
                "open_overdue_rate": (
                    overdue_count / len(valid_overdue) * 100
                    if len(valid_overdue)
                    else None
                ),
                "tasks_with_rework": rework_task_count,
                "rework_valid_tasks": len(valid_rework),
                "tasks_with_rework_rate": (
                    rework_task_count / len(valid_rework) * 100
                    if len(valid_rework)
                    else None
                ),
                "mean_execution_business_hours": safe_mean(
                    group["execution_business_hours"]
                ),
                "median_execution_business_hours": safe_median(
                    group["execution_business_hours"]
                ),
                "mean_lead_time_business_hours": safe_mean(
                    group["lead_time_business_hours"]
                ),
                "median_lead_time_business_hours": safe_median(
                    group["lead_time_business_hours"]
                ),
                "mean_time_in_review_business_hours": None,
                "mean_rework_count": safe_mean(
                    group["rework_count"]
                ),
                "total_rework_count": (
                    int(valid_rework["rework_count"].sum())
                    if len(valid_rework)
                    else None
                ),
                "history_complete_tasks": int(
                    group["history_complete"].sum()
                ),
            }
        )

        review_values: list[float] = []

        for value in group["time_in_status"]:
            if not isinstance(value, dict):
                continue

            review = value.get("In Review", {})
            hours = review.get("business_hours")

            if hours is not None:
                review_values.append(float(hours))

        if review_values:
            record["mean_time_in_review_business_hours"] = sum(
                review_values
            ) / len(review_values)

        rows.append(record)

    return pd.DataFrame(rows)


def build_overall_summary(frame: pd.DataFrame) -> pd.DataFrame:
    overall = aggregate_frame(
        frame,
        group_columns=[],
    )

    if overall.empty:
        return overall

    return overall


def flatten_for_excel(frame: pd.DataFrame) -> pd.DataFrame:
    output = frame.copy()

    for column in ["labels", "time_in_status"]:
        if column in output.columns:
            output[column] = output[column].apply(
                lambda value: json.dumps(
                    value,
                    ensure_ascii=False,
                )
                if isinstance(value, (list, dict))
                else value
            )

    return output


def write_metrics_output(
    output_path: Path,
    task_metrics: pd.DataFrame,
    overall_summary: pd.DataFrame,
    assignee_summary: pd.DataFrame,
    issue_type_summary: pd.DataFrame,
) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        flatten_for_excel(task_metrics).to_excel(
            writer,
            index=False,
            sheet_name="task_metrics",
        )

        overall_summary.to_excel(
            writer,
            index=False,
            sheet_name="overall_summary",
        )

        assignee_summary.to_excel(
            writer,
            index=False,
            sheet_name="by_assignee",
        )

        issue_type_summary.to_excel(
            writer,
            index=False,
            sheet_name="by_issue_type",
        )


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Calculate Jira task-performance metrics."
    )

    parser.add_argument(
        "--tasks",
        required=True,
        help="Normalized workbook created by jira_processor.py.",
    )

    parser.add_argument(
        "--history",
        required=True,
        help="JSON history created by jira_client.py.",
    )

    parser.add_argument(
        "--metrics-config",
        default="configs/jira_metrics_config.json",
        help="Path to jira_metrics_config.json.",
    )

    parser.add_argument(
        "--cutoff",
        required=True,
        help=(
            "Evaluation cutoff, for example "
            "2026-09-06T17:00:00+03:00."
        ),
    )

    parser.add_argument(
        "--output",
        required=True,
        help="Path for the metrics workbook.",
    )

    parser.add_argument(
        "--holiday",
        action="append",
        default=[],
        help="Optional holiday date in YYYY-MM-DD format.",
    )

    return parser


def main() -> int:
    logging.basicConfig(
        level=logging.INFO,
        format="%(levelname)s: %(message)s",
    )

    parser = build_parser()
    args = parser.parse_args()

    try:
        metrics_config = load_json(
            Path(args.metrics_config)
        )

        calendar = load_calendar(
            metrics_config,
            holiday_values=args.holiday,
        )

        cutoff = parse_timestamp(
            args.cutoff,
            timezone_name=calendar.timezone_name,
        )

        if cutoff is None:
            raise ValueError(
                "The evaluation cutoff is not a valid timestamp."
            )

        tasks_frame = pd.read_excel(
            args.tasks,
            sheet_name="normalized_tasks",
            dtype=object,
        )

        histories = load_json(Path(args.history))

        task_metrics = analyze_tasks(
            tasks_frame,
            histories,
            cutoff=cutoff,
            calendar=calendar,
        )

        overall_summary = build_overall_summary(
            task_metrics
        )

        assignee_summary = aggregate_frame(
            task_metrics,
            group_columns=["assignee_name"],
        )

        issue_type_summary = aggregate_frame(
            task_metrics,
            group_columns=["issue_type"],
        )

        write_metrics_output(
            Path(args.output),
            task_metrics,
            overall_summary,
            assignee_summary,
            issue_type_summary,
        )

        print(f"Analyzed tasks: {len(task_metrics)}")
        print(f"Output written to: {args.output}")

        return 0

    except Exception as exc:
        LOGGER.exception("Metrics calculation failed: %s", exc)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
