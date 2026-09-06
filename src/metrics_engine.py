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


def load_json(path: Path) -> Any:
    with path.open("r", encoding="utf-8") as file:
        return json.load(file)


def parse_timestamp(
    value: Any,
    timezone_name: str,
) -> pd.Timestamp | None:
    if value is None:
        return None

    if isinstance(value, float) and pd.isna(value):
        return None

    if isinstance(value, str) and not value.strip():
        return None

    try:
        timestamp = pd.Timestamp(value)

        if timestamp.tzinfo is None:
            timestamp = timestamp.tz_localize(
                ZoneInfo(timezone_name)
            )
        else:
            timestamp = timestamp.tz_convert("UTC")

        return timestamp.tz_convert("UTC")

    except Exception:
        return None


def parse_date_value(value: Any) -> date | None:
    if value is None:
        return None

    if isinstance(value, float) and pd.isna(value):
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

    return (end - start).total_seconds() / 3600


def local_window(
    current_date: date,
    hour: int,
    calendar: WorkCalendar,
) -> pd.Timestamp:
    timestamp = pd.Timestamp(
        datetime.combine(
            current_date,
            time(hour, 0),
        )
    )

    return timestamp.tz_localize(
        calendar.timezone
    ).tz_convert("UTC")


def business_hours_between(
    start: pd.Timestamp | None,
    end: pd.Timestamp | None,
    calendar: WorkCalendar,
) -> float | None:
    if start is None or end is None or end < start:
        return None

    if start == end:
        return 0.0

    start = start.tz_convert("UTC")
    end = end.tz_convert("UTC")

    start_local = start.tz_convert(calendar.timezone)
    end_local = end.tz_convert(calendar.timezone)

    current_date = start_local.date()
    last_date = end_local.date()
    seconds = 0.0

    while current_date <= last_date:
        if calendar.is_working_day(current_date):
            window_start = local_window(
                current_date,
                calendar.start_hour,
                calendar,
            )

            window_end = local_window(
                current_date,
                calendar.end_hour,
                calendar,
            )

            overlap_start = max(start, window_start)
            overlap_end = min(end, window_end)

            if overlap_end > overlap_start:
                seconds += (
                    overlap_end - overlap_start
                ).total_seconds()

        current_date += timedelta(days=1)

    return seconds / 3600


def load_calendar(
    metrics_config: dict[str, Any],
    extra_holidays: list[str],
) -> WorkCalendar:
    config = metrics_config.get("work_calendar", {})

    day_numbers = {
        "Monday": 0,
        "Tuesday": 1,
        "Wednesday": 2,
        "Thursday": 3,
        "Friday": 4,
        "Saturday": 5,
        "Sunday": 6,
    }

    default_days = [
        "Sunday",
        "Monday",
        "Tuesday",
        "Wednesday",
        "Thursday",
    ]

    working_days = frozenset(
        day_numbers[name]
        for name in config.get(
            "working_days",
            default_days,
        )
        if name in day_numbers
    )

    holiday_values = [
        *config.get("holiday_dates", []),
        *extra_holidays,
    ]

    holidays: set[date] = set()

    for value in holiday_values:
        parsed = parse_date_value(value)

        if parsed is not None:
            holidays.add(parsed)

    windows = config.get(
        "daily_windows",
        [{"start": "09:00", "end": "17:00"}],
    )

    if not windows:
        raise ValueError("No work window was configured.")

    start_hour = int(
        windows[0]["start"].split(":")[0]
    )

    end_hour = int(
        windows[0]["end"].split(":")[0]
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
    if value is None:
        return []

    if isinstance(value, float) and pd.isna(value):
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

    labels: list[str] = []

    for item in values:
        label = str(item).strip()

        if label and label not in labels:
            labels.append(label)

    return labels


def get_status_events(
    history: dict[str, Any] | None,
    cutoff: pd.Timestamp,
    timezone_name: str,
) -> list[dict[str, Any]]:
    if not history:
        return []

    events: list[dict[str, Any]] = []

    for event in history.get("status_events", []):
        changed_at = parse_timestamp(
            event.get("changed_at"),
            timezone_name,
        )

        if changed_at is None:
            continue

        if changed_at > cutoff:
            continue

        events.append(
            {
                **event,
                "_changed_at": changed_at,
            }
        )

    events.sort(
        key=lambda item: item["_changed_at"]
    )

    return events


def build_status_intervals(
    created_at: pd.Timestamp | None,
    events: list[dict[str, Any]],
    cutoff: pd.Timestamp,
) -> tuple[list[dict[str, Any]], bool]:
    if created_at is None or not events:
        return [], False

    current_status = events[0].get("from_status")

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

        if from_status and from_status != current_status:
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

    if cursor <= cutoff:
        intervals.append(
            {
                "status": current_status,
                "start": cursor,
                "end": cutoff,
            }
        )

    return intervals, True


def status_time_summary(
    intervals: list[dict[str, Any]],
    calendar: WorkCalendar,
) -> dict[str, dict[str, float]]:
    result: dict[str, dict[str, float]] = {}

    for interval in intervals:
        status = str(interval["status"])
        start = interval["start"]
        end = interval["end"]

        elapsed = elapsed_hours(start, end)
        business = business_hours_between(
            start,
            end,
            calendar,
        )

        if elapsed is None or business is None:
            continue

        if status not in result:
            result[status] = {
                "elapsed_hours": 0.0,
                "business_hours": 0.0,
            }

        result[status]["elapsed_hours"] += elapsed
        result[status]["business_hours"] += business

    return result


def first_entry(
    events: list[dict[str, Any]],
    status: str,
) -> pd.Timestamp | None:
    for event in events:
        if event.get("to_status") == status:
            return event["_changed_at"]

    return None


def latest_entry(
    events: list[dict[str, Any]],
    status: str,
) -> pd.Timestamp | None:
    values = [
        event["_changed_at"]
        for event in events
        if event.get("to_status") == status
    ]

    return max(values) if values else None


def count_transition(
    events: list[dict[str, Any]],
    from_status: str,
    to_status: str,
) -> int:
    return sum(
        1
        for event in events
        if event.get("from_status") == from_status
        and event.get("to_status") == to_status
    )


def calculate_task(
    task: dict[str, Any],
    history: dict[str, Any] | None,
    cutoff: pd.Timestamp,
    calendar: WorkCalendar,
) -> dict[str, Any]:
    issue_key = str(
        task.get("issue_key") or ""
    ).strip()

    created_at = parse_timestamp(
        task.get("created_at"),
        calendar.timezone_name,
    )

    due_date = parse_date_value(
        task.get("due_date")
    )

    planned_start_date = parse_date_value(
        task.get("planned_start_date")
    )

    current_status = task.get("current_status")

    if current_status is not None:
        current_status = str(
            current_status
        ).strip()

    labels = normalize_labels(
        task.get("labels")
    )

    events = get_status_events(
        history,
        cutoff,
        calendar.timezone_name,
    )

    intervals, history_complete = build_status_intervals(
        created_at,
        events,
        cutoff,
    )

    status_times = status_time_summary(
        intervals,
        calendar,
    )

    actual_start = first_entry(
        events,
        "In Progress",
    )

    latest_done = latest_entry(
        events,
        "Done",
    )

    final_status = (
        intervals[-1]["status"]
        if intervals
        else current_status
    )

    completed = final_status == "Done"
    rejected = final_status == "Rejected"
    open_task = not completed and not rejected
    wip = final_status in {
        "In Progress",
        "In Review",
    }

    completed_at = latest_done if completed else None

    execution_elapsed = elapsed_hours(
        actual_start,
        completed_at,
    )

    execution_business = business_hours_between(
        actual_start,
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
        actual_start,
    )

    start_business = business_hours_between(
        created_at,
        actual_start,
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

    current_status_start = (
        intervals[-1]["start"]
        if intervals
        else None
    )

    current_status_elapsed = elapsed_hours(
        current_status_start,
        cutoff,
    )

    current_status_business = business_hours_between(
        current_status_start,
        cutoff,
        calendar,
    )

    on_time: bool | None = None
    schedule_variance: int | None = None
    overdue_days: int | None = None

    cutoff_local_date = cutoff.tz_convert(
        calendar.timezone
    ).date()

    if due_date is not None:
        if completed_at is not None:
            completed_local_date = completed_at.tz_convert(
                calendar.timezone
            ).date()

            schedule_variance = (
                completed_local_date - due_date
            ).days

            on_time = schedule_variance <= 0
            overdue_days = max(
                0,
                schedule_variance,
            )

        elif open_task:
            overdue_days = max(
                0,
                (cutoff_local_date - due_date).days,
            )

    start_variance: int | None = None

    if planned_start_date and actual_start:
        actual_start_date = actual_start.tz_convert(
            calendar.timezone
        ).date()

        start_variance = (
            actual_start_date - planned_start_date
        ).days

    if not history_complete:
        history_note = (
            "Status history is incomplete or unavailable."
        )
    else:
        history_note = None

    return {
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
        "created_at": (
            created_at.isoformat()
            if created_at is not None
            else None
        ),
        "actual_start_at": (
            actual_start.isoformat()
            if actual_start is not None
            else None
        ),
        "completed_at": (
            completed_at.isoformat()
            if completed_at is not None
            else None
        ),
        "due_date": (
            due_date.isoformat()
            if due_date is not None
            else None
        ),
        "planned_start_date": (
            planned_start_date.isoformat()
            if planned_start_date is not None
            else None
        ),
        "is_completed": completed,
        "is_rejected": rejected,
        "is_open": open_task,
        "is_wip": wip,
        "history_complete": history_complete,
        "history_note": history_note,
        "execution_elapsed_hours": execution_elapsed,
        "execution_business_hours": execution_business,
        "lead_time_elapsed_hours": lead_elapsed,
        "lead_time_business_hours": lead_business,
        "time_to_start_elapsed_hours": start_elapsed,
        "time_to_start_business_hours": start_business,
        "task_age_elapsed_hours": (
            age_elapsed if open_task else None
        ),
        "task_age_business_hours": (
            age_business if open_task else None
        ),
        "current_status_age_elapsed_hours": (
            current_status_elapsed
            if open_task
            else None
        ),
        "current_status_age_business_hours": (
            current_status_business
            if open_task
            else None
        ),
        "on_time_completion": on_time,
        "schedule_variance_days": schedule_variance,
        "overdue_days": overdue_days,
        "start_schedule_variance_days": start_variance,
        "rework_count": (
            count_transition(
                events,
                "In Review",
                "In Progress",
            )
            if history_complete
            else None
        ),
        "replanning_count": (
            count_transition(
                events,
                "In Review",
                "To Do",
            )
            if history_complete
            else None
        ),
        "re_evaluation_count": (
            count_transition(
                events,
                "In Review",
                "In Triage",
            )
            if history_complete
            else None
        ),
        "reopen_count": (
            sum(
                1
                for event in events
                if event.get("from_status") == "Done"
                and event.get("to_status")
                not in {"Done", "Rejected"}
            )
            if history_complete
            else None
        ),
        "time_in_status": status_times,
        "evaluation_cutoff": cutoff.isoformat(),
        "work_calendar_timezone": calendar.timezone_name,
        "work_calendar_days": "Sunday-Thursday",
        "work_calendar_window": (
            f"{calendar.start_hour:02d}:00-"
            f"{calendar.end_hour:02d}:00"
        ),
    }


def analyze_tasks(
    tasks_frame: pd.DataFrame,
    histories: dict[str, dict[str, Any]],
    cutoff: pd.Timestamp,
    calendar: WorkCalendar,
) -> pd.DataFrame:
    records: list[dict[str, Any]] = []

    for _, row in tasks_frame.iterrows():
        task = row.to_dict()
        issue_key = str(
            task.get("issue_key") or ""
        ).strip()

        records.append(
            calculate_task(
                task,
                histories.get(issue_key),
                cutoff,
                calendar,
            )
        )

    return pd.DataFrame(records)


def safe_mean(series: pd.Series) -> float | None:
    values = pd.to_numeric(
        series,
        errors="coerce",
    ).dropna()

    return (
        float(values.mean())
        if not values.empty
        else None
    )


def safe_median(series: pd.Series) -> float | None:
    values = pd.to_numeric(
        series,
        errors="coerce",
    ).dropna()

    return (
        float(values.median())
        if not values.empty
        else None
    )


def aggregate(
    frame: pd.DataFrame,
    group_columns: list[str],
) -> pd.DataFrame:
    if frame.empty:
        return pd.DataFrame()

    if group_columns:
        groups = frame.groupby(
            group_columns,
            dropna=False,
        )
    else:
        groups = [((), frame)]

    rows: list[dict[str, Any]] = []

    for group_key, group in groups:
        if not isinstance(group_key, tuple):
            group_key = (group_key,)

        record = {
            column: value
            for column, value in zip(
                group_columns,
                group_key,
            )
        }

        total = int(
            group["issue_key"].nunique()
        )

        completed = int(
            group["is_completed"].sum()
        )

        rejected = int(
            group["is_rejected"].sum()
        )

        open_tasks = int(
            group["is_open"].sum()
        )

        wip = int(
            group["is_wip"].sum()
        )

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

        rework_count = int(
            (valid_rework["rework_count"] > 0).sum()
        )

        record.update(
            {
                "total_tasks": total,
                "completed_tasks": completed,
                "rejected_tasks": rejected,
                "open_tasks": open_tasks,
                "wip_tasks": wip,
                "completion_rate": (
                    completed / total * 100
                    if total
                    else None
                ),
                "rejection_rate": (
                    rejected / total * 100
                    if total
                    else None
                ),
                "on_time_tasks": on_time_count,
                "on_time_valid_tasks": len(valid_on_time),
                "on_time_completion_rate": (
                    on_time_count
                    / len(valid_on_time)
                    * 100
                    if len(valid_on_time)
                    else None
                ),
                "overdue_open_tasks": overdue_count,
                "overdue_valid_tasks": len(valid_overdue),
                "open_overdue_rate": (
                    overdue_count
                    / len(valid_overdue)
                    * 100
                    if len(valid_overdue)
                    else None
                ),
                "tasks_with_rework": rework_count,
                "rework_valid_tasks": len(valid_rework),
                "tasks_with_rework_rate": (
                    rework_count
                    / len(valid_rework)
                    * 100
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

        rows.append(record)

    return pd.DataFrame(rows)


def flatten_for_excel(
    frame: pd.DataFrame,
) -> pd.DataFrame:
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


def write_output(
    output_path: Path,
    task_metrics: pd.DataFrame,
    overall: pd.DataFrame,
    by_assignee: pd.DataFrame,
    by_issue_type: pd.DataFrame,
) -> None:
    output_path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    with pd.ExcelWriter(
        output_path,
        engine="openpyxl",
    ) as writer:
        flatten_for_excel(
            task_metrics
        ).to_excel(
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


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Calculate Jira performance metrics."
    )

    parser.add_argument(
        "--tasks",
        required=True,
        help="Normalized workbook path.",
    )

    parser.add_argument(
        "--history",
        required=True,
        help="Jira history JSON path.",
    )

    parser.add_argument(
        "--metrics-config",
        default="configs/jira_metrics_config.json",
        help="Metrics configuration path.",
    )

    parser.add_argument(
        "--cutoff",
        required=True,
        help="Evaluation cutoff with timezone.",
    )

    parser.add_argument(
        "--output",
        required=True,
        help="Output workbook path.",
    )

    parser.add_argument(
        "--holiday",
        action="append",
        default=[],
        help="Optional holiday date YYYY-MM-DD.",
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
            args.holiday,
        )

        cutoff = parse_timestamp(
            args.cutoff,
            calendar.timezone_name,
        )

        if cutoff is None:
            raise ValueError(
                "Invalid evaluation cutoff."
            )

        tasks_frame = pd.read_excel(
            args.tasks,
            sheet_name="normalized_tasks",
            dtype=object,
        )

        histories = load_json(
            Path(args.history)
        )

        task_metrics = analyze_tasks(
            tasks_frame,
            histories,
            cutoff,
            calendar,
        )

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

        write_output(
            Path(args.output),
            task_metrics,
            overall,
            by_assignee,
            by_issue_type,
        )

        print(
            f"Analyzed tasks: {len(task_metrics)}"
        )

        print(
            f"Output written to: {args.output}"
        )

        return 0

    except Exception as exc:
        LOGGER.exception(
            "Metrics calculation failed: %s",
            exc,
        )
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
