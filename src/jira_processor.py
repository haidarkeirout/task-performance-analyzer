"""
Jira task-performance analyzer
Stage 1:
- Load configuration files
- Import Jira Excel export
- Validate required columns
- Normalize dates, labels, assignees, and issue types
- Detect duplicates and data-quality problems
- Export a normalized workbook and validation log
"""

from __future__ import annotations

import argparse
import json
import logging
import re
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

import pandas as pd
from zoneinfo import ZoneInfo


ROOT_DIR = Path(__file__).resolve().parents[1]
DEFAULT_CONFIG_DIR = ROOT_DIR / "configs"


class ConfigurationError(ValueError):
    pass


class ImportValidationError(ValueError):
    pass


def load_json(path: Path) -> dict[str, Any]:
    if not path.exists():
        raise ConfigurationError(f"Configuration file not found: {path}")

    try:
        with path.open("r", encoding="utf-8") as file:
            value = json.load(file)
    except json.JSONDecodeError as exc:
        raise ConfigurationError(f"Invalid JSON in {path}: {exc}") from exc

    if not isinstance(value, dict):
        raise ConfigurationError(f"Configuration root must be an object: {path}")

    return value


def normalize_header(value: Any) -> str:
    text = "" if value is None else str(value)
    text = text.strip().casefold()
    text = re.sub(r"\s+", " ", text)
    return text


def is_blank(value: Any) -> bool:
    if value is None:
        return True

    if isinstance(value, float) and pd.isna(value):
        return True

    if pd.isna(value) is True:
        return True

    return isinstance(value, str) and not value.strip()


def clean_text(value: Any) -> Any:
    if is_blank(value):
        return None

    if isinstance(value, str):
        return value.strip()

    return value


def parse_labels(value: Any, delimiter: str | None = None) -> list[str] | None:
    if is_blank(value):
        return []

    if isinstance(value, list):
        values = value
    else:
        text = str(value).strip()

        if text.startswith("[") and text.endswith("]"):
            try:
                parsed = json.loads(text)
                values = parsed if isinstance(parsed, list) else [text]
            except json.JSONDecodeError:
                values = [text]
        elif delimiter and delimiter in text:
            values = text.split(delimiter)
        else:
            values = [text]

    labels: list[str] = []

    for item in values:
        if not is_blank(item):
            label = str(item).strip()
            if label and label not in labels:
                labels.append(label)

    return labels


def parse_datetime_value(
    value: Any,
    *,
    date_only: bool,
    source_timezone: ZoneInfo,
    explicit_format: str | None = None,
) -> str | None:
    if is_blank(value):
        return None

    try:
        if isinstance(value, (int, float)) and not isinstance(value, bool):
            timestamp = pd.Timestamp(
                pd.to_datetime(
                    float(value),
                    unit="D",
                    origin="1899-12-30",
                )
            )
        elif explicit_format and isinstance(value, str):
            timestamp = pd.Timestamp(
                pd.to_datetime(value.strip(), format=explicit_format)
            )
        else:
            timestamp = pd.Timestamp(pd.to_datetime(value, errors="raise"))

        if date_only:
            return timestamp.date().isoformat()

        if timestamp.tzinfo is None:
            timestamp = timestamp.tz_localize(source_timezone)
        else:
            timestamp = timestamp.tz_convert("UTC")

        return timestamp.tz_convert("UTC").isoformat()

    except Exception:
        return None


def choose_sheet(
    workbook: dict[str, pd.DataFrame],
    input_config: dict[str, Any],
    requested_sheet: str | None,
) -> tuple[str, pd.DataFrame]:
    if requested_sheet:
        if requested_sheet not in workbook:
            available = ", ".join(workbook.keys())
            raise ImportValidationError(
                f"Requested sheet '{requested_sheet}' was not found. "
                f"Available sheets: {available}"
            )
        return requested_sheet, workbook[requested_sheet]

    if len(workbook) == 1:
        sheet_name = next(iter(workbook))
        return sheet_name, workbook[sheet_name]

    required_columns = input_config.get("required_columns", [])
    required_headers = {normalize_header(item) for item in required_columns}

    matching_sheets: list[tuple[str, pd.DataFrame]] = []

    for name, frame in workbook.items():
        frame_headers = {normalize_header(item) for item in frame.columns}
        if required_headers.issubset(frame_headers):
            matching_sheets.append((name, frame))

    if len(matching_sheets) == 1:
        return matching_sheets[0]

    available = ", ".join(workbook.keys())

    raise ImportValidationError(
        "Could not select a unique worksheet automatically. "
        f"Available sheets: {available}. Use --sheet."
    )


def build_source_column_lookup(columns: list[Any]) -> dict[str, str]:
    lookup: dict[str, str] = {}

    for column in columns:
        normalized = normalize_header(column)

        if normalized in lookup:
            raise ImportValidationError(
                f"Duplicate worksheet headers detected: '{column}'"
            )

        lookup[normalized] = str(column)

    return lookup


def normalize_tasks(
    frame: pd.DataFrame,
    *,
    input_config: dict[str, Any],
    source_timezone: ZoneInfo,
    labels_delimiter: str | None,
    explicit_date_format: str | None,
) -> tuple[pd.DataFrame, list[dict[str, Any]]]:
    mapping = input_config.get("column_mapping", {})
    required_columns = input_config.get("required_columns", [])

    if not isinstance(mapping, dict):
        raise ConfigurationError("column_mapping must be an object")

    source_lookup = build_source_column_lookup(list(frame.columns))
    errors: list[dict[str, Any]] = []
    warnings: list[dict[str, Any]] = []

    normalized = frame.copy()
    normalized.insert(0, "_source_row", range(2, len(frame) + 2))

    field_sources: dict[str, str | None] = {}

    for source_name, internal_name in mapping.items():
        source_column = source_lookup.get(normalize_header(source_name))
        field_sources[internal_name] = source_column

        if source_column is None:
            normalized[internal_name] = None
            continue

        normalized[internal_name] = normalized[source_column]

    for required_name in required_columns:
        source_column = source_lookup.get(normalize_header(required_name))

        if source_column is None:
            errors.append(
                {
                    "type": "missing_required_column",
                    "column": required_name,
                    "message": f"Required column '{required_name}' is missing.",
                }
            )
            continue

        empty_rows = normalized[
            normalized[source_column].apply(is_blank)
        ]["_source_row"].tolist()

        if empty_rows:
            errors.append(
                {
                    "type": "blank_required_value",
                    "column": required_name,
                    "rows": empty_rows,
                    "message": (
                        f"Required column '{required_name}' contains blank "
                        "values."
                    ),
                }
            )

    if field_sources.get("labels"):
        normalized["labels"] = normalized["labels"].apply(
            lambda value: parse_labels(value, labels_delimiter)
        )

    date_fields = {
        "created_at": False,
        "due_date": True,
        "planned_start_date": True,
    }

    for field_name, date_only in date_fields.items():
        if field_name not in normalized.columns:
            continue

        normalized[field_name] = normalized[field_name].apply(
            lambda value: parse_datetime_value(
                value,
                date_only=date_only,
                source_timezone=source_timezone,
                explicit_format=explicit_date_format,
            )
        )

        original_source = field_sources.get(field_name)

        if original_source:
            invalid_rows = normalized[
                normalized[field_name].isna()
                & normalized[original_source].apply(lambda value: not is_blank(value))
            ]["_source_row"].tolist()

            if invalid_rows:
                warnings.append(
                    {
                        "type": "invalid_date",
                        "column": original_source,
                        "rows": invalid_rows,
                        "message": (
                            f"Some values in '{original_source}' could not be "
                            "converted to the configured date format."
                        ),
                    }
                )

    text_fields = [
        "issue_key",
        "issue_id",
        "task_name",
        "assignee_name",
        "assignee_id",
        "priority",
        "current_status",
        "resolution",
        "issue_type",
    ]

    for field_name in text_fields:
        if field_name in normalized.columns:
            normalized[field_name] = normalized[field_name].apply(clean_text)

    if "issue_key" in normalized.columns:
        duplicate_mask = normalized["issue_key"].duplicated(
            keep=False,
            na=False,
        )

        duplicate_groups = normalized[duplicate_mask].groupby(
            "issue_key",
            dropna=False,
        )

        rows_to_remove: list[int] = []

        for issue_key, group in duplicate_groups:
            comparable = group.drop(
                columns=["_source_row"],
                errors="ignore",
            ).fillna("<NULL>")

            if comparable.astype(str).drop_duplicates().shape[0] == 1:
                duplicate_rows = group["_source_row"].tolist()
                rows_to_remove.extend(group.index.tolist()[1:])

                warnings.append(
                    {
                        "type": "identical_duplicate_rows",
                        "issue_key": str(issue_key),
                        "rows": duplicate_rows,
                        "message": (
                            "Identical duplicate rows detected. "
                            "The first row was retained."
                        ),
                    }
                )
            else:
                errors.append(
                    {
                        "type": "conflicting_duplicate_rows",
                        "issue_key": str(issue_key),
                        "rows": group["_source_row"].tolist(),
                        "message": (
                            "The same Issue key appears with conflicting "
                            "values."
                        ),
                    }
                )

        if rows_to_remove:
            normalized = normalized.drop(index=rows_to_remove)

    for internal_name, source_column in field_sources.items():
        if source_column is None:
            warnings.append(
                {
                    "type": "missing_optional_column",
                    "column": internal_name,
                    "message": (
                        f"No source column was found for '{internal_name}'. "
                        "The field was left unavailable."
                    ),
                }
            )

    for item in warnings:
        errors.append(
            {
                "severity": "warning",
                **item,
            }
        )

    normalized = normalized.reset_index(drop=True)

    validation_log = errors

    return normalized, validation_log


def write_output(
    output_path: Path,
    normalized: pd.DataFrame,
    validation_log: list[dict[str, Any]],
) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)

    log_frame = pd.DataFrame(validation_log)

    if log_frame.empty:
        log_frame = pd.DataFrame(
            [
                {
                    "severity": "info",
                    "type": "validation_completed",
                    "message": "No validation issues were detected.",
                }
            ]
        )

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        export_frame = normalized.copy()

        if "labels" in export_frame.columns:
            export_frame["labels"] = export_frame["labels"].apply(
                lambda values: json.dumps(values, ensure_ascii=False)
                if isinstance(values, list)
                else values
            )

        export_frame.to_excel(
            writer,
            index=False,
            sheet_name="normalized_tasks",
        )

        log_frame.to_excel(
            writer,
            index=False,
            sheet_name="validation_log",
        )


def validate_calendar(metrics_config: dict[str, Any]) -> None:
    calendar = metrics_config.get("work_calendar", {})

    if calendar.get("timezone") != "Asia/Damascus":
        raise ConfigurationError(
            "The work calendar timezone must be Asia/Damascus."
        )

    if calendar.get("working_days") != [
        "Sunday",
        "Monday",
        "Tuesday",
        "Wednesday",
        "Thursday",
    ]:
        raise ConfigurationError(
            "The configured working days must be Sunday through Thursday."
        )

    if calendar.get("weekend_days") != ["Friday", "Saturday"]:
        raise ConfigurationError(
            "The configured weekend must be Friday and Saturday."
        )

    windows = calendar.get("daily_windows", [])

    if not windows:
        raise ConfigurationError(
            "At least one daily work window must be configured."
        )


def run(args: argparse.Namespace) -> int:
    config_dir = Path(args.config_dir)

    input_config = load_json(config_dir / "jira_input_config.json")
    metrics_config = load_json(config_dir / "jira_metrics_config.json")

    validate_calendar(metrics_config)

    source_timezone = ZoneInfo(args.source_timezone)

    workbook = pd.read_excel(
        args.input,
        sheet_name=None,
        dtype=object,
    )

    sheet_name, frame = choose_sheet(
        workbook,
        input_config,
        args.sheet,
    )

    normalized, validation_log = normalize_tasks(
        frame,
        input_config=input_config,
        source_timezone=source_timezone,
        labels_delimiter=args.labels_delimiter,
        explicit_date_format=args.date_format,
    )

    output_path = Path(args.output)

    write_output(
        output_path,
        normalized,
        validation_log,
    )

    error_count = sum(
        1
        for item in validation_log
        if item.get("severity") != "warning"
        and item.get("type") not in {
            "missing_optional_column",
            "invalid_date",
            "identical_duplicate_rows",
        }
    )

    warning_count = sum(
        1
        for item in validation_log
        if item.get("severity") == "warning"
        or item.get("type")
        in {
            "missing_optional_column",
            "invalid_date",
            "identical_duplicate_rows",
        }
    )

    print(f"Selected worksheet: {sheet_name}")
    print(f"Normalized task rows: {len(normalized)}")
    print(f"Validation errors: {error_count}")
    print(f"Validation warnings: {warning_count}")
    print(f"Output written to: {output_path}")

    if error_count:
        return 2

    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Normalize and validate a Jira Excel export."
    )

    parser.add_argument(
        "--input",
        required=True,
        help="Path to the Jira Excel workbook.",
    )

    parser.add_argument(
        "--output",
        required=True,
        help="Path for the normalized output workbook.",
    )

    parser.add_argument(
        "--sheet",
        default=None,
        help="Worksheet name. If omitted, the script selects it automatically.",
    )

    parser.add_argument(
        "--config-dir",
        default=str(DEFAULT_CONFIG_DIR),
        help="Directory containing the JSON configuration files.",
    )

    parser.add_argument(
        "--source-timezone",
        default="Asia/Damascus",
        help="Timezone for naive source timestamps.",
    )

    parser.add_argument(
        "--labels-delimiter",
        default=None,
        help="Optional explicit delimiter for multiple labels, such as ';'.",
    )

    parser.add_argument(
        "--date-format",
        default=None,
        help="Optional explicit format for non-ISO text dates.",
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
        return run(args)
    except (ConfigurationError, ImportValidationError, FileNotFoundError) as exc:
        logging.error(str(exc))
        return 2
    except Exception as exc:
        logging.exception("Unexpected processing error: %s", exc)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
