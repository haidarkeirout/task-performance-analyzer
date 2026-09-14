"""Admin-managed employee identities used by the employee analysis flow."""
from __future__ import annotations

import csv
from dataclasses import dataclass
from pathlib import Path


DIRECTORY_PATH = Path(__file__).resolve().parents[1] / "config" / "employee_directory.csv"
REQUIRED_COLUMNS = {
    "employee_name", "department", "primary_source", "jira_account_id", "clickup_user_id", "active",
}


@dataclass(frozen=True)
class EmployeeRecord:
    name: str
    department: str
    primary_source: str
    jira_account_id: str = ""
    clickup_user_id: str = ""
    active: bool = True


def _active(value: str) -> bool:
    return str(value or "").strip().casefold() in {"1", "true", "yes", "y", "on"}


def load_employee_directory(path: Path = DIRECTORY_PATH) -> list[EmployeeRecord]:
    """Load and validate the administrator-maintained directory."""
    if not path.exists():
        raise ValueError(f"Employee directory is missing: {path}")
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        columns = {str(column or "").strip() for column in (reader.fieldnames or [])}
        missing = REQUIRED_COLUMNS - columns
        if missing:
            raise ValueError("Employee directory is missing columns: " + ", ".join(sorted(missing)))
        records = []
        seen = set()
        for line_number, row in enumerate(reader, 2):
            name = str(row.get("employee_name") or "").strip()
            source = str(row.get("primary_source") or "").strip().casefold()
            if not name:
                raise ValueError(f"Employee directory row {line_number} has no employee_name.")
            if name.casefold() in seen:
                raise ValueError(f"Employee directory contains a duplicate employee: {name}.")
            if source not in {"jira", "clickup"}:
                raise ValueError(f"Employee directory row {line_number} has an invalid primary_source.")
            jira_id = str(row.get("jira_account_id") or "").strip()
            clickup_id = str(row.get("clickup_user_id") or "").strip()
            if source == "jira" and not jira_id:
                raise ValueError(f"Employee directory row {line_number} needs jira_account_id.")
            if source == "clickup" and not clickup_id:
                raise ValueError(f"Employee directory row {line_number} needs clickup_user_id.")
            seen.add(name.casefold())
            records.append(EmployeeRecord(
                name=name,
                department=str(row.get("department") or "").strip(),
                primary_source=source,
                jira_account_id=jira_id,
                clickup_user_id=clickup_id,
                active=_active(row.get("active", "")),
            ))
    return [record for record in records if record.active]
