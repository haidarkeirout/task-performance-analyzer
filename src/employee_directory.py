"""Employee directory loaded from the admin-maintained OneDrive workbook."""
from __future__ import annotations

import os
import re
import threading
import time
import unicodedata
from dataclasses import dataclass
from io import BytesIO
from typing import Iterable, Mapping
from urllib.parse import parse_qsl, urlencode, urlparse, urlunparse

import pandas as pd
import requests


# The workbook share URL is deployment configuration, not source code. Keep
# it in Streamlit Secrets or the EMPLOYEE_DIRECTORY_URL environment variable
# so a share token cannot be committed to the public repository.
DEFAULT_DIRECTORY_URL = ""
REQUIRED_COLUMNS = {
    "employee_name",
    "department",
    "primary_source",
    "jira_account_id",
    "clickup_user_id",
    "active",
}
DIRECTORY_CACHE_TTL_SECONDS = 300
_CACHE_LOCK = threading.Lock()
_DIRECTORY_CACHE: dict[str, tuple[float, tuple["EmployeeRecord", ...]]] = {}
_LAST_GOOD_DIRECTORY: dict[str, tuple["EmployeeRecord", ...]] = {}


@dataclass(frozen=True)
class EmployeeRecord:
    name: str
    department: str
    primary_source: str
    jira_account_id: str = ""
    clickup_user_id: str = ""
    jira_department: str = ""
    clickup_department: str = ""
    jira_position: str = ""
    clickup_position: str = ""
    active: bool = True

    @property
    def sources(self) -> tuple[str, ...]:
        values = []
        if self.jira_account_id:
            values.append("jira")
        if self.clickup_user_id:
            values.append("clickup")
        return tuple(values)


def _active(value: object) -> bool:
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)) and not pd.isna(value):
        return bool(value)
    return str(value or "").strip().casefold() in {
        "1",
        "true",
        "yes",
        "y",
        "on",
    }


def _text(value: object) -> str:
    """Return a stable cell value without leaking pandas' NaN marker."""
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    return str(value).strip()


def _primary_source(value: object) -> str:
    """Normalize harmless Excel variations while keeping routing unambiguous."""
    raw = unicodedata.normalize("NFKC", _text(value)).casefold()
    compact = re.sub(r"[^a-z0-9]+", "", raw)
    normalized = {
        "jira": "jira",
        "jiracloud": "jira",
        "atlassianjira": "jira",
        "clickup": "clickup",
        "clickupcloud": "clickup",
        "both": "both",
        "jiraandclickup": "both",
        "clickupandjira": "both",
        "jiraclickup": "both",
        "clickupjira": "both",
        "all": "both",
    }.get(compact)
    if normalized:
        return normalized
    if "jira" in compact and "clickup" in compact:
        return "both"
    if "jira" in compact:
        return "jira"
    if "clickup" in compact:
        return "clickup"
    return raw


def _directory_url() -> str:
    """Read an optional override while keeping the approved public link as default."""
    configured = os.getenv("EMPLOYEE_DIRECTORY_URL", "").strip()
    if configured:
        return configured

    try:
        import streamlit as st

        configured = str(st.secrets.get("EMPLOYEE_DIRECTORY_URL", "")).strip()
    except Exception:
        configured = ""

    return configured or DEFAULT_DIRECTORY_URL


def _with_download_parameter(url: str) -> str:
    parsed = urlparse(url)
    query = dict(parse_qsl(parsed.query, keep_blank_values=True))
    query["download"] = "1"
    return urlunparse(parsed._replace(query=urlencode(query)))


def _direct_download_url(url: str) -> str | None:
    parsed = urlparse(url)
    query = dict(parse_qsl(parsed.query, keep_blank_values=True))
    resid = query.get("resid")
    if not resid:
        return None
    return "https://onedrive.live.com/download?" + urlencode({"resid": resid})


def _download_workbook(url: str) -> bytes:
    """Download a publicly shared OneDrive workbook without Graph authentication."""
    if not str(url or "").strip():
        raise ValueError(
            "The employee directory URL is not configured. Set "
            "EMPLOYEE_DIRECTORY_URL in the deployment Secrets."
        )
    candidates = [_with_download_parameter(url), url]
    direct = _direct_download_url(url)
    if direct:
        candidates.insert(0, direct)

    visited = set()
    errors = []

    for candidate in candidates:
        if not candidate or candidate in visited:
            continue
        visited.add(candidate)

        try:
            response = requests.get(
                candidate,
                headers={"User-Agent": "Task-Performance-Analyzer/1.0"},
                timeout=(5, 15),
                allow_redirects=True,
            )
            response.raise_for_status()
        except requests.RequestException as exc:
            errors.append(str(exc))
            continue

        payload = response.content
        content_type = response.headers.get("content-type", "").casefold()
        if payload.startswith(b"PK") or "spreadsheet" in content_type:
            return payload

        redirected = _direct_download_url(response.url)
        if redirected and redirected not in visited:
            candidates.insert(0, redirected)

    detail = errors[-1] if errors else "OneDrive returned a web page instead of an Excel workbook."
    raise ValueError(
        "The online employee directory could not be read. "
        "Set the OneDrive file to 'Anyone with the link can view', "
        "keep it as an .xlsx workbook, and verify the share link. "
        f"Technical detail: {detail}"
    )


def _remote_rows(url: str) -> list[dict[str, object]]:
    payload = _download_workbook(url)
    try:
        workbook = pd.read_excel(
            BytesIO(payload),
            sheet_name=None,
            dtype=object,
        )
    except Exception as exc:
        raise ValueError(
            "The online employee directory was downloaded, but it is not a readable Excel workbook."
        ) from exc

    for sheet_name, frame in workbook.items():
        if frame is None:
            continue
        normalized = frame.copy()
        normalized.columns = [
            str(column or "").strip()
            for column in normalized.columns
        ]
        if REQUIRED_COLUMNS.issubset(set(normalized.columns)):
            return normalized.fillna("").to_dict("records")

    available = ", ".join(str(name) for name in workbook)
    raise ValueError(
        "The online employee directory does not contain the required columns "
        f"in any sheet. Available sheets: {available or 'none'}."
    )


def _build_records(rows: Iterable[Mapping[str, object]], source_label: str) -> list[EmployeeRecord]:
    records: dict[str, EmployeeRecord] = {}

    for line_number, row in enumerate(rows, 2):
        name = _text(row.get("employee_name"))
        active = _active(row.get("active", ""))

        # The workbook may retain inactive employees or blank template rows.
        # They are not part of Employee analysis and must not block active rows
        # because their source/ID fields are intentionally incomplete.
        if not active:
            continue

        raw_source = _text(row.get("primary_source"))
        source = _primary_source(raw_source)
        if not name:
            raise ValueError(
                f"Employee directory row {line_number} in {source_label} has no employee_name."
            )
        if source not in {"jira", "clickup", "both"}:
            raise ValueError(
                f"Employee directory row {line_number} ({name}) has an invalid "
                f"primary_source {raw_source!r}. Use 'jira' or 'clickup'."
            )

        jira_id = _text(row.get("jira_account_id"))
        clickup_id = _text(row.get("clickup_user_id"))
        if source == "jira" and not jira_id:
            raise ValueError(
                f"Employee directory row {line_number} needs jira_account_id."
            )
        if source == "clickup" and not clickup_id:
            raise ValueError(
                f"Employee directory row {line_number} needs clickup_user_id."
            )
        if source == "both" and not (jira_id or clickup_id):
            raise ValueError(
                f"Employee directory row {line_number} ({name}) needs at least one "
                "jira_account_id or clickup_user_id."
            )

        row_record = EmployeeRecord(
            name=name,
            department=_text(row.get("department")),
            primary_source="both" if jira_id and clickup_id else source,
            jira_account_id=jira_id,
            clickup_user_id=clickup_id,
            jira_department=_text(row.get("jira_department")) or (
                _text(row.get("department")) if jira_id else ""
            ),
            clickup_department=_text(row.get("clickup_department")) or (
                _text(row.get("department")) if clickup_id else ""
            ),
            jira_position=_text(row.get("jira_position")),
            clickup_position=_text(row.get("clickup_position")),
            active=active,
        )
        key = name.casefold()
        previous = records.get(key)
        if previous is None:
            records[key] = row_record
            continue

        # A person can have one Jira row and one ClickUp row. Merge only
        # cross-source rows; duplicate IDs in the same source remain an error
        # so two different people sharing a name are not silently combined.
        if previous.jira_account_id and row_record.jira_account_id:
            raise ValueError(
                f"Employee directory contains duplicate Jira accounts for: {name}."
            )
        if previous.clickup_user_id and row_record.clickup_user_id:
            raise ValueError(
                f"Employee directory contains duplicate ClickUp accounts for: {name}."
            )
        merged_jira = previous.jira_account_id or row_record.jira_account_id
        merged_clickup = previous.clickup_user_id or row_record.clickup_user_id
        records[key] = EmployeeRecord(
            name=previous.name,
            department=previous.department or row_record.department,
            primary_source="both" if merged_jira and merged_clickup else ("jira" if merged_jira else "clickup"),
            jira_account_id=merged_jira,
            clickup_user_id=merged_clickup,
            jira_department=previous.jira_department or row_record.jira_department,
            clickup_department=previous.clickup_department or row_record.clickup_department,
            jira_position=previous.jira_position or row_record.jira_position,
            clickup_position=previous.clickup_position or row_record.clickup_position,
            active=True,
        )

    return [record for record in records.values() if record.active]


def load_employee_directory_with_status() -> tuple[list[EmployeeRecord], str | None]:
    """Load the directory with a short cache and a validated last-good fallback."""
    url = _directory_url()
    now = time.monotonic()
    with _CACHE_LOCK:
        cached = _DIRECTORY_CACHE.get(url)
        if cached and now - cached[0] < DIRECTORY_CACHE_TTL_SECONDS:
            return list(cached[1]), None

    try:
        records = tuple(_build_records(_remote_rows(url), "the online Excel directory"))
    except ValueError as exc:
        with _CACHE_LOCK:
            fallback = _LAST_GOOD_DIRECTORY.get(url)
        if fallback is None:
            raise
        return list(fallback), (
            "The employee directory is temporarily unavailable. "
            "The last successfully validated version is being used. "
            f"Technical detail: {exc}"
        )

    with _CACHE_LOCK:
        _DIRECTORY_CACHE[url] = (now, records)
        _LAST_GOOD_DIRECTORY[url] = records
    return list(records), None


def load_employee_directory() -> list[EmployeeRecord]:
    """Load and validate the online employee directory."""
    records, _warning = load_employee_directory_with_status()
    return records
