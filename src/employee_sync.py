"""Automatic discovery of Jira users and ClickUp workspace members.

This module owns the first-stage employee directory sync. It deliberately keeps
Jira and ClickUp identities as separate records; cross-source matching is a
later feature.
"""
from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timezone
from typing import Any

from clickup_gateway import ClickUpCollectionError, ClickUpGateway
from employee_directory import EmployeeRecord
from jira_gateway import CollectionError, JiraGateway


@dataclass(frozen=True)
class EmployeeSyncResult:
    records: tuple[EmployeeRecord, ...]
    synced_at: str
    warnings: tuple[str, ...] = ()

    @property
    def active_records(self) -> tuple[EmployeeRecord, ...]:
        return tuple(record for record in self.records if record.active)

    @property
    def jira_count(self) -> int:
        return sum(1 for record in self.records if record.primary_source == "jira")

    @property
    def clickup_count(self) -> int:
        return sum(1 for record in self.records if record.primary_source == "clickup")


def _text(value: Any) -> str:
    return str(value or "").strip()


def _clickup_user(member: dict[str, Any]) -> dict[str, Any]:
    nested = member.get("user")
    return nested if isinstance(nested, dict) else member


def _clickup_active(user: dict[str, Any]) -> bool:
    status = _text(user.get("status")).casefold()
    if status in {"inactive", "deactivated", "deleted", "disabled"}:
        return False
    if user.get("deleted") is True or user.get("active") is False:
        return False
    return True


def _jira_records(users: list[dict[str, Any]]) -> list[EmployeeRecord]:
    records: dict[str, EmployeeRecord] = {}
    for user in users:
        # Jira returns Atlassian app/service accounts in the same directory.
        # They are not employees and must not appear in Employee Analysis.
        account_type = _text(user.get("accountType")).casefold()
        if account_type and account_type != "atlassian":
            continue
        account_id = _text(user.get("accountId"))
        name = _text(user.get("displayName")) or _text(user.get("emailAddress"))
        if not account_id or not name:
            continue
        records[account_id] = EmployeeRecord(
            name=name,
            department="",
            primary_source="jira",
            jira_account_id=account_id,
            active=user.get("active", True) is not False,
        )
    return list(records.values())


def _clickup_records(members: list[dict[str, Any]]) -> list[EmployeeRecord]:
    records: dict[str, EmployeeRecord] = {}
    for member in members:
        user = _clickup_user(member)
        member_id = _text(user.get("id"))
        name = (
            _text(user.get("username"))
            or _text(user.get("email"))
            or _text(user.get("initials"))
        )
        if not member_id or not name:
            continue
        records[member_id] = EmployeeRecord(
            name=name,
            department="",
            primary_source="clickup",
            clickup_user_id=member_id,
            active=_clickup_active(user),
        )
    return list(records.values())


def _sorted_records(records: list[EmployeeRecord]) -> tuple[EmployeeRecord, ...]:
    return tuple(
        sorted(
            records,
            key=lambda record: (
                not record.active,
                record.name.casefold(),
                record.primary_source,
                record.jira_account_id or record.clickup_user_id,
            ),
        )
    )


def sync_employees(settings) -> EmployeeSyncResult:
    """Read all accessible users/members from the configured source accounts.

    A source failure is retained as a warning when the other source succeeds.
    No names are matched or merged here.
    """
    records: list[EmployeeRecord] = []
    warnings: list[str] = []
    configured_sources = 0

    if settings.jira_url and settings.jira_email and settings.jira_token:
        configured_sources += 1
        gateway = JiraGateway(settings)
        try:
            records.extend(_jira_records(gateway.users()))
        except CollectionError as exc:
            warnings.append(f"Jira employee sync: {exc}")
        finally:
            gateway.close()

    if settings.clickup_token:
        configured_sources += 1
        gateway = ClickUpGateway(
            settings.clickup_token,
            workspace_id=str(settings.clickup_workspace_id or ""),
        )
        try:
            workspace_id = str(settings.clickup_workspace_id or "")
            members = gateway.workspace_members(workspace_id)
            records.extend(_clickup_records(members))
        except ClickUpCollectionError as exc:
            warnings.append(f"ClickUp employee sync: {exc}")
        finally:
            gateway.close()

    if not configured_sources:
        raise CollectionError("No Jira or ClickUp connection is configured.")

    if not records and warnings:
        raise CollectionError("Employee synchronization failed. " + " | ".join(warnings))

    return EmployeeSyncResult(
        records=_sorted_records(records),
        synced_at=datetime.now(timezone.utc).isoformat(),
        warnings=tuple(warnings),
    )
