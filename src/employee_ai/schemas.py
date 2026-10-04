"""Small typed envelope for Employee AI Assistant context V1."""

from __future__ import annotations

from dataclasses import asdict, dataclass
import json
from typing import Any


SCHEMA_VERSION = "employee-ai-context.v1"
SOURCE_MODES = {"jira", "clickup", "combined"}


class ContextValidationError(ValueError):
    """Raised when a context cannot safely satisfy the public contract."""


@dataclass(frozen=True)
class EmployeeAIContext:
    """Transport-safe context produced after a successful Employee analysis."""

    schema_version: str
    request: dict[str, Any]
    analysis: dict[str, Any]
    dashboard: dict[str, Any]
    evidence: dict[str, Any]
    permissions: dict[str, Any]
    data_quality: dict[str, Any]

    def to_dict(self) -> dict[str, Any]:
        payload = asdict(self)
        validate_employee_ai_context(payload)
        return payload

    def to_json(self) -> str:
        return json.dumps(
            self.to_dict(),
            ensure_ascii=False,
            sort_keys=True,
            separators=(",", ":"),
        )


def validate_employee_ai_context(payload: dict[str, Any]) -> None:
    """Validate the stable public envelope before it leaves the application."""

    required = {
        "schema_version",
        "request",
        "analysis",
        "dashboard",
        "evidence",
        "permissions",
        "data_quality",
    }
    if set(payload) != required:
        missing = sorted(required - set(payload))
        extra = sorted(set(payload) - required)
        raise ContextValidationError(
            f"Invalid Employee AI context keys; missing={missing}, extra={extra}"
        )
    if payload["schema_version"] != SCHEMA_VERSION:
        raise ContextValidationError("Unsupported Employee AI context schema version")
    analysis = payload.get("analysis") or {}
    if analysis.get("analysis_type") != "employee":
        raise ContextValidationError("Only Employee analysis is supported in V1")
    if analysis.get("source_mode") not in SOURCE_MODES:
        raise ContextValidationError("Unsupported Employee analysis source mode")
    request = payload.get("request") or {}
    if not request.get("context_fingerprint"):
        raise ContextValidationError("Context fingerprint is required")
    permissions = payload.get("permissions") or {}
    if permissions.get("mode") != "read_only" or permissions.get("source_access") != "none":
        raise ContextValidationError("Employee AI Assistant V1 must remain read-only")
