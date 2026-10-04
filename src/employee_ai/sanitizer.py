"""Allowlist helpers and final sensitive-data guard for AI context payloads."""

from __future__ import annotations

from dataclasses import asdict, is_dataclass
from datetime import date, datetime
from enum import Enum
import math
import re
from typing import Any, Mapping

from .schemas import ContextValidationError


_FORBIDDEN_KEY_TOKENS = {
    "email",
    "password",
    "passwd",
    "token",
    "secret",
    "apikey",
    "authorization",
    "cookie",
    "credential",
    "accountid",
    "jiraaccountid",
    "clickupuserid",
    "sessionstate",
    "rawresponse",
    "comment",
    "comments",
    "description",
    "attachment",
    "attachments",
}
_EMAIL_PATTERN = re.compile(r"(?<![\w.+-])[\w.+-]+@[\w.-]+\.[A-Za-z]{2,}(?![\w.-])")
_SECRET_QUERY_PATTERN = re.compile(
    r"(?i)([?&](?:token|access_token|api_key|apikey|password|secret)=)[^&#\s]+"
)


def _normalised_key(value: object) -> str:
    return re.sub(r"[^a-z0-9]", "", str(value).casefold())


def is_forbidden_key(key: object) -> bool:
    normalised = _normalised_key(key)
    return any(token in normalised for token in _FORBIDDEN_KEY_TOKENS)


def redact_sensitive_text(value: str) -> str:
    """Redact accidental email addresses or secrets embedded in safe text fields."""

    value = _EMAIL_PATTERN.sub("[redacted-email]", value)
    return _SECRET_QUERY_PATTERN.sub(r"\1[redacted]", value)


def transport_value(value: Any) -> Any:
    """Convert common analysis values into deterministic JSON-safe values."""

    if value is None or isinstance(value, (bool, int, str)):
        return redact_sensitive_text(value) if isinstance(value, str) else value
    if isinstance(value, float):
        return None if math.isnan(value) or math.isinf(value) else value
    if isinstance(value, (date, datetime)):
        return value.isoformat()
    if isinstance(value, Enum):
        return transport_value(value.value)
    if is_dataclass(value):
        return transport_value(asdict(value))
    if isinstance(value, Mapping):
        return {str(key): transport_value(item) for key, item in value.items()}
    if isinstance(value, (list, tuple, set, frozenset)):
        return [transport_value(item) for item in value]
    if hasattr(value, "item"):
        try:
            return transport_value(value.item())
        except (TypeError, ValueError):
            pass
    text = str(value)
    if text.casefold() in {"nan", "nat", "none", "null"}:
        return None
    return redact_sensitive_text(text)


def allowlisted_record(record: Mapping[str, Any], mapping: Mapping[str, str]) -> dict[str, Any]:
    """Copy only explicitly mapped source fields into a public record."""

    output: dict[str, Any] = {}
    for public_name, source_name in mapping.items():
        if source_name in record:
            output[public_name] = transport_value(record.get(source_name))
    return output


def assert_no_sensitive_fields(value: Any, path: str = "context") -> None:
    """Fail closed if a forbidden key or unredacted email escaped the allowlist."""

    if isinstance(value, Mapping):
        for key, item in value.items():
            if is_forbidden_key(key):
                raise ContextValidationError(f"Sensitive field is not allowed: {path}.{key}")
            assert_no_sensitive_fields(item, f"{path}.{key}")
        return
    if isinstance(value, (list, tuple)):
        for index, item in enumerate(value):
            assert_no_sensitive_fields(item, f"{path}[{index}]")
        return
    if isinstance(value, str) and _EMAIL_PATTERN.search(value):
        raise ContextValidationError(f"Unredacted email is not allowed: {path}")
