"""Deterministic, source-neutral Employee analysis identity helpers."""

from __future__ import annotations

import hashlib
import hmac
import json
from typing import Any

from .sanitizer import transport_value


def canonical_json(value: Any) -> str:
    return json.dumps(
        transport_value(value),
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
    )


def stable_subject_hash(subject_key: str, secret: str) -> str:
    """Hide the internal employee key while retaining stable identity."""

    if not str(subject_key).strip():
        raise ValueError("Employee subject key is required")
    if not str(secret):
        raise ValueError("Employee subject hash secret is required")
    return hmac.new(
        str(secret).encode("utf-8"),
        str(subject_key).strip().encode("utf-8"),
        hashlib.sha256,
    ).hexdigest()


def context_fingerprint(identity: Any) -> str:
    """Fingerprint the current analysis scope without embedding source secrets."""

    return hashlib.sha256(canonical_json(identity).encode("utf-8")).hexdigest()
