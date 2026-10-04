"""Small, fail-closed client for the Employee AI Activepieces webhook."""

from __future__ import annotations

from copy import deepcopy
import json
from typing import Any, Mapping
from urllib.error import HTTPError, URLError
from urllib.parse import urlparse
from urllib.request import Request, urlopen

from .schemas import EmployeeAIContext, validate_employee_ai_context


WEBHOOK_SECRET_NAME = "ACTIVEPIECES_EMPLOYEE_AI_WEBHOOK_URL"
# A synchronous Activepieces flow includes an external AI call.  Thirty seconds
# is often too short even when the webhook and the response contract are valid.
DEFAULT_TIMEOUT_SECONDS = 75
MAX_QUESTION_LENGTH = 2_000
CLIENT_USER_AGENT = "Performance-Management-Employee-AI/1.0"


class EmployeeAIWebhookError(RuntimeError):
    """Raised when the external AI assistant cannot safely answer a request."""


def configured_webhook_url(secrets: Mapping[str, Any] | None) -> str | None:
    """Read a production webhook URL without ever putting it in source control."""

    value = str((secrets or {}).get(WEBHOOK_SECRET_NAME, "") or "").strip()
    if not value:
        return None
    parsed = urlparse(value)
    if parsed.scheme != "https" or not parsed.netloc or parsed.username or parsed.password:
        raise EmployeeAIWebhookError("The Employee AI webhook URL must be a valid HTTPS URL")
    return value


def _question_text(question: Any) -> str:
    text = str(question or "").strip()
    if not text:
        raise EmployeeAIWebhookError("Please enter a question about the current analysis")
    if len(text) > MAX_QUESTION_LENGTH:
        raise EmployeeAIWebhookError("The question is too long. Please keep it under 2,000 characters")
    return text


def _request_payload(context: EmployeeAIContext, question: Any) -> tuple[dict[str, Any], str, str]:
    """Attach one user question without changing the analysis fingerprint."""

    text = _question_text(question)
    payload_context = deepcopy(context.to_dict())
    payload_context["request"]["question"] = text
    validate_employee_ai_context(payload_context)
    request = payload_context["request"]
    request_id = str(request["request_id"])
    fingerprint = str(request["context_fingerprint"])
    return {
        "question": text,
        "context": payload_context,
        "request_id": request_id,
        "context_fingerprint": fingerprint,
    }, request_id, fingerprint


def ask_employee_ai(
    *,
    context: EmployeeAIContext,
    question: Any,
    webhook_url: str,
    timeout_seconds: int = DEFAULT_TIMEOUT_SECONDS,
) -> dict[str, str]:
    """Send only sanitized context and reject a mismatched webhook response."""

    url = configured_webhook_url({WEBHOOK_SECRET_NAME: webhook_url})
    if not url:
        raise EmployeeAIWebhookError("The Employee AI assistant is not configured")
    payload, request_id, expected_fingerprint = _request_payload(context, question)
    try:
        request = Request(
            url,
            data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
            headers={
                "Content-Type": "application/json",
                "Accept": "application/json",
                "User-Agent": CLIENT_USER_AGENT,
            },
            method="POST",
        )
        with urlopen(request, timeout=max(1, int(timeout_seconds))) as response:
            raw_body = response.read()
    except HTTPError as exc:
        raise EmployeeAIWebhookError(
            f"The AI assistant webhook returned HTTP {exc.code}"
        ) from exc
    except TimeoutError as exc:
        raise EmployeeAIWebhookError(
            "The AI assistant took too long to respond. Please try again."
        ) from exc
    except (URLError, OSError) as exc:
        raise EmployeeAIWebhookError("The AI assistant could not be reached") from exc
    try:
        body = json.loads(raw_body.decode("utf-8"))
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise EmployeeAIWebhookError("The AI assistant returned an invalid response") from exc
    if not isinstance(body, Mapping):
        raise EmployeeAIWebhookError("The AI assistant returned an invalid response")

    answer = body.get("answer")
    response_fingerprint = body.get("context_fingerprint")
    if not isinstance(answer, str) or not answer.strip():
        raise EmployeeAIWebhookError("The AI assistant returned no answer")
    if response_fingerprint != expected_fingerprint:
        raise EmployeeAIWebhookError("The AI assistant response does not match the current analysis")
    response_request_id = str(body.get("request_id") or request_id)
    if response_request_id != request_id:
        raise EmployeeAIWebhookError("The AI assistant response does not match the current request")
    return {
        "answer": answer.strip(),
        "request_id": request_id,
        "context_fingerprint": expected_fingerprint,
    }
