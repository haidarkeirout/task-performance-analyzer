"""Read-only Employee AI context contract.

This package never collects source data and never calculates performance
metrics.  It only adapts an already-rendered Employee analysis into a small,
sanitised payload for the future assistant integration.
"""

from .context_builder import (
    build_clickup_employee_context,
    build_combined_employee_context,
    build_employee_ai_context,
    build_jira_employee_context,
)
from .client import EmployeeAIWebhookError, ask_employee_ai, configured_webhook_url
from .schemas import EmployeeAIContext, SCHEMA_VERSION

__all__ = [
    "EmployeeAIContext",
    "SCHEMA_VERSION",
    "build_employee_ai_context",
    "build_jira_employee_context",
    "build_clickup_employee_context",
    "build_combined_employee_context",
    "EmployeeAIWebhookError",
    "ask_employee_ai",
    "configured_webhook_url",
]
