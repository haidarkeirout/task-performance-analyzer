import unittest
from unittest.mock import patch

from employee_ai import build_employee_ai_context
from employee_ai.client import (
    EmployeeAIWebhookError,
    WEBHOOK_SECRET_NAME,
    ask_employee_ai,
    configured_webhook_url,
)


CONTEXT = build_employee_ai_context(
    source_mode="jira",
    employee={"display_name": "Test Employee", "subject_key": "employee-1"},
    collection={"collection_id": "collection-1"},
    cards=[],
    tasks=[],
    subject_secret="test-secret",
    request_id="request-1",
    sent_at="2026-10-04T12:00:00+00:00",
)
WEBHOOK_URL = "https://cloud.activepieces.com/api/v1/webhooks/test/sync"


class _Response:
    def __init__(self, body):
        self.body = body

    def __enter__(self):
        return self

    def __exit__(self, *_):
        return None

    def read(self):
        import json
        return json.dumps(self.body).encode("utf-8")


class EmployeeAIClientTests(unittest.TestCase):
    def test_configured_webhook_url_requires_https(self):
        self.assertEqual(
            configured_webhook_url({WEBHOOK_SECRET_NAME: WEBHOOK_URL}),
            WEBHOOK_URL,
        )
        with self.assertRaises(EmployeeAIWebhookError):
            configured_webhook_url({WEBHOOK_SECRET_NAME: "http://example.test/webhook"})

    @patch("employee_ai.client.urlopen")
    def test_posts_only_sanitized_context_and_requires_matching_identity(self, post):
        fingerprint = CONTEXT.to_dict()["request"]["context_fingerprint"]
        post.return_value = _Response({
            "answer": "The supplied analysis shows 1 task.",
            "request_id": "request-1",
            "context_fingerprint": fingerprint,
        })
        result = ask_employee_ai(context=CONTEXT, question="Explain the result", webhook_url=WEBHOOK_URL)

        self.assertEqual(result["answer"], "The supplied analysis shows 1 task.")
        request = post.call_args.args[0]
        import json
        sent = json.loads(request.data.decode("utf-8"))
        self.assertEqual(sent["question"], "Explain the result")
        self.assertEqual(sent["context"]["request"]["question"], "Explain the result")
        self.assertEqual(sent["context_fingerprint"], fingerprint)

    @patch("employee_ai.client.urlopen")
    def test_rejects_response_for_another_analysis(self, post):
        post.return_value = _Response({
            "answer": "Wrong analysis.",
            "request_id": "request-1",
            "context_fingerprint": "wrong-fingerprint",
        })
        with self.assertRaisesRegex(EmployeeAIWebhookError, "does not match the current analysis"):
            ask_employee_ai(context=CONTEXT, question="Explain", webhook_url=WEBHOOK_URL)

    def test_rejects_empty_question_before_network_call(self):
        with self.assertRaisesRegex(EmployeeAIWebhookError, "Please enter a question"):
            ask_employee_ai(context=CONTEXT, question="   ", webhook_url=WEBHOOK_URL)
