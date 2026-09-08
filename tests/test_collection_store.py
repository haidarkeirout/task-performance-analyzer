import unittest
from unittest.mock import Mock

import requests

from collection_store import CollectionStore, PersistenceError, persistence_owner_key
from test_automation import SETTINGS


class FakeResponse:
    def __init__(self, status=200, payload=None, headers=None):
        self.status_code = status
        self._payload = payload
        self.headers = headers or {}
        self.content = b"x" if payload is not None else b""

    def json(self):
        if isinstance(self._payload, Exception):
            raise self._payload
        return self._payload


class SequenceSession:
    def __init__(self, outcomes):
        self.outcomes = list(outcomes)
        self.headers = {}
        self.calls = 0

    def post(self, *args, **kwargs):
        self.calls += 1
        outcome = self.outcomes.pop(0)
        if isinstance(outcome, Exception):
            raise outcome
        return outcome

    def close(self):
        pass


class StoreTests(unittest.TestCase):
    def test_owner_key_is_stable_and_secret_derived(self):
        first = persistence_owner_key(SETTINGS)
        second = persistence_owner_key(SETTINGS)
        self.assertEqual(first, second)
        self.assertEqual(len(first), 64)
        self.assertNotIn(SETTINGS.username, first)

    def test_network_failure_retries_before_success(self):
        session = SequenceSession([
            requests.ConnectionError("one"),
            requests.Timeout("two"),
            FakeResponse(200, {"ok": True}),
        ])
        sleeps = []
        store = CollectionStore("https://example.invalid", "key", session=session, sleep=sleeps.append)
        result = store.latest_resumable("x" * 64)
        self.assertEqual(result, {"ok": True})
        self.assertEqual(session.calls, 3)
        self.assertEqual(sleeps, [1, 2])

    def test_rate_limit_and_server_error_retry(self):
        session = SequenceSession([
            FakeResponse(429, {"message": "slow"}, headers={"Retry-After": "0"}),
            FakeResponse(503, {"message": "down"}),
            FakeResponse(200, None),
        ])
        sleeps = []
        store = CollectionStore("https://example.invalid", "key", session=session, sleep=sleeps.append)
        self.assertIsNone(store.latest_resumable("x" * 64))
        self.assertEqual(session.calls, 3)
        self.assertEqual(len(sleeps), 2)

    def test_permission_error_fails_without_retry(self):
        session = SequenceSession([FakeResponse(403, {"message": "denied"})])
        store = CollectionStore("https://example.invalid", "key", session=session, sleep=Mock())
        with self.assertRaises(PersistenceError) as raised:
            store.latest_resumable("x" * 64)
        self.assertFalse(raised.exception.retryable)
        self.assertEqual(session.calls, 1)


if __name__ == "__main__":
    unittest.main()
