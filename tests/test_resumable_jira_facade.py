from __future__ import annotations

from copy import deepcopy
from types import SimpleNamespace
import unittest
from unittest.mock import patch

from jira_gateway import CollectionError
from resumable_jira import collect_jira_query


class FakeStore:
    def __init__(self):
        self.jobs = {}

    def close(self):
        pass

    def begin(self, owner_key, job_id, fingerprint, space, query, definitions, cutoff):
        self.jobs.setdefault(job_id, {
            "job_id": job_id, "fingerprint": fingerprint, "space": deepcopy(space),
            "query": query, "definitions": deepcopy(definitions), "cutoff": cutoff,
            "items": [], "status": "running", "stage": "Ready",
        })

    def seed_items(self, owner_key, job_id, items):
        job = self.jobs[job_id]
        if not job["items"]:
            job["items"] = [
                {"position": index, "issue_key": item["key"], "seed_item": deepcopy(item),
                 "completed": False, "item_data": None, "history_data": None}
                for index, item in enumerate(items)
            ]

    def save_item(self, owner_key, job_id, issue_key, item, history):
        row = next(row for row in self.jobs[job_id]["items"] if row["issue_key"] == issue_key)
        row.update(completed=True, item_data=deepcopy(item), history_data=deepcopy(history))

    def update_job(self, owner_key, job_id, status, stage, current_issue, message,
                   error_detail=None, result_meta=None):
        self.jobs[job_id].update(
            status=status, stage=stage, current_issue=current_issue, message=message,
            error_detail=deepcopy(error_detail), result_meta=deepcopy(result_meta),
        )

    def latest_for_fingerprint(self, owner_key, fingerprint):
        matches = [job for job in self.jobs.values() if job["fingerprint"] == fingerprint]
        return deepcopy(matches[-1]) if matches else None


class FakeGateway:
    calls = {"ENG-1": 0, "ENG-2": 0}
    fail_second_once = True

    def __init__(self, settings):
        self.progress = None

    def close(self):
        pass

    def fields(self):
        return []

    def all_issues(self, query, progress=None):
        return [self._issue("ENG-1"), self._issue("ENG-2")]

    def project_details(self, project_id):
        return {"id": "10", "key": "ENG", "name": "Engineering"}

    @staticmethod
    def _issue(key):
        return {
            "id": key.split("-")[-1], "key": key,
            "fields": {
                "summary": key, "project": {"id": "10", "key": "ENG", "name": "Engineering"},
                "status": {"name": "To Do"}, "created": "2026-09-01T09:00:00Z",
            },
        }

    def complete_issue(self, item):
        key = item["key"]
        self.calls[key] += 1
        if key == "ENG-2" and self.fail_second_once:
            self.__class__.fail_second_once = False
            raise CollectionError("Temporary test interruption.", retryable=False)
        return self._issue(key), {
            "issue_key": key, "history_complete": True,
            "history_through": "2026-09-30T20:59:59Z",
            "initial_status": "To Do", "status_events": [],
        }


class ResumableJiraFacadeTests(unittest.TestCase):
    def setUp(self):
        FakeGateway.calls = {"ENG-1": 0, "ENG-2": 0}
        FakeGateway.fail_second_once = True
        self.store = FakeStore()
        self.settings = SimpleNamespace(
            username="user", password="secret", password_hash="", source_timezone="Asia/Damascus",
            start_date_field="", jira_url="https://example.atlassian.net",
            jira_email="a@example.com", jira_token="token", jira_cloud_id="",
        )

    def test_retry_resumes_after_the_last_persisted_work_item(self):
        with patch("resumable_jira.CollectionStore.configured", return_value=self.store), \
             patch("resumable_jira.JiraGateway", FakeGateway):
            with self.assertRaises(CollectionError):
                collect_jira_query(
                    self.settings,
                    space={"id": "10", "key": "ENG", "name": "Engineering"},
                    query='project = "ENG"',
                    fingerprint="project:eng",
                    cutoff="2026-09-30T20:59:59Z",
                )
            result = collect_jira_query(
                self.settings,
                space={"id": "10", "key": "ENG", "name": "Engineering"},
                query='project = "ENG"',
                fingerprint="project:eng",
                cutoff="2026-09-30T20:59:59Z",
            )

        self.assertEqual(result.count, 2)
        self.assertEqual(FakeGateway.calls["ENG-1"], 1)
        self.assertEqual(FakeGateway.calls["ENG-2"], 2)


if __name__ == "__main__":
    unittest.main()
