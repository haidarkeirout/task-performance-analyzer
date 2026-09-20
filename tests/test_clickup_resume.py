from __future__ import annotations

import unittest

from clickup_gateway import ClickUpCollectionError, ClickUpGateway


class ResumableClickUpGateway(ClickUpGateway):
    def __init__(self):
        self.calls = []
        self.fail_page_one_once = True

    def folderless_lists(self, space_id, archived=False):
        return [{"id": "list-1", "name": "List"}]

    def folders(self, space_id, archived=False):
        return []

    def tasks(self, list_id, page=0):
        self.calls.append(page)
        if page == 0:
            return {"tasks": [{"id": f"T-{index}"} for index in range(100)]}
        if self.fail_page_one_once:
            self.fail_page_one_once = False
            raise ClickUpCollectionError("Temporary interruption")
        return {"tasks": [{"id": "T-100"}]}


class ClickUpResumeTests(unittest.TestCase):
    def test_space_collection_resumes_from_the_saved_page(self):
        gateway = ResumableClickUpGateway()
        checkpoint = {}
        with self.assertRaises(ClickUpCollectionError):
            gateway.all_tasks_for_space(
                "space-1",
                checkpoint=checkpoint,
                checkpoint_callback=lambda state: None,
            )
        result = gateway.all_tasks_for_space(
            "space-1",
            checkpoint=checkpoint,
            checkpoint_callback=lambda state: None,
        )
        self.assertEqual(len(result), 101)
        self.assertEqual(gateway.calls, [0, 1, 1])
        self.assertTrue(checkpoint["list_states"]["list-1"]["complete"])

    def test_fresh_space_collection_replaces_stale_snapshot(self):
        gateway = ResumableClickUpGateway()
        gateway.fail_page_one_once = False
        checkpoint = {
            "lists": [{"id": "list-1", "name": "List"}],
            "tasks": {"OLD": {"id": "OLD", "status": {"status": "in progress"}}},
            "list_states": {"list-1": {"complete": True, "next_page": 0}},
        }
        result = gateway.all_tasks_for_space(
            "space-1",
            checkpoint=checkpoint,
            checkpoint_callback=lambda state: None,
            fresh=True,
        )
        self.assertNotIn("OLD", {task["id"] for task in result})
        self.assertEqual(len(result), 101)
        self.assertTrue(checkpoint["list_states"]["list-1"]["complete"])


if __name__ == "__main__":
    unittest.main()
