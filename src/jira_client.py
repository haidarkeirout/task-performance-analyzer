"""
Secure Jira REST API client.

Environment variables:
    JIRA_BASE_URL=https://your-domain.atlassian.net
    JIRA_EMAIL=your-email@example.com
    JIRA_API_TOKEN=your-api-token
"""

from __future__ import annotations

import argparse
import json
import logging
import os
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import requests


LOGGER = logging.getLogger(__name__)


class JiraApiError(RuntimeError):
    """Raised when Jira returns an unsuccessful response."""


@dataclass
class JiraClient:
    base_url: str
    email: str
    api_token: str
    timeout_seconds: int = 30
    max_retries: int = 3

    def __post_init__(self) -> None:
        self.base_url = self.base_url.rstrip("/")

        if not self.base_url.startswith(("https://", "http://")):
            raise ValueError("JIRA_BASE_URL must start with http:// or https://")

        if not self.email.strip():
            raise ValueError("Jira email is required.")

        if not self.api_token.strip():
            raise ValueError("Jira API token is required.")

        self.session = requests.Session()
        self.session.auth = (self.email, self.api_token)
        self.session.headers.update(
            {
                "Accept": "application/json",
                "Content-Type": "application/json",
                "User-Agent": "task-performance-analyzer/1.0",
            }
        )

    @classmethod
    def from_environment(cls) -> "JiraClient":
        base_url = os.getenv("JIRA_BASE_URL", "").strip()
        email = os.getenv("JIRA_EMAIL", "").strip()
        api_token = os.getenv("JIRA_API_TOKEN", "").strip()

        missing = []

        if not base_url:
            missing.append("JIRA_BASE_URL")

        if not email:
            missing.append("JIRA_EMAIL")

        if not api_token:
            missing.append("JIRA_API_TOKEN")

        if missing:
            raise ValueError(
                "Missing Jira environment variables: "
                + ", ".join(missing)
            )

        return cls(
            base_url=base_url,
            email=email,
            api_token=api_token,
        )

    def _request(
        self,
        method: str,
        path: str,
        *,
        params: dict[str, Any] | None = None,
    ) -> requests.Response:
        url = f"{self.base_url}{path}"

        for attempt in range(self.max_retries + 1):
            response = self.session.request(
                method=method,
                url=url,
                params=params,
                timeout=self.timeout_seconds,
            )

            if response.status_code == 429:
                retry_after = response.headers.get("Retry-After", "5")

                try:
                    delay = min(float(retry_after), 60.0)
                except ValueError:
                    delay = 5.0

                if attempt < self.max_retries:
                    LOGGER.warning(
                        "Jira rate limit reached. Retrying in %.1f seconds.",
                        delay,
                    )
                    time.sleep(delay)
                    continue

            if 500 <= response.status_code <= 599:
                if attempt < self.max_retries:
                    delay = 2**attempt
                    LOGGER.warning(
                        "Jira server error %s. Retrying in %s seconds.",
                        response.status_code,
                        delay,
                    )
                    time.sleep(delay)
                    continue

            if not response.ok:
                try:
                    detail = response.json()
                except ValueError:
                    detail = response.text[:500]

                raise JiraApiError(
                    f"Jira API request failed: "
                    f"{response.status_code} {method} {path} - {detail}"
                )

            return response

        raise JiraApiError(f"Jira request failed after retries: {method} {path}")

    def verify_connection(self) -> dict[str, Any]:
        response = self._request("GET", "/rest/api/3/myself")
        return response.json()

    def get_issue_snapshot(self, issue_key: str) -> dict[str, Any]:
        response = self._request(
            "GET",
            f"/rest/api/3/issue/{issue_key}",
            params={
                "fields": (
                    "summary,status,issuetype,assignee,priority,"
                    "created,duedate,labels,resolution"
                )
            },
        )

        payload = response.json()
        fields = payload.get("fields", {})

        assignee = fields.get("assignee") or {}
        issue_type = fields.get("issuetype") or {}
        status = fields.get("status") or {}
        priority = fields.get("priority") or {}

        return {
            "issue_key": payload.get("key", issue_key),
            "issue_id": payload.get("id"),
            "task_name": fields.get("summary"),
            "current_status": status.get("name"),
            "issue_type": issue_type.get("name"),
            "assignee_id": assignee.get("accountId"),
            "assignee_name": assignee.get("displayName"),
            "priority": priority.get("name"),
            "created_at": fields.get("created"),
            "due_date": fields.get("duedate"),
            "labels": fields.get("labels") or [],
            "resolution": (
                (fields.get("resolution") or {}).get("name")
                if fields.get("resolution")
                else None
            ),
            "source": "jira_api_current_snapshot",
        }

    def get_issue_changelog(
        self,
        issue_key: str,
        *,
        page_size: int = 100,
    ) -> list[dict[str, Any]]:
        all_changes: list[dict[str, Any]] = []
        start_at = 0

        while True:
            response = self._request(
                "GET",
                f"/rest/api/3/issue/{issue_key}/changelog",
                params={
                    "startAt": start_at,
                    "maxResults": page_size,
                },
            )

            payload = response.json()
            values = payload.get("values", [])

            if not values:
                break

            all_changes.extend(values)

            start_at += len(values)
            total = payload.get("total")

            if total is not None and start_at >= int(total):
                break

            if len(values) < page_size:
                break

        return all_changes

    @staticmethod
    def _extract_status_events(
        issue_key: str,
        changelog: list[dict[str, Any]],
    ) -> list[dict[str, Any]]:
        events: list[dict[str, Any]] = []

        for change in changelog:
            changed_at = change.get("created")
            change_id = change.get("id")
            author = change.get("author") or {}

            for item in change.get("items", []):
                if str(item.get("field", "")).casefold() != "status":
                    continue

                events.append(
                    {
                        "issue_key": issue_key,
                        "from_status": item.get("fromString"),
                        "to_status": item.get("toString"),
                        "changed_at": changed_at,
                        "change_id": change_id,
                        "author_id": author.get("accountId"),
                        "author_name": author.get("displayName"),
                        "source": "jira_api_changelog",
                    }
                )

        return events

    @staticmethod
    def _extract_assignee_events(
        issue_key: str,
        changelog: list[dict[str, Any]],
    ) -> list[dict[str, Any]]:
        events: list[dict[str, Any]] = []

        for change in changelog:
            changed_at = change.get("created")
            change_id = change.get("id")
            author = change.get("author") or {}

            for item in change.get("items", []):
                field_name = str(item.get("field", "")).casefold()

                if field_name not in {"assignee", "assignee field"}:
                    continue

                events.append(
                    {
                        "issue_key": issue_key,
                        "from_assignee": item.get("fromString"),
                        "to_assignee": item.get("toString"),
                        "from_assignee_id": item.get("from"),
                        "to_assignee_id": item.get("to"),
                        "changed_at": changed_at,
                        "change_id": change_id,
                        "author_id": author.get("accountId"),
                        "author_name": author.get("displayName"),
                        "source": "jira_api_changelog",
                    }
                )

        return events

    def fetch_issue_history(self, issue_key: str) -> dict[str, Any]:
        snapshot = self.get_issue_snapshot(issue_key)
        changelog = self.get_issue_changelog(issue_key)

        return {
            "issue_key": issue_key,
            "snapshot": snapshot,
            "status_events": self._extract_status_events(
                issue_key,
                changelog,
            ),
            "assignee_events": self._extract_assignee_events(
                issue_key,
                changelog,
            ),
            "raw_changelog": changelog,
        }

    def fetch_many_issue_histories(
        self,
        issue_keys: list[str],
    ) -> dict[str, dict[str, Any]]:
        results: dict[str, dict[str, Any]] = {}

        unique_keys = list(
            dict.fromkeys(
                key.strip()
                for key in issue_keys
                if key and key.strip()
            )
        )

        for issue_key in unique_keys:
            LOGGER.info("Fetching Jira history for %s", issue_key)

            try:
                results[issue_key] = self.fetch_issue_history(issue_key)
            except JiraApiError:
                LOGGER.exception(
                    "Could not fetch Jira history for %s",
                    issue_key,
                )
                results[issue_key] = {
                    "issue_key": issue_key,
                    "error": "history_fetch_failed",
                    "snapshot": None,
                    "status_events": [],
                    "assignee_events": [],
                    "raw_changelog": [],
                }

        return results


def save_history(
    history: dict[str, dict[str, Any]],
    output_path: Path,
) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with output_path.open("w", encoding="utf-8") as file:
        json.dump(
            history,
            file,
            ensure_ascii=False,
            indent=2,
        )


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Fetch Jira issue snapshots and changelog history."
    )

    parser.add_argument(
        "--issue-keys",
        required=True,
        help="Comma-separated Jira issue keys, for example WEB-1,WEB-2.",
    )

    parser.add_argument(
        "--output",
        required=True,
        help="Path for the JSON history output.",
    )

    return parser


def main() -> int:
    logging.basicConfig(
        level=logging.INFO,
        format="%(levelname)s: %(message)s",
    )

    parser = build_parser()
    args = parser.parse_args()

    try:
        client = JiraClient.from_environment()

        user = client.verify_connection()
        LOGGER.info(
            "Connected to Jira as %s",
            user.get("displayName", "authenticated user"),
        )

        issue_keys = args.issue_keys.split(",")
        history = client.fetch_many_issue_histories(issue_keys)

        save_history(
            history,
            Path(args.output),
        )

        print(f"Fetched Jira history for {len(history)} issue(s).")
        print(f"Saved output to: {args.output}")

        return 0

    except (ValueError, JiraApiError) as exc:
        LOGGER.error(str(exc))
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
