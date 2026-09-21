"""Read-only Jira Cloud access for the collection workflow.

Uses enhanced search cursors and independently pages histories/comments/worklogs.
No credentials, raw HTTP bodies, or authorization headers enter error messages.
"""
from __future__ import annotations

import re
import time
from datetime import datetime, timezone
from urllib.parse import quote

import requests

from automation_auth import Settings
from jira_client import JiraClient


class CollectionError(RuntimeError):
    def __init__(self, message, status=None, retryable=False):
        super().__init__(message)
        self.status = status
        self.retryable = retryable


class JiraGateway:
    def __init__(self, settings: Settings, session=None, sleep=time.sleep):
        self.settings = settings
        self.base = (f"https://api.atlassian.com/ex/jira/{settings.jira_cloud_id}"
                     if settings.jira_cloud_id else settings.jira_url)
        self.session = session or requests.Session()
        self.session.auth = (settings.jira_email, settings.jira_token)
        self.session.headers.update({"Accept": "application/json", "Content-Type": "application/json"})
        self.sleep = sleep
        self.progress = None

    def close(self):
        self.session.close()

    def request(self, method, path, *, params=None, body=None):
        if method not in {"GET", "POST"} or not path.startswith("/rest/api/3/"):
            raise CollectionError("The requested Jira operation is unavailable.")
        # POST is only permitted for documented read operations.
        if method == "POST" and path not in {
            "/rest/api/3/search/jql", "/rest/api/3/jql/autocompletedata", "/rest/api/3/jql/parse"
        }:
            raise CollectionError("The requested Jira operation is unavailable.")

        # Gateway-level retries absorb short network/rate-limit interruptions before
        # the CollectionJob has to retry an entire work item.
        for attempt in range(4):
            try:
                response = self.session.request(
                    method, self.base + path, params=params, json=body,
                    timeout=(10, 60), allow_redirects=False,
                )
            except requests.RequestException as exc:
                if attempt < 3:
                    if self.progress:
                        self.progress(f"Connection interrupted; retry {attempt + 1} of 3...")
                    self.sleep(2 ** attempt)
                    continue
                raise CollectionError(
                    "Could not reach Jira after repeated connection attempts.",
                    retryable=True,
                ) from None

            status = response.status_code
            if status == 429 or 500 <= status <= 599:
                if attempt < 3:
                    try:
                        delay = max(0.0, float(response.headers.get("Retry-After", 2 ** attempt)))
                    except (ValueError, TypeError):
                        delay = float(2 ** attempt)
                    if delay > 30:
                        raise CollectionError(
                            f"Jira is temporarily unavailable or rate-limited (HTTP {status}).",
                            status=status,
                            retryable=True,
                        )
                    if self.progress:
                        self.progress(
                            f"Jira HTTP {status}; retry {attempt + 1} of 3 in {delay:g} seconds..."
                        )
                    self.sleep(delay)
                    continue
                raise CollectionError(
                    f"Jira is temporarily unavailable or rate-limited (HTTP {status}).",
                    status=status,
                    retryable=True,
                )

            if status in (401, 403):
                raise CollectionError(
                    "Jira access could not be verified. Ask the administrator to check the connection and its permissions.",
                    status=status,
                )
            if status == 400:
                raise CollectionError(
                    "Jira could not apply these filters. Check the selected fields, operators, and values.",
                    status=status,
                )
            if status == 404:
                raise CollectionError(
                    "The selected Jira resource is unavailable or is no longer accessible.",
                    status=status,
                )
            if status < 200 or status >= 300:
                raise CollectionError(
                    f"Jira could not complete the request (HTTP {status}). Please try again.",
                    status=status,
                )
            try:
                return response.json()
            except ValueError:
                raise CollectionError(
                    "Jira returned an unreadable response. Please try again.",
                    status=status,
                    retryable=True,
                ) from None

    def _paged(self, path, key="values", params=None):
        start, result, ids, expected_total = 0, [], set(), None
        while True:
            page = self.request("GET", path, params={**(params or {}), "startAt": start, "maxResults": 100})
            values = page.get(key)
            if not isinstance(values, list):
                raise CollectionError("Jira returned an incomplete page. No data has been marked ready.")
            if page.get("startAt", start) != start:
                raise CollectionError(
                    "Jira repeated or skipped a page. Please collect the data again.",
                    retryable=True,
                )
            total = page.get("total")
            if total is not None:
                total = int(total)
                if expected_total is not None and total != expected_total:
                    raise CollectionError("Jira data changed during collection. Please start a new collection.")
                expected_total = total
            for value in values:
                identity = value.get("id")
                if identity is not None:
                    identity = str(identity)
                    if identity in ids:
                        raise CollectionError("Jira repeated a record during collection. Please try again.")
                    ids.add(identity)
            result.extend(values)
            start += len(values)
            if total is not None and start > total:
                raise CollectionError("Jira returned inconsistent page totals. Please try again.")
            if page.get("isLast") is True or (total is not None and start == total):
                if total is not None and start != total:
                    raise CollectionError(
                        "Jira ended before all records were collected. Please try again.",
                        retryable=True,
                    )
                return result
            if not values:
                if (total is not None and start < total) or page.get("isLast") is False:
                    raise CollectionError(
                        "Jira ended before all records were collected. Please try again.",
                        retryable=True,
                    )
                return result
            if total is None and "isLast" not in page and len(values) < page.get("maxResults", 100):
                return result

    def identity(self):
        return self.request("GET", "/rest/api/3/myself")

    def spaces(self):
        return self._paged("/rest/api/3/project/search", params={"orderBy": "name"})

    def project_details(self, project_id):
        return self.request("GET", "/rest/api/3/project/" + quote(str(project_id), safe=""))

    def fields(self):
        data = self.request("GET", "/rest/api/3/field")
        if not isinstance(data, list):
            raise CollectionError("Jira field information is unavailable.")
        return data

    def filter_reference(self, project_id):
        return self.request(
            "POST", "/rest/api/3/jql/autocompletedata",
            body={"projectIds": [int(project_id)], "includeCollapsedFields": False},
        )

    def suggestions(self, field, value=""):
        return self.request(
            "GET", "/rest/api/3/jql/autocompletedata/suggestions",
            params={"fieldName": field, "fieldValue": value},
        ).get("results", [])

    def validate(self, jql):
        data = self.request(
            "POST", "/rest/api/3/jql/parse", params={"validation": "strict"},
            body={"queries": [jql]},
        )
        queries = data.get("queries", [])
        if not queries or queries[0].get("errors"):
            raise CollectionError(
                "Some filters are not valid for this Jira space. Check their values or review the JQL query."
            )

    def search_page(self, jql, token=None, all_fields=False):
        body = {
            "jql": jql,
            "maxResults": 100 if all_fields else 50,
            "fields": ["*all"] if all_fields else [
                "summary", "assignee", "issuetype", "status", "priority",
                "resolution", "created", "duedate", "project",
            ],
        }
        if token:
            body["nextPageToken"] = token
        if all_fields:
            body["expand"] = "names,schema"
        page = self.request("POST", "/rest/api/3/search/jql", body=body)
        if not isinstance(page.get("issues"), list):
            raise CollectionError("Jira did not return a valid work item list.")
        if page.get("isLast") is not True and not page.get("nextPageToken"):
            raise CollectionError(
                "Jira did not provide the next page. No partial export was prepared.",
                retryable=True,
            )
        return page

    def all_issues(self, jql, progress=None):
        self.validate(jql)
        items, seen_keys, tokens, token = [], set(), set(), None
        while True:
            page = self.search_page(jql, token, all_fields=True)
            for item in page["issues"]:
                key = item.get("key")
                if not key or key in seen_keys:
                    raise CollectionError("Jira returned a missing or repeated work item. Please collect again.")
                seen_keys.add(key)
                items.append(item)
            if progress:
                progress(f"Collecting work items: {len(items)} received...")
            if page.get("isLast") is True:
                return items
            token = page["nextPageToken"]
            if token in tokens:
                raise CollectionError(
                    "Jira repeated a search page. Please collect again.",
                    retryable=True,
                )
            tokens.add(token)

    @staticmethod
    def _issue_path(key):
        if not re.fullmatch(r"[A-Za-z][A-Za-z0-9_]*-\d+", key):
            raise CollectionError("Jira returned an invalid work item key.")
        return "/rest/api/3/issue/" + quote(key, safe="")

    def complete_issue(self, item):
        """Verify a stable snapshot around each full history fetch; retry one edit race."""
        path = self._issue_path(item["key"])
        for attempt in range(2):
            if self.progress:
                self.progress(f"{item['key']}: reading task fields (attempt {attempt + 1})...")
            snapshot = self.request("GET", path, params={"fields": "*all", "expand": "names,schema"})
            fields = snapshot.get("fields", {})
            if self.progress:
                self.progress(f"{item['key']}: reading complete transition history...")
            changes = self._paged(path + "/changelog")
            for name, endpoint, list_key in [
                ("comment", "comment", "comments"),
                ("worklog", "worklog", "worklogs"),
            ]:
                embedded = fields.get(name)
                if isinstance(embedded, dict):
                    if self.progress:
                        self.progress(f"{item['key']}: reading {name} pages...")
                    full = self._paged(path + "/" + endpoint, key=list_key)
                    fields[name] = {"startAt": 0, "total": len(full), list_key: full}
            watches = fields.get("watches")
            if isinstance(watches, dict) and watches.get("watchCount", 0):
                if self.progress:
                    self.progress(f"{item['key']}: reading watcher details...")
                try:
                    fields["watches"] = self.request("GET", path + "/watchers")
                except CollectionError as exc:
                    if exc.status not in (403, 404):
                        raise
                    # Jira can allow issue reading but deny the watcher identities.
                    fields["watches"] = {**watches, "identities_unavailable": True}

            # The final consistency read must cover events received with the snapshot.
            through = datetime.now(timezone.utc).isoformat()
            check = self.request("GET", path, params={"fields": "updated,status"}).get("fields", {})
            if check.get("updated") != fields.get("updated") or check.get("status") != fields.get("status"):
                continue

            status_events = JiraClient._extract_status_events(item["key"], changes)
            current = (fields.get("status") or {}).get("name")
            if not current:
                raise CollectionError("A work item has no readable status. No data has been marked ready.")
            history = {
                "issue_key": item["key"],
                "snapshot": {"current_status": current},
                "history_complete": True,
                "history_through": through,
                "initial_status": status_events[0].get("from_status") if status_events else current,
                "status_events": status_events,
                "assignee_events": JiraClient._extract_assignee_events(item["key"], changes),
                "raw_changelog": changes,
            }
            return snapshot, history

        raise CollectionError(
            "A work item changed while it was being collected. Please retry the collection.",
            retryable=True,
        )
