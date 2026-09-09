"""Read-only ClickUp API v2 access for the collection workflow."""
from __future__ import annotations

import os
import time
import requests

from clickup_browser_history import ClickUpBrowserHistory, ClickUpBrowserHistoryError


class ClickUpCollectionError(RuntimeError):
    pass


class ClickUpActivityUnavailable(ClickUpCollectionError):
    """ClickUp did not expose usable task Activity History data."""


class ClickUpGateway:
    base = "https://api.clickup.com/api/v2"
    default_frontdoor_base = "https://frontdoor-prod-eu-west-1-2.clickup.com"

    def __init__(
        self,
        token: str,
        session=None,
        sleep=time.sleep,
        workspace_id: str = "",
        frontdoor_base_url: str | None = None,
        browser_profile_dir: str = "",
        storage_state_json: str = "",
        storage_state_path: str = "",
        browser_history_base_url: str = "",
        browser_headless: bool = True,
        browser_task_wait_ms: int = 4500,
    ):
        if not token or not token.startswith("pk_"):
            raise ClickUpCollectionError("ClickUp Personal API Token is missing or invalid.")
        self.token = token
        self.session = session or requests.Session()
        self.session.headers.update({"Authorization": token, "Accept": "application/json"})
        self.sleep = sleep
        self.workspace_id = str(workspace_id or "")
        self.frontdoor_base_url = (
            frontdoor_base_url
            or os.environ.get("CLICKUP_FRONTDOOR_BASE_URL")
            or self.default_frontdoor_base
        ).rstrip("/")
        self._web_history_disabled = False
        self._web_history_disable_reason = ""
        self._browser_history = ClickUpBrowserHistory(
            profile_dir=browser_profile_dir,
            storage_state_json=storage_state_json,
            storage_state_path=storage_state_path,
            history_base_url=browser_history_base_url,
            headless=browser_headless,
            task_wait_ms=browser_task_wait_ms,
        )

    def close(self):
        self.session.close()

    def request(self, path, *, params=None):
        for attempt in range(4):
            try:
                response = self.session.get(self.base + path, params=params, timeout=(10, 60))
            except requests.RequestException:
                if attempt < 3:
                    self.sleep(2 ** attempt)
                    continue
                raise ClickUpCollectionError("تعذر الاتصال بـ ClickUp. تحقق من الاتصال والتوكن.") from None
            if response.status_code == 429 or 500 <= response.status_code <= 599:
                if attempt < 3:
                    self.sleep(min(30, 2 ** attempt))
                    continue
                raise ClickUpCollectionError("ClickUp مشغول أو تم تجاوز حد الطلبات. حاول مرة أخرى لاحقاً.")
            if response.status_code in (401, 403):
                raise ClickUpCollectionError("توكن ClickUp غير صالح أو لا يملك الصلاحية المطلوبة.")
            if response.status_code == 404:
                raise ClickUpCollectionError("مورد ClickUp المطلوب غير متاح لهذا الحساب.")
            if response.status_code < 200 or response.status_code >= 300:
                raise ClickUpCollectionError(f"فشل طلب ClickUp (HTTP {response.status_code}).")
            try:
                return response.json()
            except ValueError:
                raise ClickUpCollectionError("ClickUp أعاد استجابة غير قابلة للقراءة.") from None
        raise ClickUpCollectionError("فشل الاتصال بـ ClickUp.")

    def _frontdoor_request(self, path, *, params=None, task_id: str = ""):
        """Call ClickUp's web Activity service without opening a browser."""
        if not self.workspace_id:
            raise ClickUpActivityUnavailable(
                "ClickUp web Activity History needs the ClickUp Workspace ID."
            )
        if self._web_history_disabled:
            raise ClickUpActivityUnavailable(
                self._web_history_disable_reason
                or "ClickUp web Activity History is unavailable for this account."
            )
        common_headers = {
            "Accept": "application/json",
            "Origin": "https://app.clickup.com",
            "Referer": "https://app.clickup.com/",
            "User-Agent": (
                "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
                "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
            ),
            "X-Requested-With": "XMLHttpRequest",
            "X-Workspace-ID": self.workspace_id,
        }
        if task_id:
            # These context headers are ignored by the public API, but are
            # accepted by different ClickUp web deployments and keep each
            # private-history request scoped to the selected task.
            common_headers.update({
                "X-Task-ID": str(task_id),
                "X-ClickUp-Task-ID": str(task_id),
                "X-Context-Task-ID": str(task_id),
            })

        auth_values = (self.token, f"Bearer {self.token}")
        for attempt in range(3):
            last_response = None
            for auth_index, authorization in enumerate(auth_values):
                headers = {**common_headers, "Authorization": authorization}
                try:
                    response = self.session.get(
                        self.frontdoor_base_url + path,
                        params=params,
                        headers=headers,
                        timeout=(10, 60),
                    )
                except requests.RequestException as exc:
                    if attempt < 2:
                        self.sleep(2 ** attempt)
                        break
                    raise ClickUpActivityUnavailable(
                        "تعذر الاتصال بمسار ClickUp الخاص بسجل النشاط."
                    ) from exc
                last_response = response
                if response.status_code in (401, 403) and auth_index == 0:
                    # Some web deployments accept Bearer while the public API
                    # accepts the raw Personal API Token (and vice versa).
                    continue
                break

            if last_response is None:
                continue
            response = last_response
            if response.status_code == 429 or 500 <= response.status_code <= 599:
                if attempt < 2:
                    self.sleep(min(30, 2 ** attempt))
                    continue
                raise ClickUpActivityUnavailable(
                    f"فشل مسار ClickUp الخاص بسجل النشاط (HTTP {response.status_code})."
                )
            if response.status_code in (401, 403):
                self._web_history_disabled = True
                self._web_history_disable_reason = (
                    "ClickUp web Activity History requires an authenticated web session for this account."
                )
                raise ClickUpActivityUnavailable(self._web_history_disable_reason)
            if response.status_code == 404:
                self._web_history_disabled = True
                self._web_history_disable_reason = (
                    "ClickUp web Activity History is not available for this account."
                )
                raise ClickUpActivityUnavailable(self._web_history_disable_reason)
            if response.status_code < 200 or response.status_code >= 300:
                raise ClickUpActivityUnavailable(
                    f"ClickUp web Activity History returned HTTP {response.status_code}."
                )
            try:
                payload = response.json()
            except ValueError as exc:
                raise ClickUpActivityUnavailable(
                    "ClickUp web Activity History returned unreadable data."
                ) from exc
            if not isinstance(payload, dict):
                raise ClickUpActivityUnavailable(
                    "ClickUp web Activity History returned an unexpected response."
                )
            return payload
        raise ClickUpActivityUnavailable("ClickUp web Activity History request failed.")

    @staticmethod
    def _normalise_history(history, task_id: str):
        result = []
        for event in history if isinstance(history, list) else []:
            if not isinstance(event, dict):
                continue
            item = dict(event)
            item.setdefault("task_id", str(task_id))
            if "timestamp" not in item and item.get("date") is not None:
                item["timestamp"] = item["date"]
            if "from" not in item and "before" in item:
                item["from"] = item.get("before")
            if "to" not in item and "after" in item:
                item["to"] = item.get("after")
            result.append(item)
        return result

    def _history_index(self, task_id: str):
        """Return the web history index, scoped to one task."""
        endpoint = f"/task-v3/experience/{self.workspace_id}/tasks/history"
        candidates = (
            [("fields[]", "users"), ("task_id", str(task_id))],
            [("fields[]", "users"), ("task_ids[]", str(task_id))],
        )
        last_error = None
        for params in candidates:
            try:
                data = self._frontdoor_request(endpoint, params=params, task_id=task_id)
            except ClickUpActivityUnavailable as exc:
                last_error = exc
                continue
            ids = data.get("ids")
            if isinstance(ids, list):
                return [str(value) for value in ids if value], data
        if last_error:
            raise last_error
        return [], {}

    def _history_details(self, task_id: str, history_ids):
        details = []
        ids = [str(value) for value in history_ids if value]
        for start in range(0, len(ids), 50):
            params = [("hist_ids[]", value) for value in ids[start:start + 50]]
            data = self._frontdoor_request(
                f"/tasks/v1/task/{task_id}/historyItems",
                params=params,
                task_id=task_id,
            )
            details.extend(data.get("history", []) if isinstance(data, dict) else [])
        return self._normalise_history(details, task_id)

    def _web_activity(self, task_id: str):
        task_id = str(task_id)
        if not self.workspace_id:
            raise ClickUpActivityUnavailable(
                "ClickUp web Activity History needs the ClickUp Workspace ID."
            )

        # A few ClickUp deployments return all items directly when hist_ids[]
        # is omitted. Try this first; otherwise use the scoped history index.
        direct_error = None
        try:
            direct = self._frontdoor_request(
                f"/tasks/v1/task/{task_id}/historyItems",
                task_id=task_id,
            )
            direct_history = self._normalise_history(direct.get("history", []), task_id)
            if direct_history:
                return direct_history
        except ClickUpActivityUnavailable as exc:
            direct_error = exc

        try:
            history_ids, index = self._history_index(task_id)
            if not history_ids:
                raise ClickUpActivityUnavailable(
                    "ClickUp returned no History IDs for the selected task."
                )
            events = self._history_details(task_id, history_ids)
            by_id = {
                str(item.get("id")): item
                for item in events
                if item.get("id") is not None
            }

            # Comment entries are part of the index but are not always returned
            # by the historyItems endpoint. Preserve them as Activity events so
            # the exported History ID column remains complete.
            comments = index.get("comments") if isinstance(index, dict) else {}
            users = index.get("userMap") if isinstance(index, dict) else {}
            for history_id, comment in (comments or {}).items():
                if not isinstance(comment, dict):
                    continue
                if str(comment.get("parent")) != task_id:
                    continue
                user_id = comment.get("userid")
                user = users.get(str(user_id), user_id) if isinstance(users, dict) else user_id
                by_id.setdefault(str(history_id), {
                    "id": str(history_id),
                    "task_id": task_id,
                    "date": comment.get("date"),
                    "timestamp": comment.get("date"),
                    "type": "comment",
                    "field": "comment",
                    "comment": comment.get("text_content", ""),
                    "user": user,
                    "source": comment.get("source"),
                    "from": None,
                    "to": None,
                })

            ordered = [by_id[value] for value in history_ids if value in by_id]
            # Keep any server-returned event whose ID was not present in the
            # index, but never leak an event belonging to another task.
            ordered.extend(
                item for item in events
                if str(item.get("id")) not in {str(value) for value in history_ids}
                and str(item.get("task_id", task_id)) == task_id
            )
            if not ordered:
                raise ClickUpActivityUnavailable(
                    "ClickUp returned History IDs but no task-scoped Activity events."
                )
            return ordered
        except ClickUpActivityUnavailable:
            raise
        except ClickUpCollectionError as exc:
            raise ClickUpActivityUnavailable(
                "ClickUp web Activity History could not be collected for this task."
            ) from exc
        except Exception as exc:
            raise ClickUpActivityUnavailable(
                "ClickUp web Activity History returned an unexpected payload."
            ) from exc

    @property
    def browser_history_configured(self):
        """Whether the optional ClickUp-only headless collector is configured."""
        return self._browser_history.configured

    def browser_activity(self, task_ids, progress=None):
        """Collect task Activity History through ClickUp's authenticated web UI."""
        if not self.browser_history_configured:
            raise ClickUpActivityUnavailable(
                "ClickUp web history needs a browser profile or storage state."
            )
        try:
            return self._browser_history.collect(task_ids, progress=progress)
        except ClickUpBrowserHistoryError as exc:
            raise ClickUpActivityUnavailable(str(exc)) from exc

    def workspaces(self):
        data = self.request("/team")
        teams = data.get("teams")
        if not isinstance(teams, list):
            raise ClickUpCollectionError("لم يتم العثور على Workspaces في ClickUp.")
        return teams

    def spaces(self, workspace_id: str, archived: bool = False):
        data = self.request(f"/team/{workspace_id}/space", params={"archived": str(archived).lower()})
        spaces = data.get("spaces")
        if not isinstance(spaces, list):
            raise ClickUpCollectionError("لم يتم العثور على Spaces في ClickUp.")
        return spaces

    def folders(self, space_id: str, archived: bool = False):
        data = self.request(f"/space/{space_id}/folder", params={"archived": str(archived).lower()})
        return data.get("folders", [])

    def folderless_lists(self, space_id: str, archived: bool = False):
        data = self.request(f"/space/{space_id}/list", params={"archived": str(archived).lower()})
        return data.get("lists", [])

    def lists(self, folder_id: str, archived: bool = False):
        data = self.request(f"/folder/{folder_id}/list", params={"archived": str(archived).lower()})
        return data.get("lists", [])

    def tasks(self, list_id: str, page: int = 0):
        return self.request(f"/list/{list_id}/task", params={
            "page": page, "include_closed": "true", "subtasks": "true", "include_timl": "true",
        })

    def all_tasks_for_space(self, space_id: str, progress=None):
        lists = list(self.folderless_lists(space_id))
        for folder in self.folders(space_id):
            lists.extend(self.lists(str(folder["id"])))
        result, seen = [], set()
        for current_list in lists:
            page = 0
            while True:
                data = self.tasks(str(current_list["id"]), page)
                tasks = data.get("tasks", [])
                for task in tasks:
                    task_id = str(task.get("id", ""))
                    if task_id and task_id not in seen:
                        seen.add(task_id)
                        result.append(task)
                if progress:
                    progress(f"ClickUp: تم جمع {len(result)} مهمة...")
                if len(tasks) < 100:
                    break
                page += 1
        return result

    def activity(self, task_id: str):
        try:
            data = self.request(f"/task/{task_id}/activity")
        except ClickUpCollectionError as public_error:
            if str(public_error) != "مورد ClickUp المطلوب غير متاح لهذا الحساب.":
                raise
            try:
                return self._web_activity(task_id)
            except ClickUpActivityUnavailable as web_error:
                raise ClickUpActivityUnavailable(
                    f"ClickUp public Activity History is unavailable; web Activity History also failed: {web_error}"
                ) from public_error
        return data.get("activity", data.get("events", [])) if isinstance(data, dict) else []

