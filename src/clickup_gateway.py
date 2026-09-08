"""Read-only ClickUp API v2 access for the collection workflow."""
from __future__ import annotations

import time
import requests


class ClickUpCollectionError(RuntimeError):
    pass


class ClickUpGateway:
    base = "https://api.clickup.com/api/v2"

    def __init__(self, token: str, session=None, sleep=time.sleep):
        if not token or not token.startswith("pk_"):
            raise ClickUpCollectionError("ClickUp Personal API Token is missing or invalid.")
        self.session = session or requests.Session()
        self.session.headers.update({"Authorization": token, "Accept": "application/json"})
        self.sleep = sleep

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
        data = self.request(f"/task/{task_id}/activity")
        return data.get("activity", data.get("events", [])) if isinstance(data, dict) else []

