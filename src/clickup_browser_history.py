"""Headless ClickUp web Activity History collection.

The public ClickUp API does not expose the complete task history used by the
ClickUp web application.  This module reproduces the proven, read-only browser
flow from ``clickup_backend_history_test.py`` without showing a browser window:
open each task, open Activity, capture the history index, then fetch the
task-scoped history details with the authenticated browser context.

The dependency is optional at import time.  Jira never imports or executes this
module; the ClickUp gateway opts in only when a ClickUp browser profile or
storage-state secret is configured.
"""
from __future__ import annotations

import asyncio
import json
import os
from pathlib import Path
from urllib.parse import urlencode, urlsplit


class ClickUpBrowserHistoryError(RuntimeError):
    """The authenticated headless ClickUp history flow could not complete."""


def _load_storage_state(raw: str = "", path: str = ""):
    value = raw.strip()
    if not value and path:
        try:
            value = Path(path).expanduser().read_text(encoding="utf-8")
        except OSError as exc:
            raise ClickUpBrowserHistoryError(
                "ClickUp browser storage state could not be read."
            ) from exc
    if not value:
        return None
    try:
        state = json.loads(value)
    except json.JSONDecodeError as exc:
        raise ClickUpBrowserHistoryError(
            "CLICKUP_STORAGE_STATE_JSON is not valid JSON."
        ) from exc
    if not isinstance(state, dict):
        raise ClickUpBrowserHistoryError(
            "CLICKUP_STORAGE_STATE_JSON must be a Playwright storage-state object."
        )
    return state


def _history_url_origin(url: str) -> str:
    parsed = urlsplit(url)
    if not parsed.scheme or not parsed.netloc:
        raise ClickUpBrowserHistoryError("ClickUp returned an invalid history URL.")
    return f"{parsed.scheme}://{parsed.netloc}"


class ClickUpBrowserHistory:
    """Collect task Activity History using an authenticated headless context."""

    def __init__(
        self,
        *,
        profile_dir: str = "",
        storage_state_json: str = "",
        storage_state_path: str = "",
        app_base_url: str = "https://app.clickup.com",
        history_base_url: str = "",
        headless: bool = True,
        task_wait_ms: int = 4500,
        show_more_limit: int = 50,
        executable_path: str = "",
    ):
        self.profile_dir = str(profile_dir or os.environ.get("CLICKUP_BROWSER_PROFILE_DIR", "")).strip()
        self.storage_state_json = storage_state_json or os.environ.get("CLICKUP_STORAGE_STATE_JSON", "")
        self.storage_state_path = storage_state_path or os.environ.get("CLICKUP_STORAGE_STATE_PATH", "")
        self.app_base_url = (app_base_url or "https://app.clickup.com").rstrip("/")
        self.history_base_url = (history_base_url or "").rstrip("/")
        self.headless = bool(headless)
        self.task_wait_ms = max(500, int(task_wait_ms))
        self.show_more_limit = max(1, int(show_more_limit))
        self.executable_path = str(executable_path or os.environ.get("CLICKUP_BROWSER_EXECUTABLE", "")).strip()
        if not self.executable_path:
            for candidate in ("/usr/bin/chromium", "/usr/bin/chromium-browser", "/usr/bin/google-chrome"):
                if Path(candidate).exists():
                    self.executable_path = candidate
                    break

    @property
    def configured(self) -> bool:
        return bool(self.profile_dir or self.storage_state_json or self.storage_state_path)

    def collect(self, task_ids, progress=None):
        """Return ``{task_id: [history events]}`` for the supplied tasks."""
        if not self.configured:
            raise ClickUpBrowserHistoryError(
                "ClickUp web history needs CLICKUP_BROWSER_PROFILE_DIR or CLICKUP_STORAGE_STATE_JSON."
            )
        unique = []
        seen = set()
        for value in task_ids or []:
            task_id = str(value or "").strip()
            if task_id and task_id not in seen:
                seen.add(task_id)
                unique.append(task_id)
        if not unique:
            return {}
        try:
            asyncio.get_running_loop()
        except RuntimeError:
            return asyncio.run(self._collect(unique, progress))
        raise ClickUpBrowserHistoryError(
            "ClickUp web history cannot start inside an already-running event loop."
        )

    async def _open_context(self, playwright):
        try:
            chromium = playwright.chromium
            launch_options = {
                "headless": self.headless,
                "viewport": {"width": 1440, "height": 900},
            }
            if self.executable_path:
                launch_options["executable_path"] = self.executable_path
            if self.profile_dir:
                context = await chromium.launch_persistent_context(
                    str(Path(self.profile_dir).expanduser()),
                    **launch_options,
                )
                return context, None
            browser = await chromium.launch(**launch_options)
            context = await browser.new_context(
                storage_state=_load_storage_state(self.storage_state_json, self.storage_state_path),
            )
            return context, browser
        except ClickUpBrowserHistoryError:
            raise
        except Exception as exc:
            raise ClickUpBrowserHistoryError(
                "ClickUp headless browser could not start. Install Playwright Chromium and verify the browser configuration."
            ) from exc

    async def _collect(self, task_ids, progress):
        try:
            from playwright.async_api import async_playwright
        except ImportError as exc:
            raise ClickUpBrowserHistoryError(
                "ClickUp web history requires the optional Playwright dependency."
            ) from exc

        results = {}
        errors = {}
        async with async_playwright() as playwright:
            context, browser = await self._open_context(playwright)
            try:
                page = context.pages[0] if context.pages else await context.new_page()
                for number, task_id in enumerate(task_ids, 1):
                    if progress:
                        progress(f"Collecting ClickUp Activity: {number} of {len(task_ids)} tasks...")
                    try:
                        results[task_id] = await self._collect_task(context, page, task_id)
                    except ClickUpBrowserHistoryError as exc:
                        errors[task_id] = str(exc)
                        results[task_id] = []
            finally:
                await context.close()
                if browser is not None:
                    await browser.close()

        if not results:
            raise ClickUpBrowserHistoryError(
                "ClickUp web history returned no task results."
            )
        # Keep per-task failures available to the exporter without discarding
        # successful tasks.  The exporter treats these as ClickUp-only notes.
        results["__errors__"] = errors
        return results

    async def _collect_task(self, context, page, task_id: str):
        captured = []

        async def capture_history(response):
            url = response.url.lower()
            is_index = "/task-v3/experience/" in url and "/tasks/history" in url
            is_detail = "/tasks/v1/task/" in url and "/historyitems" in url
            if not (is_index or is_detail):
                return
            try:
                body = await response.json()
                if isinstance(body, dict):
                    captured.append({"url": response.url, "status": response.status, "body": body})
            except Exception:
                pass

        page.on("response", capture_history)
        try:
            await page.goto(
                f"{self.app_base_url}/t/{task_id}",
                wait_until="domcontentloaded",
                timeout=30_000,
            )
            await page.wait_for_timeout(self.task_wait_ms)
            if any(marker in page.url.lower() for marker in ("/login", "/signin", "/sign-in")):
                raise ClickUpBrowserHistoryError(
                    "ClickUp browser session is not authenticated."
                )

            activity = page.get_by_text("Activity", exact=True)
            if await activity.count():
                try:
                    await activity.last.click(timeout=1500)
                except Exception:
                    pass
                await page.wait_for_timeout(1500)

            for _ in range(self.show_more_limit):
                more = page.get_by_text("Show more", exact=True)
                if not await more.count():
                    break
                try:
                    await more.first.click(timeout=1200)
                    await page.wait_for_timeout(600)
                except Exception:
                    break
            await page.wait_for_timeout(1000)
        except ClickUpBrowserHistoryError:
            raise
        except Exception as exc:
            raise ClickUpBrowserHistoryError(
                f"ClickUp Activity could not be opened for task {task_id}."
            ) from exc
        finally:
            page.remove_listener("response", capture_history)

        index = None
        for request in captured:
            body = request.get("body") or {}
            if "/task-v3/experience/" in request.get("url", "").lower() and isinstance(body.get("ids"), list):
                index = request
        if index is None:
            # A task can legitimately have no history; return an empty list.
            return []

        index_body = index["body"]
        history_ids = []
        for value in index_body.get("ids", []):
            value = str(value or "")
            if value and value not in history_ids:
                history_ids.append(value)
        if not history_ids:
            return []

        origin = self.history_base_url or _history_url_origin(index["url"])
        events = []
        request_context = context.request
        for start in range(0, len(history_ids), 6):
            batch = history_ids[start:start + 6]
            query = urlencode([( "hist_ids[]", value) for value in batch])
            url = f"{origin}/tasks/v1/task/{task_id}/historyItems?{query}"
            try:
                response = await request_context.get(url, timeout=60_000)
                if not response.ok:
                    raise ClickUpBrowserHistoryError(
                        f"ClickUp Activity details returned HTTP {response.status}."
                    )
                payload = await response.json()
                if isinstance(payload, dict):
                    events.extend(payload.get("history", []) or [])
            except Exception as exc:
                raise ClickUpBrowserHistoryError(
                    f"ClickUp Activity details could not be read for task {task_id}."
                ) from exc

        by_id = {}
        for event in events:
            if not isinstance(event, dict):
                continue
            event_id = str(event.get("id") or "")
            event_task = str(event.get("task_id") or task_id)
            if event_id and event_task == task_id:
                item = dict(event)
                item.setdefault("task_id", task_id)
                item.setdefault("timestamp", item.get("date"))
                item.setdefault("from", item.get("before"))
                item.setdefault("to", item.get("after"))
                by_id[event_id] = item

        comments = index_body.get("comments") or {}
        users = index_body.get("userMap") or {}
        if isinstance(comments, dict):
            for history_id, comment in comments.items():
                if not isinstance(comment, dict):
                    continue
                parent = comment.get("parent") or comment.get("parent_id")
                if str(parent) != task_id:
                    continue
                history_id = str(history_id)
                user_id = comment.get("userid") or comment.get("user_id")
                user = users.get(str(user_id), user_id) if isinstance(users, dict) else user_id
                by_id.setdefault(history_id, {
                    "id": history_id,
                    "task_id": task_id,
                    "date": comment.get("date"),
                    "timestamp": comment.get("date"),
                    "type": "comment",
                    "field": "comment",
                    "comment": comment.get("text_content", comment.get("text", "")),
                    "user": user,
                    "source": comment.get("source"),
                    "from": None,
                    "to": None,
                })

        # Preserve every ID observed in the web history index.  An index-only
        # placeholder is explicit and task-scoped; it never invents a field or
        # status value, but keeps the exported History IDs complete.
        for history_id in history_ids:
            by_id.setdefault(history_id, {
                "id": history_id,
                "task_id": task_id,
                "type": "history_id",
                "field": None,
                "source": "ClickUp history index",
                "from": None,
                "to": None,
            })
        return [by_id[history_id] for history_id in history_ids]


__all__ = ["ClickUpBrowserHistory", "ClickUpBrowserHistoryError"]
