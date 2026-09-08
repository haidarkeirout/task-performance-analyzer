"""Single account authentication. Configuration stays on the server, outside Git.

This module has no default account, registration, or administrative roles.
"""
from __future__ import annotations

import hashlib
import hmac
import os
import secrets
from dataclasses import dataclass, field
from typing import Mapping
from urllib.parse import urlsplit


class SetupError(ValueError):
    pass


@dataclass(frozen=True)
class Settings:
    username: str
    password: str = field(repr=False)
    password_hash: str = field(repr=False)
    jira_url: str
    jira_email: str = field(repr=False)
    jira_token: str = field(repr=False)
    jira_cloud_id: str = ""
    source_timezone: str = "Asia/Damascus"
    start_date_field: str = ""

    @property
    def revision(self) -> str:
        # Invalidates an existing session if its server-side credentials change.
        value = "\0".join(str(getattr(self, name)) for name in self.__dataclass_fields__)
        return hashlib.sha256(value.encode()).hexdigest()


def read_settings(values: Mapping, environment: Mapping | None = None) -> Settings:
    environment = os.environ if environment is None else environment

    def get(name, default=""):
        return str(values.get(name, environment.get(name, default)))

    username = get("APP_USERNAME").strip()
    password, password_hash = get("APP_PASSWORD"), get("APP_PASSWORD_HASH")
    url = get("JIRA_BASE_URL").strip().rstrip("/")
    email, token = get("JIRA_EMAIL").strip(), get("JIRA_API_TOKEN").strip()
    if not username or not (password or password_hash) or not all([url, email, token]):
        raise SetupError("This application is not configured yet. Please contact the administrator.")
    parsed = urlsplit(url)
    if (parsed.scheme != "https" or not parsed.hostname or parsed.username or parsed.password
            or parsed.query or parsed.fragment or parsed.path or parsed.port not in (None, 443)):
        raise SetupError("The Jira connection configuration is invalid. Please contact the administrator.")
    cloud_id = get("JIRA_CLOUD_ID").strip()
    if cloud_id and any(c not in "0123456789abcdefABCDEF-" for c in cloud_id):
        raise SetupError("The Jira connection configuration is invalid. Please contact the administrator.")
    return Settings(username, password, password_hash, url, email, token, cloud_id,
                    get("SOURCE_TIMEZONE", "Asia/Damascus"), get("JIRA_START_DATE_FIELD_ID").strip())


def hash_password(password: str, iterations: int = 600_000) -> str:
    salt = secrets.token_hex(16)
    digest = hashlib.pbkdf2_hmac("sha256", password.encode(), salt.encode(), iterations).hex()
    return f"pbkdf2_sha256${iterations}${salt}${digest}"


def credentials_match(username: str, password: str, settings: Settings) -> bool:
    user_ok = hmac.compare_digest(username.strip().encode(), settings.username.encode())
    if settings.password_hash:
        try:
            algorithm, iterations, salt, expected = settings.password_hash.split("$")
            rounds = int(iterations)
            if algorithm != "pbkdf2_sha256" or not 100_000 <= rounds <= 2_000_000:
                return False
            actual = hashlib.pbkdf2_hmac("sha256", password.encode(), salt.encode(), rounds).hex()
            password_ok = hmac.compare_digest(actual, expected)
        except (ValueError, TypeError):
            return False
    else:
        # The simple setup stores the password in Streamlit Secrets, never in code.
        password_ok = hmac.compare_digest(password.encode(), settings.password.encode())
    return user_ok and password_ok
