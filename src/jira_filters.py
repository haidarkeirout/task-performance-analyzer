"""Build project-scoped JQL from Jira's own searchable-field metadata."""
from __future__ import annotations

import hashlib
import html
import json
import re
from dataclasses import dataclass
from datetime import date, timedelta


class FilterError(ValueError):
    pass


@dataclass(frozen=True)
class FilterField:
    id: str
    label: str
    jql: str
    kind: str
    operators: tuple[str, ...]
    auto: bool = False


ALIASES = {"type": "issuetype", "issue type": "issuetype", "worktype": "issuetype",
           "space": "project", "work item type": "issuetype", "createddate": "created",
           "duedate": "due", "updateddate": "updated", "resolved": "resolutiondate"}
BASIC_IDS = ("assignee", "issuetype", "status", "created")
DISPLAY_NAMES = {"project": "Space", "issuetype": "Type", "assignee": "Assignee",
                 "status": "Status", "created": "Created", "due": "Due date", "updated": "Updated",
                 "resolutiondate": "Resolved", "fixversion": "Fix version", "affectedversion": "Affects version",
                 "reporter": "Reporter", "priority": "Priority", "resolution": "Resolution",
                 "labels": "Labels", "component": "Component", "text": "Text"}


def plain_label(text):
    return html.unescape(re.sub(r"<[^>]+>", "", str(text)))


def unquote_value(text):
    value = str(text)
    if value.startswith('"') and value.endswith('"'):
        try:
            decoded = json.loads(value)
            if isinstance(decoded, str):
                return decoded
        except ValueError:
            pass
    return value


def quoted(value):
    # Always quote literals, including IDs, to prevent filter values changing query structure.
    return json.dumps(str(value), ensure_ascii=False)


def field_catalog(reference, definitions):
    by_id = {f["id"]: f for f in definitions if f.get("id")}
    catalog = {}
    for entry in reference.get("visibleFieldNames", []):
        if str(entry.get("searchable", "true")).lower() == "false" or not entry.get("operators"):
            continue
        raw = unquote_value(entry.get("cfid") or entry.get("value", ""))
        if not raw:
            continue
        fid = ALIASES.get(raw.casefold(), raw.casefold())
        custom = re.fullmatch(r"cf\[(\d+)\]", fid)
        definition = by_id.get("customfield_" + custom[1] if custom else fid, {})
        schema = definition.get("schema") or {}
        types = " ".join(entry.get("types", [])) + " " + str(schema.get("type", ""))
        if re.search(r"date|timestamp", types, re.I) or fid in {"created", "updated", "due", "resolutiondate", "lastviewed"}:
            kind = "date"
        elif re.search(r"long|double|integer|number|float", types, re.I):
            kind = "number"
        elif re.search(r"user", types, re.I):
            kind = "user"
        else:
            kind = "value"
        label = definition.get("name") or plain_label(entry.get("displayName") or raw)
        if custom:
            label = re.sub(r"\s*-\s*cf\[\d+\]$", "", label) + f" ({fid})"
        else:
            label = DISPLAY_NAMES.get(fid, label[:1].upper() + label[1:])
        jql_name = fid if re.fullmatch(r"[a-z][a-z0-9_.]*|cf\[\d+\]", fid) else quoted(raw)
        catalog[fid] = FilterField(fid, label, jql_name, kind,
                                   tuple(op.lower() for op in entry["operators"]),
                                   str(entry.get("auto", "false")).lower() == "true")
    # Standard controls have stable Jira semantics if an account's metadata omits a basic alias.
    for fid in BASIC_IDS:
        if fid not in catalog:
            ops = (">=", "<=", ">", "<", "=", "!=", "is", "is not") if fid == "created" else ("in", "not in", "=", "!=", "is", "is not")
            catalog[fid] = FilterField(fid, DISPLAY_NAMES[fid], fid, "date" if fid == "created" else "value", ops, fid != "created")
    return catalog


def field_clause(field, operator, values=(), value_mode="Values"):
    operator = operator.lower()
    if operator not in field.operators:
        raise FilterError(f"{field.label}: select a supported comparison.")
    if operator in {"is", "is not"}:
        return f"{field.jql} {operator.upper()} EMPTY"
    if operator == "changed":
        return f"{field.jql} CHANGED"
    values = [str(v) for v in values if str(v).strip()]
    if not values:
        return ""
    rendered = []
    for value in values:
        if value_mode == "JQL expression":
            validate_condition(value)
            rendered.append(value)
        elif field.kind == "number":
            if not re.fullmatch(r"-?\d+(?:\.\d+)?", value.strip()):
                raise FilterError(f"{field.label}: enter a numeric value.")
            rendered.append(value.strip())
        else:
            rendered.append(quoted(value))
    if operator in {"in", "not in", "was in", "was not in"}:
        right = "(" + ", ".join(rendered) + ")"
    elif len(rendered) != 1:
        raise FilterError(f"{field.label}: select one value for this comparison.")
    else:
        right = rendered[0]
    return f"{field.jql} {operator.upper()} {right}"


def date_clause(field, mode, start=None, end=None, days=7):
    name = field.jql
    if mode == "Any time":
        return ""
    if mode == "Is empty":
        return field_clause(field, "is")
    if mode == "Is not empty":
        return field_clause(field, "is not")
    if mode == "Within the last":
        return f'{name} >= "-{int(days)}d" AND {name} <= now()'
    functions = {"This week": "startOfWeek()", "This month": "startOfMonth()", "This year": "startOfYear()"}
    if mode in functions:
        return f"{name} >= {functions[mode]} AND {name} <= now()"
    if not isinstance(start, date):
        raise FilterError(f"{field.label}: select a date.")
    if mode == "On":
        return f"{name} >= {quoted(start.isoformat())} AND {name} < {quoted((start + timedelta(days=1)).isoformat())}"
    if mode == "Before":
        return f"{name} < {quoted(start.isoformat())}"
    if mode == "After":
        return f"{name} >= {quoted((start + timedelta(days=1)).isoformat())}"
    if mode == "Between":
        if not isinstance(end, date) or start > end:
            raise FilterError(f"{field.label}: the end date must be on or after the start date.")
        return f"{name} >= {quoted(start.isoformat())} AND {name} < {quoted((end + timedelta(days=1)).isoformat())}"
    raise FilterError("Select a valid date comparison.")


def validate_condition(value):
    """Reject an unbalanced clause before placing it inside the fixed project scope."""
    depth, quote_char, escaped, outside = 0, None, False, []
    for char in value:
        if quote_char:
            if escaped:
                escaped = False
            elif char == "\\":
                escaped = True
            elif char == quote_char:
                quote_char = None
            outside.append(" ")
        elif char in "\"'":
            quote_char = char
            outside.append(" ")
        else:
            outside.append(char)
            if char == "(": depth += 1
            elif char == ")": depth -= 1
            if depth < 0:
                raise FilterError("The JQL condition contains an unmatched parenthesis.")
    if quote_char or depth:
        raise FilterError("The JQL condition contains an unfinished quote or parenthesis.")
    if re.search(r"\border\s+by\b", "".join(outside), re.I):
        raise FilterError("Enter the JQL filter condition without ORDER BY. Work items are sorted automatically.")


def build_query(project_key, clauses=(), text="", advanced=""):
    if not re.fullmatch(r"[A-Za-z][A-Za-z0-9_]*", project_key):
        raise FilterError("Select a valid Jira space.")
    conditions = [f"project = {quoted(project_key)}"]
    for clause in clauses:
        if clause:
            validate_condition(clause)
            conditions.append(f"({clause})")
    if text.strip():
        conditions.append(f"text ~ {quoted(text.strip())}")
    if advanced.strip():
        validate_condition(advanced)
        conditions.append(f"({advanced.strip()})")
    return " AND ".join(conditions) + " ORDER BY key ASC"


def query_fingerprint(project_id, query, revision=""):
    return hashlib.sha256(json.dumps([str(project_id), query, revision]).encode()).hexdigest()
