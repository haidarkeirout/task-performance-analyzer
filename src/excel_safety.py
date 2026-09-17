"""Shared safeguards for every generated Excel workbook."""

from __future__ import annotations

from typing import Any

from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE


EXCEL_CELL_LIMIT = 32767


def safe_excel_value(value: Any) -> Any:
    """Return an Excel-compatible value without executable string formulas."""
    if not isinstance(value, str):
        return value
    cleaned = ILLEGAL_CHARACTERS_RE.sub(
        lambda match: f"\\u{ord(match.group(0)):04x}", value
    )
    return cleaned[:EXCEL_CELL_LIMIT]


def write_excel_cell(sheet: Any, row: int, column: int, value: Any):
    """Write one typed cell and force all source strings to remain plain text."""
    cleaned = safe_excel_value(value)
    cell = sheet.cell(row=row, column=column, value=cleaned)
    if isinstance(cleaned, str):
        cell.data_type = "s"
    return cell
