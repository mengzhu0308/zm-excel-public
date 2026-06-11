"""Shared helpers for extract_template.py and write_back.py.

Centralizes:
  - select_worksheet(): pick a worksheet by name or fall back to active sheet
  - detect_header_row(): pick the row among the first 3 with the most non-empty cells
  - find_sample_row(): locate the last data row aligned to the header
  - copy_styles(): copy visual styles from one cell to another
  - load_workbook(): explicit data_only policy wrapper
  - normalize_header(): collapse internal whitespace, replace newlines, strip

Keeping these in one place avoids drift between extract_template.py and
write_back.py.
"""

from __future__ import annotations

import re
import sys
from copy import copy
from pathlib import Path
from typing import Optional

import openpyxl

_WHITESPACE_RE = re.compile(r"\s+")


HEADER_SCAN_ROWS = 3

# File extensions this skill actually supports. openpyxl can only read/write
# .xlsx and .xlsm, so anything else should be rejected early with a clear
# message instead of letting openpyxl raise a confusing zipfile error.
SUPPORTED_EXCEL_EXTENSIONS = (".xlsx", ".xlsm")


def validate_excel_extension(path) -> None:
    """Raise ValueError if path does not have a supported Excel extension.

    Keeps the user-facing extension check in one place so extract_template.py
    and write_back.py stay aligned. The check is purely about the extension
    string — actual openpyxl errors are still surfaced as-is for malformed
    .xlsx files.
    """
    suffix = Path(path).suffix.lower()
    if suffix not in SUPPORTED_EXCEL_EXTENSIONS:
        raise ValueError(
            f"Unsupported file extension: '{suffix or '(none)'}'. "
            f"This skill only supports {', '.join(SUPPORTED_EXCEL_EXTENSIONS)}. "
            f"If you have a legacy .xls file, please convert it to .xlsx first."
        )


def select_worksheet(wb, sheet_name: Optional[str]):
    """Pick a worksheet by name, or fall back to the active sheet.

    When sheet_name is None and the workbook has multiple sheets, prints a
    stderr note listing the available sheets so the user can confirm the
    active sheet is the intended target. This keeps extract_template.py
    and write_back.py aligned on multi-sheet handling.

    Args:
        wb: An openpyxl Workbook.
        sheet_name: Optional worksheet name. If provided and exists, it is used.

    Returns:
        The selected openpyxl Worksheet.

    Raises:
        ValueError: If sheet_name is provided but not found, or if the workbook
                    has no active sheet.
    """
    if sheet_name:
        if sheet_name in wb.sheetnames:
            return wb[sheet_name]
        available = ", ".join(f"'{s}'" for s in wb.sheetnames)
        raise ValueError(
            f"Worksheet '{sheet_name}' not found. "
            f"Available sheets: {available}"
        )
    ws = wb.active
    if ws is None:
        raise ValueError("Excel file has no active worksheet")
    if len(wb.sheetnames) > 1:
        sheets_str = ", ".join(f"'{s}'" for s in wb.sheetnames)
        print(
            f"Note: File has multiple sheets ({sheets_str}). "
            f"Using active sheet '{ws.title}'. Use --sheet to specify another.",
            file=sys.stderr,
        )
    return ws


def detect_header_row(ws) -> int:
    """Auto-detect header row: pick the row among the first HEADER_SCAN_ROWS
    with the most non-empty cells.

    Raises:
        ValueError: If the worksheet has no readable rows.
    """
    def _non_empty_count(row_cells):
        return sum(1 for c in row_cells if c.value is not None)

    candidate_rows = []
    upper = min(HEADER_SCAN_ROWS, ws.max_row if ws.max_row else 0) + 1
    for r in range(1, max(2, upper)):
        candidate_rows.append((_non_empty_count(ws[r]), r))

    if not candidate_rows:
        raise ValueError("Excel file has no readable rows")

    _, header_row_idx = max(candidate_rows, key=lambda x: x[0])
    return header_row_idx


def find_sample_row(ws, header_row_idx: int) -> Optional[int]:
    """Find the last row that looks like a data row (aligned to headers).

    A data row must have values in at least half of the *named* header
    columns. The "named" set is derived from read_headers() so that
    whitespace-only header cells are treated the same way as empty header
    cells, matching what write_back considers a real column.
    """
    headers = read_headers(ws, header_row_idx)
    header_cols = [
        cell.column
        for cell, name in zip(ws[header_row_idx], headers)
        if name != ""
    ]
    if not header_cols:
        return None

    threshold = max(1, len(header_cols) // 2)
    for row_idx in range(ws.max_row, header_row_idx, -1):
        vals = sum(
            1
            for c in header_cols
            if ws.cell(row=row_idx, column=c).value is not None
        )
        if vals >= threshold:
            return row_idx
    return None


def copy_styles(src_cell, dst_cell) -> None:
    """Copy visual styles from src_cell to dst_cell."""
    if src_cell.has_style:
        dst_cell.font = copy(src_cell.font)
        dst_cell.border = copy(src_cell.border)
        dst_cell.fill = copy(src_cell.fill)
        dst_cell.number_format = copy(src_cell.number_format)
        dst_cell.protection = copy(src_cell.protection)
        dst_cell.alignment = copy(src_cell.alignment)


def normalize_header(value) -> str:
    """Normalize a header cell value into a string for matching.

    - None -> ""
    - str -> collapse all runs of whitespace (including \\n) into a
      single space, then strip
    - other -> str(...)

    Collapsing internal whitespace ensures that headers like
    "姓 名" and "姓  名" are treated as the same field name when
    matching between Excel and Markdown.
    """
    if value is None:
        return ""
    return _WHITESPACE_RE.sub(" ", str(value)).strip()


def read_headers(ws, header_row_idx: int) -> list[str]:
    """Read the header row into a normalized list of strings.

    Empty strings are preserved in their original column position so that the
    column index in the headers list matches the column index in the sheet.
    """
    headers = []
    for cell in ws[header_row_idx]:
        headers.append(normalize_header(cell.value))
    return headers


def load_workbook(path: str, *, data_only: bool):
    """Wrapper around openpyxl.load_workbook with explicit data_only policy.

    Use data_only=True when you need computed values (template extraction).
    Use data_only=False when you need formulas preserved (write-back).
    """
    return openpyxl.load_workbook(path, data_only=data_only)


def assert_headers_present(headers: list[str]) -> None:
    """Raise ValueError when no named headers are present.

    A single helper to keep extract_template.py and write_back.py aligned:
    both scripts must refuse to continue when the detected header row is
    entirely empty or whitespace-only (which normalize_header collapses
    to ""). Without this guard, write_back would silently produce an
    all-None new row.
    """
    if not headers or all(h == "" for h in headers):
        raise ValueError("No headers found")


__all__ = [
    "HEADER_SCAN_ROWS",
    "SUPPORTED_EXCEL_EXTENSIONS",
    "select_worksheet",
    "detect_header_row",
    "find_sample_row",
    "copy_styles",
    "normalize_header",
    "read_headers",
    "load_workbook",
    "validate_excel_extension",
    "assert_headers_present",
]
