from __future__ import annotations

import openpyxl
from openpyxl import Workbook
from openpyxl.utils import column_index_from_string
from openpyxl.worksheet.worksheet import Worksheet

from .config import URL_COL, DEST_COL, GRANTED_COL, PUB_COL, JUSTIA_BASE, MAX_CHARS, ROW_HEIGHT


def load_workbook(filepath: str) -> tuple[Workbook, Worksheet]:
    wb = openpyxl.load_workbook(filepath)
    return wb, wb.active


def _needs_fetch(dest_val: str) -> bool:
    """Return True if the destination cell is empty or contains a failed extraction marker."""
    v = dest_val.strip()
    return not v or v.startswith("[擷取失敗]")


def load_urls_from_file(
    filepath: str,
    start_row: int = 2,
    url_col: str = "AV",
    single_row: bool = False,
    skip_done: bool = True,
) -> list[tuple[int, str]]:
    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb.active
    col_idx = column_index_from_string(url_col.upper().strip())
    dest_col_idx = col_idx + 1

    if single_row:
        scan_range = range(max(start_row, 2), start_row + 1)
    else:
        scan_range = range(2, ws.max_row + 1)

    # Try reading URLs directly from the specified column
    url_col_has_data = False
    rows = []
    for i in scan_range:
        val = ws.cell(row=i, column=col_idx).value
        if not val or not str(val).strip() or str(val).startswith("="):
            continue
        url_col_has_data = True
        if skip_done:
            dest_val = str(ws.cell(row=i, column=dest_col_idx).value or "")
            if not _needs_fetch(dest_val):
                continue
        rows.append((i, str(val).strip()))

    if url_col_has_data:
        return rows

    # Fallback: build Justia URLs from K (Granted) / M (Pub) columns
    rows = []
    for i in scan_range:
        granted = str(ws.cell(row=i, column=GRANTED_COL).value or "").strip()
        pub     = str(ws.cell(row=i, column=PUB_COL).value or "").strip()
        patent_num = granted if granted else pub
        if not patent_num:
            continue
        if skip_done:
            dest_val = str(ws.cell(row=i, column=dest_col_idx).value or "")
            if not _needs_fetch(dest_val):
                continue
        rows.append((i, JUSTIA_BASE + patent_num))
    return rows


def write_result(ws: Worksheet, row_num: int, text: str, dest_col: int = DEST_COL) -> None:
    ws.cell(row=row_num, column=dest_col).value = text[:MAX_CHARS]
    ws.row_dimensions[row_num].height = ROW_HEIGHT


def save(wb: Workbook, filepath: str) -> None:
    wb.save(filepath)
