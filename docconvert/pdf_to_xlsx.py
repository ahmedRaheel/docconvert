"""PDF -> Excel.

Every detected table becomes real cells (numbers are converted to numeric cells,
header row styled). Pages without ruled tables fall back to splitting text lines
on wide whitespace gaps so column-aligned data still lands in columns.
"""
from __future__ import annotations

import re

from .utils import coerce_cell, parse_pages

_TEXT_STRATEGY = {"vertical_strategy": "text", "horizontal_strategy": "text"}


def _clean(cell):
    if cell is None:
        return None
    return re.sub(r"[ \t]+", " ", str(cell)).strip() or None


def _text_rows(page):
    txt = page.extract_text(layout=True) or ""
    rows = []
    for line in txt.splitlines():
        if not line.strip():
            continue
        rows.append([c for c in re.split(r"\s{2,}", line.strip())])
    return rows


def pdf_to_xlsx(src, dst, *, pages=None, sheet_mode: str = "page", table_strategy: str = "auto",
                text_fallback: bool = True, **_):
    """
    sheet_mode:      ``"page"`` (one sheet per page, tables stacked) or ``"table"`` (one sheet per table)
    table_strategy:  ``"auto"`` (ruled tables), ``"lines"`` or ``"text"`` (infer from alignment)
    text_fallback:   when a page has no table, put its text lines into rows/columns
    """
    import pdfplumber
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    if sheet_mode not in ("page", "table"):
        raise ValueError("sheet_mode must be 'page' or 'table'")
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    head_fill = PatternFill("solid", fgColor="D9E1F2")

    wb = Workbook()
    wb.remove(wb.active)

    def write_block(ws, top_row, rows, header=True):
        width = max(len(r) for r in rows)
        for r_i, row in enumerate(rows):
            for c_i in range(width):
                raw = row[c_i] if c_i < len(row) else None
                val, fmt = coerce_cell(_clean(raw))
                cell = ws.cell(row=top_row + r_i, column=c_i + 1, value=val)
                if fmt:
                    cell.number_format = fmt
                cell.border = border
                cell.alignment = Alignment(wrap_text=True, vertical="top")
                cell.font = Font(name="Arial", size=10, bold=(header and r_i == 0))
                if header and r_i == 0:
                    cell.fill = head_fill
        return top_row + len(rows)

    def autosize(ws):
        for col in ws.columns:
            longest = max((len(str(c.value).split("\n")[0]) for c in col if c.value is not None), default=8)
            ws.column_dimensions[get_column_letter(col[0].column)].width = min(60, max(8, longest + 2))

    with pdfplumber.open(str(src)) as pdf:
        for i in parse_pages(pages, len(pdf.pages)):
            page = pdf.pages[i]
            tables = []
            if table_strategy in ("auto", "lines"):
                tables = page.find_tables()
            if table_strategy == "text" or (table_strategy == "auto" and not tables and False):
                tables = page.find_tables(_TEXT_STRATEGY)
            extracted = [t.extract() for t in tables]
            extracted = [t for t in extracted if t and any(any(c for c in r) for r in t)]

            blocks = []
            if extracted:
                blocks = [(f"Table {n + 1}", t, True) for n, t in enumerate(extracted)]
            elif text_fallback:
                rows = _text_rows(page)
                if rows:
                    blocks = [("Text", rows, False)]
            if not blocks:
                continue

            if sheet_mode == "page":
                ws = wb.create_sheet(f"Page {i + 1}")
                row = 1
                for title, rows, hdr in blocks:
                    if len(blocks) > 1:
                        ws.cell(row=row, column=1, value=title).font = Font(name="Arial", bold=True, size=11)
                        row += 1
                    row = write_block(ws, row, rows, hdr) + 1
                autosize(ws)
            else:
                for n, (title, rows, hdr) in enumerate(blocks, 1):
                    ws = wb.create_sheet(f"P{i + 1}_{'T' + str(n) if hdr else 'Text'}")
                    write_block(ws, 1, rows, hdr)
                    autosize(ws)

    if not wb.sheetnames:
        ws = wb.create_sheet("No data")
        ws["A1"] = "No tables or extractable text found. The PDF may be a scan (OCR it first)."
    wb.save(str(dst))
    return dst
