"""PDF -> Excel, keeping the table's look.

For every ruled table the cell grid is rebuilt from the PDF geometry, so merged cells, column widths and row
heights match. Each cell gets the formatting found in the PDF: font name, size, bold, italic, colour (also
mixed formatting inside one cell as rich text), background fill, borders (colour and weight), horizontal and
vertical alignment. Numbers, percentages and thousands separators become real numeric cells.
Pages without ruled tables fall back to text lines split into columns, each keeping its font/size/colour.
"""
from __future__ import annotations

import re
from collections import Counter

from .utils import (chars_to_segments, clean_font, coerce_cell, color_hex, font_is_bold, font_is_italic,
                    group_runs, parse_pages)

_TEXT_STRATEGY = {"vertical_strategy": "text", "horizontal_strategy": "text"}


def _snap(values, tol=1.8):
    reps = []
    for v in sorted(values):
        if not reps or v - reps[-1] > tol:
            reps.append(v)
    return reps


def _idx(reps, v):
    return min(range(len(reps)), key=lambda i: abs(reps[i] - v))


def _cell_text(chars):
    """chars -> text with line breaks, inserting spaces where the PDF only has gaps."""
    lines = []
    for ch in sorted(chars, key=lambda c: (round(c["top"] / 2), c["x0"])):
        if lines and abs(lines[-1][0] - ch["top"]) < 2.5:
            lines[-1][1].append(ch)
        else:
            lines.append([ch["top"], [ch]])
    out = []
    for _, row in lines:
        row.sort(key=lambda c: c["x0"])
        s, prev = "", None
        for c in row:
            if prev is not None and c["x0"] - prev["x1"] > max(c.get("size", 9), 1) * 0.18 and c["text"] != " " and prev["text"] != " ":
                s += " "
            s += c["text"]
            prev = c
        out.append(s.strip())
    return "\n".join(t for t in out if t)


def _dominant(chars):
    runs = group_runs([(c["text"], c) for c in chars if c["text"].strip()])
    if not runs:
        return None
    weight = Counter()
    for r in runs:
        weight[(r["font"], r["bold"], r["italic"], r["size"], r["color"])] += len(r["text"])
    font, bold, italic, size, color = weight.most_common(1)[0][0]
    return dict(font=font, bold=bold, italic=italic, size=size, color=color)


def _rich(chars, base):
    """Per-run formatting inside one cell when it differs from the dominant style."""
    runs = group_runs([(c["text"], c) for c in sorted(chars, key=lambda c: (round(c["top"] / 2), c["x0"]))])
    if len({(r["bold"], r["italic"], r["size"], r["color"], r["font"]) for r in runs if r["text"].strip()}) < 2:
        return None
    try:
        from openpyxl.cell.rich_text import CellRichText, TextBlock
        from openpyxl.cell.text import InlineFont
        blocks = []
        for r in runs:
            blocks.append(TextBlock(InlineFont(rFont=r["font"] or base["font"] or "Arial", sz=r["size"], b=r["bold"], i=r["italic"],
                                               color="FF" + r["color"]), r["text"]))
        return CellRichText(*blocks)
    except Exception:
        return None


def _fill_for(page, bbox):
    cx, cy = (bbox[0] + bbox[2]) / 2, (bbox[1] + bbox[3]) / 2
    best = None
    for r in page.rects:
        if not r.get("fill"):
            continue
        if r["x0"] - 0.5 <= cx <= r["x1"] + 0.5 and r["top"] - 0.5 <= cy <= r["bottom"] + 0.5:
            area = (r["x1"] - r["x0"]) * (r["bottom"] - r["top"])
            if area > 0.85 * page.width * page.height:
                continue
            hx = color_hex(r.get("non_stroking_color"))
            if hx == "FFFFFF":
                hx = None
            if best is None or area <= best[0]:
                best = (area, hx)
    return best[1] if best else None


def _table_lines(page, tbbox):
    """(colour, width) most used by the stroked lines/rects inside a table."""
    cols, widths = Counter(), []
    for obj in list(page.lines) + [r for r in page.rects if r.get("stroke")]:
        if tbbox[0] - 2 <= obj["x0"] <= tbbox[2] + 2 and tbbox[1] - 2 <= obj["top"] <= tbbox[3] + 2:
            cols[color_hex(obj.get("stroking_color"))] += 1
            widths.append(obj.get("linewidth") or 0.5)
    colour = cols.most_common(1)[0][0] if cols else "000000"
    w = sorted(widths)[len(widths) // 2] if widths else 0.5
    return colour, w


def _align(chars, bbox):
    xs0, xs1 = min(c["x0"] for c in chars), max(c["x1"] for c in chars)
    ys0, ys1 = min(c["top"] for c in chars), max(c["bottom"] for c in chars)
    width, height = bbox[2] - bbox[0], bbox[3] - bbox[1]
    gl, gr = xs0 - bbox[0], bbox[2] - xs1
    if abs(gl - gr) < 0.12 * width and gl > 3:
        h = "center"
    elif gr < gl * 0.5 and gr < 8:
        h = "right"
    else:
        h = "left"
    gt, gb = ys0 - bbox[1], bbox[3] - ys1
    v = "center" if abs(gt - gb) < 0.2 * height else ("top" if gt < gb else "bottom")
    return h, v


def pdf_to_xlsx(src, dst, *, pages=None, sheet_mode: str = "page", table_strategy: str = "auto",
                text_fallback: bool = True, **_):
    """
    sheet_mode:      ``"page"`` (one sheet per page, tables stacked) or ``"table"`` (one sheet per table)
    table_strategy:  ``"auto"``/``"lines"`` (ruled tables) or ``"text"`` (infer from alignment)
    text_fallback:   when a page has no table, put its text lines into rows/columns
    """
    import pdfplumber
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    if sheet_mode not in ("page", "table"):
        raise ValueError("sheet_mode must be 'page' or 'table'")
    wb = Workbook()
    wb.remove(wb.active)

    def style_cell(cell, st, fill=None, align=None, borders=None):
        cell.font = Font(name=(st or {}).get("font") or "Arial", size=(st or {}).get("size") or 10,
                         bold=bool((st or {}).get("bold")), italic=bool((st or {}).get("italic")),
                         color="FF" + ((st or {}).get("color") or "000000"))
        if fill:
            cell.fill = PatternFill("solid", fgColor="FF" + fill)
        if align:
            cell.alignment = Alignment(horizontal=align[0], vertical=align[1], wrap_text=True)
        if borders:
            side = Side(style=borders[1], color="FF" + borders[0])
            cell.border = Border(left=side, right=side, top=side, bottom=side)

    def write_table(ws, row0, page, tbl):
        cells = [c for c in tbl.cells if c]
        if not cells:
            return row0
        xs = _snap([c[0] for c in cells] + [c[2] for c in cells])
        ys = _snap([c[1] for c in cells] + [c[3] for c in cells])
        lcol, lw = _table_lines(page, tbl.bbox)
        bstyle = "medium" if lw >= 1.4 else "thin"
        used = {}
        for bb in cells:
            c0, c1, r0, r1 = _idx(xs, bb[0]), _idx(xs, bb[2]), _idx(ys, bb[1]), _idx(ys, bb[3])
            if c1 <= c0 or r1 <= r0:
                continue
            inside = [ch for ch in page.chars if bb[0] - 0.5 <= (ch["x0"] + ch["x1"]) / 2 <= bb[2] + 0.5
                      and bb[1] - 0.5 <= (ch["top"] + ch["bottom"]) / 2 <= bb[3] + 0.5]
            text = _cell_text(inside)
            st = _dominant(inside)
            fill = _fill_for(page, bb)
            al = _align([c for c in inside if c["text"].strip()], bb) if text else ("left", "center")
            val, fmt = coerce_cell(text)
            cell = ws.cell(row=row0 + r0, column=c0 + 1)
            rich = _rich(inside, st) if st and isinstance(val, str) and val else None
            cell.value = rich if rich is not None else val
            if fmt:
                cell.number_format = fmt
            style_cell(cell, st, fill, al, (lcol, bstyle))
            if c1 - c0 > 1 or r1 - r0 > 1:
                ws.merge_cells(start_row=row0 + r0, start_column=c0 + 1, end_row=row0 + r1 - 1, end_column=c1)
            used[(r0, c0)] = bb
        for i in range(len(xs) - 1):
            letter = get_column_letter(i + 1)
            w = (xs[i + 1] - xs[i]) / 5.4
            ws.column_dimensions[letter].width = max(ws.column_dimensions[letter].width or 0, min(max(w, 4), 90))
        for j in range(len(ys) - 1):
            ws.row_dimensions[row0 + j].height = max(12, ys[j + 1] - ys[j])
        return row0 + len(ys) + 1

    def write_lines(ws, row0, page):
        lines = page.extract_text_lines(return_chars=True, strip=True)
        r = row0
        for ln in lines:
            segs = chars_to_segments(ln["chars"], gap_factor=1.0)
            for ci, seg in enumerate(segs, 1):
                runs = group_runs(seg)
                text = "".join(x["text"] for x in runs).strip()
                big = max(runs, key=lambda x: len(x["text"]))
                val, fmt = coerce_cell(text)
                cell = ws.cell(row=r, column=ci, value=val)
                if fmt:
                    cell.number_format = fmt
                bb = (min(c["x0"] for _, c in seg), ln["top"], max(c["x1"] for _, c in seg), ln["bottom"])
                style_cell(cell, big, _fill_for(page, bb))
            r += 1
        for col in ws.columns:
            longest = max((len(str(c.value).split("\n")[0]) for c in col if c.value is not None), default=8)
            ws.column_dimensions[get_column_letter(col[0].column)].width = min(70, max(8, longest * 1.15 + 2))
        return r

    with pdfplumber.open(str(src)) as pdf:
        for i in parse_pages(pages, len(pdf.pages)):
            page = pdf.pages[i]
            tables = []
            if table_strategy in ("auto", "lines"):
                tables = page.find_tables()
            elif table_strategy == "text":
                tables = page.find_tables(_TEXT_STRATEGY)
            tables = [t for t in tables if t.cells]

            def coverage(tbls):
                if not page.chars:
                    return 0.0
                inside = sum(1 for ch in page.chars if any(t.bbox[0] - 1 <= ch["x0"] and ch["x1"] <= t.bbox[2] + 1
                                                          and t.bbox[1] - 1 <= ch["top"] and ch["bottom"] <= t.bbox[3] + 1 for t in tbls))
                return inside / len(page.chars)

            if tables and coverage(tables) < 0.5:          # shading boxes, not a real grid of cells
                tables = []
            if not tables and table_strategy == "auto":      # borderless tables: infer columns from alignment
                guess = [t for t in page.find_tables(_TEXT_STRATEGY) if t.cells and len(t.rows) >= 2]
                if guess and coverage(guess) >= 0.7 and max(len(t.rows[0].cells) for t in guess) >= 2:
                    tables = guess
            if tables:
                if sheet_mode == "page":
                    ws = wb.create_sheet(f"Page {i + 1}")
                    row = 1
                    for t in tables:
                        row = write_table(ws, row, page, t)
                else:
                    for n, t in enumerate(tables, 1):
                        ws = wb.create_sheet(f"P{i + 1}_T{n}")
                        write_table(ws, 1, page, t)
            elif text_fallback and page.chars:
                ws = wb.create_sheet(f"Page {i + 1}")
                write_lines(ws, 1, page)

    if not wb.sheetnames:
        ws = wb.create_sheet("No data")
        ws["A1"] = "No tables or extractable text found. The PDF may be a scan (OCR it first)."
    wb.save(str(dst))
    return dst
