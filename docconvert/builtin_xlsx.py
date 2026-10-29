"""Excel -> PDF without LibreOffice: cell fonts (name/size/colour/bold/italic/underline/strike), fills,
borders, alignment (incl. wrap, indent, vertical), merged cells, number formats, exact column widths and
row heights, hidden rows/columns, and fit-to-width scaling."""
from __future__ import annotations

import datetime as dt
import re

from .builtin_fonts import THEME, THEME_INDEX, esc, face, families, tint

_INDEXED = {0: "000000", 1: "FFFFFF", 2: "FF0000", 3: "00FF00", 4: "0000FF", 5: "FFFF00", 6: "FF00FF", 7: "00FFFF",
            8: "000000", 9: "FFFFFF", 10: "FF0000", 11: "00FF00", 12: "0000FF", 13: "FFFF00", 14: "FF00FF", 15: "00FFFF",
            16: "800000", 17: "008000", 18: "000080", 19: "808000", 20: "800080", 21: "008080", 22: "C0C0C0", 23: "808080"}


def xl_color(c, default=None):
    """openpyxl Color -> 'RRGGBB' (rgb, theme+tint or indexed)."""
    try:
        if c is None:
            return default
        if c.type == "rgb" and isinstance(c.rgb, str) and len(c.rgb) in (6, 8):
            return c.rgb[-6:]
        if c.type == "theme" and c.theme is not None and c.theme < len(THEME_INDEX):
            base = THEME[THEME_INDEX[c.theme]]
            return tint(base, c.tint) if c.tint else base
        if c.type == "indexed" and c.indexed in _INDEXED:
            return _INDEXED[c.indexed]
    except Exception:
        pass
    return default


_DATE_TOKENS = [("yyyy", "%Y"), ("yy", "%y"), ("mmmm", "%B"), ("mmm", "%b"), ("dddd", "%A"), ("ddd", "%a"),
                ("dd", "%d"), ("hh", "%H"), ("ss", "%S"), ("AM/PM", "%p"), ("mm", "%m"), ("d", "%-d"), ("m", "%-m")]


def format_value(v, fmt: str) -> str:
    """Render a cell value the way Excel's number format would show it (common formats)."""
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    fmt = fmt or "General"
    section = fmt.split(";")[0]
    if isinstance(v, (dt.datetime, dt.date, dt.time)):
        if section == "General" or not re.search(r"[ymdhs]", section, re.I):
            section = "yyyy-mm-dd" if not isinstance(v, dt.time) else "hh:mm:ss"
        py, i, out = re.sub(r'"[^"]*"|\\.', lambda m: m.group(0), section), 0, ""
        s = section.replace("mm", "mm")
        # minutes follow hours
        s = re.sub(r"(h+:?)mm", lambda m: m.group(1) + "MIN", s, flags=re.I)
        for tok, repl in _DATE_TOKENS:
            s = re.sub(tok, repl.replace("%-", "%") if tok in ("d", "m") else repl, s) if tok not in ("d", "m") else s
        s = s.replace("MIN", "%M").replace("AM/PM", "%p")
        s = re.sub(r"(?<![%\w])d(?!\w)", "%d", s)
        try:
            return v.strftime(re.sub(r"[\"\\]", "", s)) if not isinstance(v, dt.date) or isinstance(v, dt.datetime) else dt.datetime(v.year, v.month, v.day).strftime(re.sub(r"[\"\\]", "", s))
        except Exception:
            return str(v)
    if isinstance(v, (int, float)):
        if section in ("General", "@", ""):
            return f"{v:.10g}" if isinstance(v, float) else str(v)
        pct = "%" in section
        num = v * 100 if pct else v
        m = re.search(r"0+\.(0+|#+)|0\.0+|#,##0\.(0+)|\.0+", section)
        dec = len(re.search(r"\.([0#]+)", section).group(1)) if re.search(r"\.([0#]+)", section) else 0
        body = f"{abs(num):,.{dec}f}" if "," in section else f"{abs(num):.{dec}f}"
        pre = "".join(re.findall(r'"([^"]*)"|([$€£¥])', section.split("0")[0].split("#")[0]) and
                      [a or b for a, b in re.findall(r'"([^"]*)"|([$€£¥])', section.split("0")[0].split("#")[0])]) or ""
        suf = "%" if pct else ""
        text = f"{pre}{body}{suf}"
        neg = num < 0
        if neg:
            text = f"({text})" if "(" in section else f"-{text}"
        return text
    return str(v)


def _border(side):
    if side is None or side.style is None:
        return None
    width = {"hair": .25, "thin": .5, "dotted": .5, "dashed": .5, "medium": 1.2, "mediumDashed": 1.2, "thick": 2.0, "double": 1.5}.get(side.style, .5)
    return width, xl_color(side.color, "000000")


def _col_pts(width):
    w = width if width is not None else 8.43
    return max((w * 7 + 5) * 0.75, 4)


def xlsx_to_pdf(src, dst):
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, letter, landscape
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.platypus import PageBreak, Paragraph, SimpleDocTemplate, Spacer
    from reportlab.platypus import Table as RTable, TableStyle

    families()
    wb = load_workbook(str(src), data_only=True)
    story = []
    first_page = None
    for ws in wb.worksheets:
        if ws.sheet_state != "visible" or ws.max_row == 0:
            continue
        min_r, max_r, min_c, max_c = ws.min_row, ws.max_row, ws.min_column, ws.max_column
        if ws.max_row == 1 and ws.max_column == 1 and ws["A1"].value is None:
            continue
        hidden_cols = {c for c, d in ws.column_dimensions.items() if d.hidden}
        widths_by_idx = {}
        for key, d in ws.column_dimensions.items():
            lo, hi = (d.min or 0), (d.max or 0)
            if lo and hi:
                for i in range(lo, min(hi, max_c) + 1):
                    widths_by_idx[i] = d
        cols = [c for c in range(min_c, max_c + 1) if not (widths_by_idx.get(c) is not None and widths_by_idx[c].hidden)]
        rows = [r for r in range(min_r, max_r + 1) if not (ws.row_dimensions[r].hidden)]
        col_pos = {c: i for i, c in enumerate(cols)}
        row_pos = {r: i for i, r in enumerate(rows)}
        merged = {(mr.min_row, mr.min_col): mr for mr in ws.merged_cells.ranges}
        base_w = [_col_pts(widths_by_idx[c].width if c in widths_by_idx else ws.sheet_format.defaultColWidth) for c in cols]
        from reportlab.pdfbase.pdfmetrics import stringWidth

        def text_w(cell):
            f = cell.font
            return stringWidth(format_value(cell.value, cell.number_format).split("\n")[0], face(f.name, bool(f.b), bool(f.i)), f.sz or 11) + 6

        for c in cols:
            need = 0
            for r in rows:
                cell = ws.cell(row=r, column=c)
                if cell.value is None or cell.alignment.wrap_text or (r, c) in merged:
                    continue
                nxt_filled = ws.cell(row=r, column=c + 1).value is not None if c + 1 <= max_c else True
                if isinstance(cell.value, (int, float, dt.date, dt.datetime)) or nxt_filled:
                    need = max(need, text_w(cell))
            base_w[col_pos[c]] = max(base_w[col_pos[c]], min(need, 320))
        total = sum(base_w)
        orient_land = (ws.page_setup.orientation == "landscape") or total > A4[0] - 60
        paper = letter if ws.page_setup.paperSize == 1 else A4
        page = landscape(paper) if orient_land else paper
        if first_page is None:
            first_page = page
        avail = page[0] - 50
        k = min(1.0, avail / total) if total else 1.0
        k = max(k, 0.35)
        widths = [w * k for w in base_w]

        covered = set()
        for mr in ws.merged_cells.ranges:
            merged[(mr.min_row, mr.min_col)] = mr
            for r in range(mr.min_row, mr.max_row + 1):
                for c in range(mr.min_col, mr.max_col + 1):
                    if (r, c) != (mr.min_row, mr.min_col):
                        covered.add((r, c))

        data, cmds, heights = [], [], []
        for r in rows:
            row_cells = []
            dim_h = ws.row_dimensions[r].height
            for c in cols:
                cell = ws.cell(row=r, column=c)
                f = cell.font
                size = (f.sz or 11) * k
                fg = xl_color(f.color, "000000")
                is_num = isinstance(cell.value, (int, float)) and not isinstance(cell.value, bool)
                text = "" if (r, c) in covered else format_value(cell.value, cell.number_format)
                inner = esc(text).replace("\n", "<br/>")
                if f.u:
                    inner = f"<u>{inner}</u>"
                if f.strike:
                    inner = f"<strike>{inner}</strike>"
                h = cell.alignment.horizontal
                align = {"center": 1, "centerContinuous": 1, "right": 2, "justify": 4, "left": 0}.get(h)
                if align is None:
                    align = 2 if (is_num or isinstance(cell.value, (dt.date, dt.datetime))) else 0
                indent = (cell.alignment.indent or 0) * 6 * k
                style = ParagraphStyle("c", fontName=face(f.name, bool(f.b), bool(f.i)), fontSize=size, leading=size * 1.2,
                                       textColor=colors.HexColor("#" + fg), alignment=align,
                                       leftIndent=indent if align == 0 else 0, splitLongWords=1, wordWrap="CJK" if cell.alignment.wrap_text else None)
                if not cell.alignment.wrap_text:
                    style.wordWrap = None
                row_cells.append(Paragraph(inner, style) if inner else "")
                ri, ci = row_pos[r], col_pos[c]
                if inner and align == 0 and not cell.alignment.wrap_text and (r, c) not in merged and (r, c) not in covered:
                    have, j = base_w[ci], ci + 1
                    while have < text_w(cell) and j < len(cols):
                        nc = cols[j]
                        if ws.cell(row=r, column=nc).value is not None or (r, nc) in covered or (r, nc) in merged:
                            break
                        have += base_w[j]
                        j += 1
                    if j - 1 > ci and have >= text_w(cell) * 0.98:
                        cmds.append(("SPAN", (ci, ri), (j - 1, ri)))
                if cell.fill is not None and cell.fill.fill_type == "solid":
                    fill = xl_color(cell.fill.fgColor)
                    if fill:
                        cmds.append(("BACKGROUND", (ci, ri), (ci, ri), colors.HexColor("#" + fill)))
                va = {"top": "TOP", "center": "MIDDLE", "bottom": "BOTTOM"}.get(cell.alignment.vertical, "BOTTOM")
                cmds.append(("VALIGN", (ci, ri), (ci, ri), va))
                b = cell.border
                for side, name, pos in ((b.top, "LINEABOVE", (ci, ri)), (b.bottom, "LINEBELOW", (ci, ri)),
                                        (b.left, "LINEBEFORE", (ci, ri)), (b.right, "LINEAFTER", (ci, ri))):
                    spec = _border(side)
                    if spec:
                        cmds.append((name, pos, pos, spec[0], colors.HexColor("#" + spec[1])))
            data.append(row_cells)
            heights.append(dim_h * k if dim_h else None)
        for (r, c), mr in merged.items():
            if r in row_pos and c in col_pos:
                end_r = max((x for x in rows if x <= mr.max_row), default=r)
                end_c = max((x for x in cols if x <= mr.max_col), default=c)
                cmds.append(("SPAN", (col_pos[c], row_pos[r]), (col_pos[end_c], row_pos[end_r])))
        if not data:
            continue
        if story:
            story.append(PageBreak())
        t = RTable(data, colWidths=widths, rowHeights=heights, repeatRows=0)
        t.setStyle(TableStyle([("LEFTPADDING", (0, 0), (-1, -1), 2), ("RIGHTPADDING", (0, 0), (-1, -1), 2),
                               ("TOPPADDING", (0, 0), (-1, -1), 1), ("BOTTOMPADDING", (0, 0), (-1, -1), 1)] + cmds))
        story.append(t)
    if not story:
        story = [Spacer(1, 1)]
    SimpleDocTemplate(str(dst), pagesize=first_page or landscape(A4), leftMargin=25, rightMargin=25, topMargin=30, bottomMargin=30).build(story)
