"""Word -> PDF without LibreOffice. Reads the real formatting (run size/colour/font/bold/italic/underline/
strike/highlight, paragraph alignment/indent/spacing, list numbering, table shading/borders/merges,
headers/footers) from the .docx, following style inheritance."""
from __future__ import annotations

import io
import re

from .builtin_fonts import THEME, esc, face, family_kind, families, tint

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
_HIGHLIGHT = {"yellow": "FFFF00", "green": "00FF00", "cyan": "00FFFF", "magenta": "FF00FF", "blue": "0000FF",
              "red": "FF0000", "darkBlue": "000080", "darkCyan": "008080", "darkGreen": "008000",
              "darkMagenta": "800080", "darkRed": "800000", "darkYellow": "808000", "darkGray": "808080",
              "lightGray": "C0C0C0", "black": "000000", "white": "FFFFFF"}


def _chain(style):
    while style is not None:
        yield style
        style = style.base_style


def _theme_rgb(color):
    try:
        if color.rgb is not None:
            return str(color.rgb)
        tc = color.theme_color
        if tc is not None:
            name = {"TEXT_1": "dk1", "BACKGROUND_1": "lt1", "TEXT_2": "dk2", "BACKGROUND_2": "lt2", "ACCENT_1": "accent1",
                    "ACCENT_2": "accent2", "ACCENT_3": "accent3", "ACCENT_4": "accent4", "ACCENT_5": "accent5",
                    "ACCENT_6": "accent6", "HYPERLINK": "hlink", "FOLLOWED_HYPERLINK": "folHlink"}.get(tc.name)
            if name:
                base = THEME[name]
                return tint(base, -0.25) if name.startswith("accent") else base   # Word headings use a darker accent
    except Exception:
        pass
    return None


class _Doc:
    def __init__(self, doc):
        from docx.oxml.ns import qn
        self.doc = doc
        sz = doc.styles.element.xpath("w:docDefaults/w:rPrDefault/w:rPr/w:sz/@w:val")
        self.default_size = int(sz[0]) / 2 if sz else 11.0
        self.counters: dict = {}
        self.numbering = None
        try:
            self.numbering = doc.part.numbering_part.element
        except Exception:
            pass
        self.qn = qn

    # ---- run properties with inheritance
    def run_props(self, run, para_style, extra_chain=()):
        p = dict(size=None, color=None, name=None, bold=None, italic=None, underline=None, strike=None, hl=None, sub=None, sup=None)
        sources = [run.font]
        try:
            if run.style is not None:
                sources += [s.font for s in _chain(run.style)]
        except Exception:
            pass
        sources += [s.font for s in _chain(para_style)]
        for f in sources:
            if p["size"] is None and f.size:
                p["size"] = f.size.pt
            if p["color"] is None:
                p["color"] = _theme_rgb(f.color)
            if p["name"] is None and f.name:
                p["name"] = f.name
            for k, attr in (("bold", "bold"), ("italic", "italic"), ("strike", "strike")):
                if p[k] is None and getattr(f, attr) is not None:
                    p[k] = getattr(f, attr)
            if p["underline"] is None and f.underline is not None:
                p["underline"] = bool(f.underline)
            if p["hl"] is None and f.highlight_color is not None:
                p["hl"] = _HIGHLIGHT.get(f.highlight_color.name.lower().replace("_", ""), None) or _HIGHLIGHT.get(
                    {"DARK_BLUE": "darkBlue", "DARK_RED": "darkRed", "DARK_YELLOW": "darkYellow", "TEAL": "darkCyan",
                     "GRAY_25": "lightGray", "GRAY_50": "darkGray", "BRIGHT_GREEN": "green", "TURQUOISE": "cyan",
                     "PINK": "magenta", "VIOLET": "darkMagenta"}.get(f.highlight_color.name, ""), None)
            if p["sub"] is None and f.subscript is not None:
                p["sub"] = f.subscript
            if p["sup"] is None and f.superscript is not None:
                p["sup"] = f.superscript
        p["size"] = p["size"] or self.default_size
        return p

    def pf(self, para, attr):
        v = getattr(para.paragraph_format, attr)
        if v is not None:
            return v
        for s in _chain(para.style):
            v = getattr(s.paragraph_format, attr)
            if v is not None:
                return v
        return None

    # ---- lists
    def list_marker(self, para):
        ppr = para._p.pPr
        num_id = ilvl = None
        if ppr is not None and ppr.numPr is not None:
            num_id = ppr.numPr.numId.val if ppr.numPr.numId is not None else None
            ilvl = ppr.numPr.ilvl.val if ppr.numPr.ilvl is not None else 0
        name = para.style.name if para.style is not None else ""
        if num_id is None:
            if name.startswith("List Bullet"):
                return "•", 18 * (1 + (int(name[-1]) - 1 if name[-1].isdigit() else 0))
            if name.startswith("List Number"):
                key = ("style", name)
                self.counters[key] = self.counters.get(key, 0) + 1
                return f"{self.counters[key]}.", 18
            return None, 0
        fmt, text, left = "decimal", "%1.", 18 * (ilvl + 1)
        if self.numbering is not None and num_id:
            q = self.qn
            for num in self.numbering.findall(q("w:num")):
                if num.get(q("w:numId")) == str(num_id):
                    aid = num.find(q("w:abstractNumId")).get(q("w:val"))
                    for ab in self.numbering.findall(q("w:abstractNum")):
                        if ab.get(q("w:abstractNumId")) == aid:
                            for lvl in ab.findall(q("w:lvl")):
                                if lvl.get(q("w:ilvl")) == str(ilvl):
                                    f = lvl.find(q("w:numFmt"))
                                    t = lvl.find(q("w:lvlText"))
                                    fmt = f.get(q("w:val")) if f is not None else fmt
                                    text = t.get(q("w:val")) if t is not None else text
                                    ind = lvl.find(q("w:pPr") + "/" + q("w:ind"))
                                    if ind is not None and ind.get(q("w:left")):
                                        left = int(ind.get(q("w:left"))) / 20
        if fmt == "bullet" or (text and ord(text[0]) >= 0xF000):
            return "•" if ilvl == 0 else "–", left
        key = (num_id, ilvl)
        self.counters[key] = self.counters.get(key, 0) + 1
        for k in [k for k in self.counters if isinstance(k[0], int) and k[0] == num_id and k[1] > ilvl]:
            del self.counters[k]
        n = self.counters[key]
        shown = {"lowerLetter": chr(96 + (n - 1) % 26 + 1), "upperLetter": chr(64 + (n - 1) % 26 + 1),
                 "lowerRoman": _roman(n).lower(), "upperRoman": _roman(n)}.get(fmt, str(n))
        return re.sub(r"%\d", shown, text or "%1."), left


def _roman(n):
    out = ""
    for v, s in ((1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"), (90, "XC"), (50, "L"), (40, "XL"),
                 (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")):
        while n >= v:
            out += s
            n -= v
    return out


def _markup(ctx, runs_iter, para_style):
    parts, size = [], 0.0
    for run, link in runs_iter:
        text = run.text
        if not text and not run._r.xpath(".//w:br | .//w:tab"):
            continue
        p = ctx.run_props(run, para_style)
        if link:
            p["color"] = p["color"] or "0563C1"
            p["underline"] = True
        size = max(size, p["size"])
        t = esc(text).replace("\t", "&nbsp;" * 4).replace("\n", "<br/>")
        attrs = [f'name="{face(p["name"], bool(p["bold"]), bool(p["italic"]))}"', f'size="{p["size"]:.1f}"']
        if p["color"]:
            attrs.append(f'color="#{p["color"]}"')
        if p["hl"]:
            attrs.append(f'backcolor="#{p["hl"]}"')
        t = f"<font {' '.join(attrs)}>{t}</font>"
        if p["underline"]:
            t = f"<u>{t}</u>"
        if p["strike"]:
            t = f"<strike>{t}</strike>"
        if p["sup"]:
            t = f"<super>{t}</super>"
        elif p["sub"]:
            t = f"<sub>{t}</sub>"
        if link:
            t = f'<a href="{esc(link).replace(chr(34), "&quot;")}">{t}</a>'
        parts.append(t)
    return "".join(parts), size


def _inner_runs(par):
    """Yield (run, hyperlink_url|None) in order, including runs inside hyperlinks."""
    try:
        for item in par.iter_inner_content():
            if hasattr(item, "runs") and hasattr(item, "url"):
                for r in item.runs:
                    yield r, item.url or "#"
            else:
                yield item, None
    except AttributeError:
        for r in par.runs:
            yield r, None


def paragraph_flowables(ctx, par, avail, in_cell=False):
    from docx.enum.text import WD_LINE_SPACING, WD_PARAGRAPH_ALIGNMENT as A
    from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT, TA_RIGHT
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.utils import ImageReader
    from reportlab.lib import colors
    from reportlab.platypus import Image as RImage, PageBreak, Paragraph, Spacer

    flow = []
    if ctx.pf(par, "page_break_before") and not in_cell:
        flow.append(PageBreak())
    runs = list(_inner_runs(par))
    for run, _ in runs:
        for rid in run._r.xpath(".//a:blip/@r:embed"):
            try:
                blob = ctx.doc.part.related_parts[rid].blob
                iw, ih = ImageReader(io.BytesIO(blob)).getSize()
                ext = run._r.xpath(".//wp:extent")
                w = int(ext[0].get("cx")) / 12700 if ext else iw * 0.75
                h = int(ext[0].get("cy")) / 12700 if ext else ih * 0.75
                k = min(1.0, avail / w) if w else 1
                flow.append(RImage(io.BytesIO(blob), width=w * k, height=h * k))
            except Exception:
                pass
        if run._r.xpath('.//w:br[@w:type="page"]') and not in_cell:
            flow.append(PageBreak())

    style = par.style
    text, size = _markup(ctx, runs, style)
    marker, left = ctx.list_marker(par)
    if not text.strip():
        if not flow:
            flow.append(Spacer(1, (size or ctx.default_size) * 1.1))
        return flow
    size = size or ctx.default_size
    al = ctx.pf(par, "alignment")
    align = {A.CENTER: TA_CENTER, A.RIGHT: TA_RIGHT, A.JUSTIFY: TA_JUSTIFY}.get(al, TA_LEFT)
    ind_left = ctx.pf(par, "left_indent")
    first = ctx.pf(par, "first_line_indent")
    left_pt = (ind_left.pt if ind_left is not None else 0) or (left if marker else 0)
    first_pt = first.pt if first is not None else 0
    ls = ctx.pf(par, "line_spacing")
    rule = ctx.pf(par, "line_spacing_rule")
    leading = size * 1.2
    if ls is not None:
        if isinstance(ls, float) or (isinstance(ls, int) and ls < 10):
            leading = size * 1.2 * float(ls)
        else:
            leading = ls.pt if rule == WD_LINE_SPACING.EXACTLY else max(ls.pt, size * 1.2)
    sb, sa = ctx.pf(par, "space_before"), ctx.pf(par, "space_after")
    st = ParagraphStyle("p", fontName=face(None), fontSize=size, leading=leading, alignment=align,
                        leftIndent=left_pt, firstLineIndent=first_pt if not marker else 0,
                        spaceBefore=sb.pt if sb is not None else 0, spaceAfter=sa.pt if sa is not None else (0 if in_cell else 8),
                        autoLeading="max", bulletIndent=max(left_pt - 18, 0))
    # paragraph shading / border box
    shd = par._p.xpath("w:pPr/w:shd/@w:fill")
    if shd and shd[0] not in ("auto", "FFFFFF"):
        st.backColor = colors.HexColor("#" + shd[0])
    flow.append(Paragraph(text, st, bulletText=marker))
    return flow


def _cell_fill(tc):
    v = tc.xpath("w:tcPr/w:shd/@w:fill")
    return v[0] if v and v[0] not in ("auto", "FFFFFF") else None


def _table_flowable(ctx, tbl, avail):
    from reportlab.lib import colors
    from reportlab.platypus import Spacer
    from reportlab.platypus import Table as RTable, TableStyle
    from docx.table import _Cell

    grid = [int(x) / 20 for x in tbl._tbl.xpath("w:tblGrid/w:gridCol/@w:w")]
    ncols = len(grid) or max(len(r._tr.tc_lst) for r in tbl.rows)
    if not grid:
        grid = [avail / ncols] * ncols
    k = min(1.0, avail / sum(grid)) if sum(grid) else 1
    widths = [w * k for w in grid]
    data, cmds = [], []
    styles_text = (tbl.style.name if tbl.style is not None else "") or ""
    has_borders = bool(tbl._tbl.xpath("w:tblPr/w:tblBorders")) or "grid" in styles_text.lower() or "table grid" in styles_text.lower()
    open_vmerge: dict = {}
    for r_i, row in enumerate(tbl.rows):
        out, col = [""] * ncols, 0
        for tc in row._tr.tc_lst:
            span = int(tc.xpath("w:tcPr/w:gridSpan/@w:val")[0]) if tc.xpath("w:tcPr/w:gridSpan/@w:val") else 1
            vm = tc.xpath("w:tcPr/w:vMerge")
            vm_val = (vm[0].get(ctx.qn("w:val")) if vm else None)
            if col >= ncols:
                break
            if vm and vm_val != "restart":                       # continuation of a vertically merged cell
                start = open_vmerge.get(col)
                if start:
                    cmds.append(("SPAN", (col, start), (col + span - 1, r_i)))
                col += span
                continue
            if vm and vm_val == "restart":
                open_vmerge[col] = r_i
            else:
                open_vmerge.pop(col, None)
            cell = _Cell(tc, tbl)
            cell_w = sum(widths[col:col + span]) - 10
            items = []
            for p in cell.paragraphs:
                items += paragraph_flowables(ctx, p, cell_w, in_cell=True)
            out[col] = items or ""
            fill = _cell_fill(tc)
            if fill:
                cmds.append(("BACKGROUND", (col, r_i), (col + span - 1, r_i), colors.HexColor("#" + fill)))
            if span > 1:
                cmds.append(("SPAN", (col, r_i), (col + span - 1, r_i)))
            col += span
        data.append(out)
    # collapse vertical merges that were registered start->end
    final_cmds = []
    spans = {}
    for c in cmds:
        if c[0] == "SPAN" and c[1][1] != c[2][1]:
            spans[(c[1][0], c[1][1])] = max(spans.get((c[1][0], c[1][1]), 0), c[2][1])
        else:
            final_cmds.append(c)
    for (col, start), end in spans.items():
        final_cmds.append(("SPAN", (col, start), (col, end)))
    t = RTable(data, colWidths=widths, repeatRows=0)
    base = [("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 5), ("RIGHTPADDING", (0, 0), (-1, -1), 5),
            ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3)]
    if has_borders:
        base.append(("GRID", (0, 0), (-1, -1), 0.5, colors.black))
    t.setStyle(TableStyle(base + final_cmds))
    return [t, Spacer(1, 8)]


def docx_to_pdf(src, dst):
    from docx import Document
    from docx.oxml.ns import qn
    from docx.table import Table
    from docx.text.paragraph import Paragraph as DPara
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer

    families()
    doc = Document(str(src))
    ctx = _Doc(doc)
    sec = doc.sections[0]
    pw, ph = sec.page_width.pt, sec.page_height.pt
    lm, rm, tm, bm = sec.left_margin.pt, sec.right_margin.pt, sec.top_margin.pt, sec.bottom_margin.pt
    avail = pw - lm - rm
    story = []
    for child in doc.element.body.iterchildren():
        if child.tag == qn("w:p"):
            story += paragraph_flowables(ctx, DPara(child, doc), avail)
        elif child.tag == qn("w:tbl"):
            story += _table_flowable(ctx, Table(child, doc), avail)
    if not story:
        story = [Spacer(1, 1)]

    def furniture(canvas, _doc):
        for part, y in ((sec.header, ph - sec.header_distance.pt), (sec.footer, sec.footer_distance.pt)):
            try:
                for p in part.paragraphs:
                    text, size = _markup(ctx, [(r, None) for r in p.runs], p.style)
                    if not text.strip():
                        continue
                    from docx.enum.text import WD_PARAGRAPH_ALIGNMENT as A
                    al = {A.CENTER: 1, A.RIGHT: 2}.get(p.alignment, 0)
                    para = Paragraph(text, ParagraphStyle("h", fontName=face(None), fontSize=size or 9, leading=(size or 9) * 1.2, alignment=al))
                    w, h = para.wrap(avail, 100)
                    para.drawOn(canvas, lm, y - h if part is sec.header else y)
                    y += -h if part is sec.header else h
            except Exception:
                pass

    SimpleDocTemplate(str(dst), pagesize=(pw, ph), leftMargin=lm, rightMargin=rm, topMargin=tm, bottomMargin=bm).build(
        story, onFirstPage=furniture, onLaterPages=furniture)
