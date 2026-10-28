"""PDF -> Word, exact-layout mode (for forms, letterheads, certificates, anything positional).

Flow-based conversion rebuilds a document as paragraphs and tables, which breaks
forms: letters typed into comb boxes drift out of their cells, rotated side labels
are mangled and fields overlap. This mode keeps every page exactly where it was:

* all lines, boxes, logos, check marks and rotated text stay in a page-sized
  picture placed *behind* the text (the text itself is cut out of that picture);
* every piece of real horizontal text becomes an editable Word text frame that is
  pinned to the same coordinates, with the original font, size, colour, bold and italic;
* letters typed into comb boxes get one frame per letter so they stay inside their cells.

Needs PyMuPDF (installed with the ``fidelity`` extra).
"""
from __future__ import annotations

import re
from statistics import median

from .errors import MissingDependencyError
from .utils import clean_font, font_is_bold, font_is_italic, parse_pages

_SYMBOLIC = re.compile(r"dingbat|symbol|wingding|webding|zapf", re.I)


def _fitz():
    try:
        import fitz  # PyMuPDF
    except ImportError as exc:
        raise MissingDependencyError("Exact-layout mode needs PyMuPDF: pip install pymupdf") from exc
    return fitz


def looks_like_form(src, pages=None) -> bool:
    """Heuristic used by ``layout="auto"``: form fields, comb boxes or heavy line art."""
    try:
        fitz = _fitz()
    except MissingDependencyError:
        return False
    doc = fitz.open(str(src))
    try:
        for i in parse_pages(pages, len(doc))[:3]:
            page = doc[i]
            if list(page.widgets() or []):
                return True
            comb = 0
            for b in page.get_text("rawdict")["blocks"]:
                for ln in b.get("lines", []):
                    if abs(ln["dir"][1]) > 0.02:
                        continue
                    for sp in ln["spans"]:
                        if len(sp["chars"]) == 1 and sp["chars"][0]["c"].strip():
                            comb += 1
            if comb >= 25 or len(page.get_drawings()) >= 400:
                return True
        return False
    finally:
        doc.close()


def _rgb_int(c) -> int:
    try:
        if isinstance(c, (int, float)):
            c = (c, c, c)
        c = tuple(c)
        if len(c) == 1:
            c = c * 3
        return (round(c[0] * 255) << 16) | (round(c[1] * 255) << 8) | round(c[2] * 255)
    except Exception:
        return 0


def _overlap(a, b) -> float:
    w = min(a[2], b[2]) - max(a[0], b[0])
    h = min(a[3], b[3]) - max(a[1], b[1])
    if w <= 0 or h <= 0:
        return 0.0
    small = min((a[2] - a[0]) * (a[3] - a[1]), (b[2] - b[0]) * (b[3] - b[1])) or 1
    return w * h / small


def _collect(page):
    """Text worth rebuilding as Word text, in drawing order.

    Uses the text trace (which knows drawing order) so that text hidden underneath
    later text - common in edited forms - is dropped instead of being doubled.
    """
    items = []
    for t in page.get_texttrace():
        if t.get("type") == 3 or t.get("opacity", 1) == 0:
            continue
        dx, dy = t.get("dir", (1, 0))
        if abs(dy) > 0.02 or dx <= 0:                      # rotated text stays in the picture
            continue
        chars = [dict(c=chr(c[0]), bbox=tuple(c[3]), oy=c[2][1]) for c in t["chars"]]
        if _SYMBOLIC.search(t["font"]):
            continue                                       # symbols, boxes stay in the picture
        size = t["size"]
        # one text object can hold several lines / far-apart words: split into visual pieces
        pieces, cur = [], []
        for c in chars:
            if cur:
                pv = cur[-1]
                if (abs(c["oy"] - pv["oy"]) > 0.4 * size or c["bbox"][0] - pv["bbox"][2] > 1.5 * size
                        or c["bbox"][0] - pv["bbox"][2] < -2 * size):
                    pieces.append(cur)
                    cur = []
            cur.append(c)
        if cur:
            pieces.append(cur)
        for pc in pieces:
            while pc and not pc[0]["c"].strip():
                pc.pop(0)
            while pc and not pc[-1]["c"].strip():
                pc.pop()
            text = "".join(c["c"] for c in pc)
            if not text or not re.search(r"\w", text):
                continue                                   # dashes and lone symbols stay in the picture
            bb = (min(c["bbox"][0] for c in pc), min(c["bbox"][1] for c in pc),
                  max(c["bbox"][2] for c in pc), max(c["bbox"][3] for c in pc))
            items.append(dict(chars=pc, text=text, size=size, font=t["font"], flags=t.get("flags", 0),
                              color=_rgb_int(t.get("color", 0)), y=round(pc[0]["oy"], 1), bbox=bb,
                              seq=t.get("seqno", 0)))
    kept = []
    for it in sorted(items, key=lambda s: -s["seq"]):      # newest first; older overlapped text is hidden
        if not any(_overlap(it["bbox"], k["bbox"]) > 0.15 for k in kept):
            kept.append(it)
    items = kept
    # merge neighbouring spans of the same style on one baseline (but not comb letters)
    items.sort(key=lambda s: (s["y"], s["bbox"][0]))
    merged = []
    for it in items:
        p = merged[-1] if merged else None
        if (p and abs(p["y"] - it["y"]) < 1.0 and p["font"] == it["font"] and abs(p["size"] - it["size"]) < .3
                and p["color"] == it["color"] and 0 <= it["bbox"][0] - p["bbox"][2] < 0.35 * it["size"]):
            gap = it["bbox"][0] - p["bbox"][2]
            if gap > 0.15 * it["size"] and not p["text"].endswith(" "):
                p["text"] += " "
                p["chars"] = p["chars"] + [dict(c=" ", bbox=(p["bbox"][2], p["bbox"][1], it["bbox"][0], p["bbox"][3]))]
            p["text"] += it["text"]
            p["chars"] = p["chars"] + it["chars"]
            p["bbox"] = (p["bbox"][0], min(p["bbox"][1], it["bbox"][1]), it["bbox"][2], max(p["bbox"][3], it["bbox"][3]))
        else:
            merged.append(dict(it))
    spans = []
    for s in merged:
        ch = s["chars"]
        gaps = [ch[k + 1]["bbox"][0] - ch[k]["bbox"][2] for k in range(len(ch) - 1) if ch[k]["c"] != " " and ch[k + 1]["c"] != " "]
        comb = len(ch) >= 3 and gaps and median(gaps) > 0.35 * s["size"]
        if comb:
            for c in ch:
                if c["c"].strip():
                    spans.append(dict(s, text=c["c"], bbox=tuple(c["bbox"])))
        else:
            spans.append(s)
    return spans


def _visible_spans(fitz, page, spans, dpi: int = 144):
    """Cut the spans out of the page and keep only those whose removal changes the picture,
    i.e. text that is really visible (not covered by a later white box or clipped away)."""
    import io

    from PIL import Image, ImageChops

    def snap():
        pix = page.get_pixmap(dpi=dpi, alpha=False)
        return Image.frombytes("RGB", (pix.width, pix.height), pix.samples)

    before = snap()
    for s in spans:
        page.add_redact_annot(fitz.Rect(s["bbox"]) + (0.3, 0.3, -0.3, -0.3), fill=False)
    page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE, graphics=fitz.PDF_REDACT_LINE_ART_NONE)
    after = snap()
    k = dpi / 72.0
    keep = []
    for s in spans:
        x0, y0, x1, y1 = s["bbox"]
        ix, iy = min(1.5, 0.2 * (x1 - x0)), min(1.0, 0.2 * (y1 - y0))   # ignore ink of neighbours at the edges
        box = (max(int((x0 + ix) * k), 0), max(int((y0 + iy) * k), 0),
               min(int((x1 - ix) * k) + 1, before.width), min(int((y1 - iy) * k) + 1, before.height))
        if box[2] <= box[0] or box[3] <= box[1]:
            continue
        diff = ImageChops.difference(before.crop(box), after.crop(box)).convert("L")
        if diff.getextrema()[1] >= 40:
            keep.append(s)
    return keep


def _frame(doc, s, page_w, dy):
    from docx.enum.text import WD_LINE_SPACING
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Pt, RGBColor

    x0, y0, x1, y1 = s["bbox"]
    h = max(y1 - y0, s["size"] * 1.1)
    w = min((x1 - x0) * 1.3 + 8, max(page_w - x0, x1 - x0 + 2))
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = pf.space_after = Pt(0)
    pf.left_indent = pf.right_indent = Pt(0)
    pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    pf.line_spacing = Pt(h)
    fr = OxmlElement("w:framePr")
    for k, v in (("w", int(w * 20)), ("h", int(h * 20) + 20), ("hRule", "exact"), ("wrap", "around"),
                 ("vAnchor", "page"), ("hAnchor", "page"), ("x", int(x0 * 20)), ("y", int((y0 + dy) * 20))):
        fr.set(qn("w:" + k), str(v))
    p._p.get_or_add_pPr().insert(0, fr)
    run = p.add_run(s["text"])
    f = run.font
    f.size = Pt(s["size"])
    name = clean_font(s["font"])
    if name:
        f.name = name
        rpr = run._r.get_or_add_rPr()
        rf = rpr.find(qn("w:rFonts"))
        rf.set(qn("w:cs"), name)
        rf.set(qn("w:eastAsia"), name)
    f.bold = True if (font_is_bold(s["font"]) or s["flags"] & 16) else None
    f.italic = True if (font_is_italic(s["font"]) or s["flags"] & 2) else None
    f.color.rgb = RGBColor.from_string("%06X" % (s["color"] & 0xFFFFFF))


def _background(doc, png: bytes, w_pt, h_pt, first: bool):
    """Page-sized picture anchored to the page, behind the text."""
    import io

    from docx.enum.text import WD_LINE_SPACING
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls
    from docx.shared import Emu, Pt

    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = pf.space_after = Pt(0)
    pf.line_spacing_rule, pf.line_spacing = WD_LINE_SPACING.EXACTLY, Pt(1)
    pf.page_break_before = not first
    run = p.add_run()
    run.font.size = Pt(1)
    run.add_picture(io.BytesIO(png), width=Emu(int(w_pt * 12700)), height=Emu(int(h_pt * 12700)))
    inline = run._r.xpath(".//wp:inline")[0]
    anchor = parse_xml(
        f'<wp:anchor {nsdecls("wp")} distT="0" distB="0" distL="0" distR="0" simplePos="0" '
        'relativeHeight="0" behindDoc="1" locked="0" layoutInCell="1" allowOverlap="1">'
        '<wp:simplePos x="0" y="0"/>'
        '<wp:positionH relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionH>'
        '<wp:positionV relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionV>'
        '<wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/></wp:anchor>')
    ext = inline.find("{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}extent")
    anchor.insert(2, ext)
    anchor.insert(4, inline.find("{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}docPr"))
    anchor.append(inline.find("{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}cNvGraphicFramePr"))
    anchor.append(inline.find("{http://schemas.openxmlformats.org/drawingml/2006/main}graphic"))
    inline.getparent().replace(inline, anchor)


def pdf_to_docx_fixed(src, dst, *, pages=None, dpi: int = 200, y_offset: float = 0.0, **_):
    fitz = _fitz()
    from docx import Document
    from docx.shared import Pt

    pdf = fitz.open(str(src))
    try:
        idx = parse_pages(pages, len(pdf))
        first = pdf[idx[0]].rect
        doc = Document()
        sec = doc.sections[0]
        sec.page_width, sec.page_height = Pt(first.width), Pt(first.height)
        sec.left_margin = sec.right_margin = sec.top_margin = sec.bottom_margin = Pt(0)
        sec.header_distance = sec.footer_distance = Pt(0)
        # the template ships with an empty paragraph; remove it so page 1 holds only our content
        for para in list(doc.paragraphs):
            para._p.getparent().remove(para._p)

        for n, i in enumerate(idx):
            page = pdf[i]
            # cut the text out of the picture (lines and boxes stay) and keep only visible text
            spans = _visible_spans(fitz, page, _collect(page))
            png = page.get_pixmap(dpi=dpi, alpha=False).tobytes("png")
            for s in spans:
                _frame(doc, s, page.rect.width, y_offset)
            _background(doc, png, page.rect.width, page.rect.height, first=(n == 0))
        doc.save(str(dst))
    finally:
        pdf.close()
    return dst
