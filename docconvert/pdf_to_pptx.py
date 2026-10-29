"""PDF -> PowerPoint.

Modes
-----
``auto`` (default)
    Forms and positional pages (boxes, comb cells, heavy line art) use ``exact``; everything else ``editable``.
``exact``
    Page graphics (lines, boxes, logos, rotated labels) as one picture per slide with every visible piece of
    text on top as a real text box at the original position, font, size, colour, bold and italic.
``editable``
    One slide per page. Text is rebuilt as real, positioned text boxes (font,
    size, bold/italic and colour preserved), filled rectangles become shapes and
    embedded pictures are cropped from a render of the page. Pages without
    extractable text (scans) automatically fall back to ``image`` mode.
``image``
    One slide per page containing a full-slide picture of the page. Pixel-perfect
    but not editable.
"""
from __future__ import annotations

import io

from .errors import ConversionError
from .utils import (EMU_PER_PT, chars_to_segments, clean_font, color_hex, font_is_bold, font_is_italic, group_runs,
                    most_common_size, parse_pages)


def _exact(src, dst, pages, dpi):
    """Forms and other positional pages: page graphics as a picture, visible text as real text boxes on top."""
    import pdfplumber
    from pptx import Presentation
    from pptx.dml.color import RGBColor
    from pptx.util import Emu, Pt

    from .pdf_layout import render_page

    with pdfplumber.open(str(src)) as pl:
        idx = parse_pages(pages, len(pl.pages))
        W, H = float(pl.pages[idx[0]].width), float(pl.pages[idx[0]].height)
    prs = Presentation()
    prs.slide_width, prs.slide_height = Emu(int(W * EMU_PER_PT)), Emu(int(H * EMU_PER_PT))
    for i in idx:
        spans, png, (pw, ph) = render_page(src, i, dpi * 4 // 3)
        k = min(W / pw, H / ph)
        ox, oy = (W - pw * k) / 2, (H - ph * k) / 2
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        slide.shapes.add_picture(io.BytesIO(png), Emu(int(ox * EMU_PER_PT)), Emu(int(oy * EMU_PER_PT)),
                                 Emu(int(pw * k * EMU_PER_PT)), Emu(int(ph * k * EMU_PER_PT)))
        for sp in spans:
            x0, y0, x1, y1 = sp["bbox"]
            tb = slide.shapes.add_textbox(Emu(int((ox + x0 * k) * EMU_PER_PT)), Emu(int((oy + y0 * k) * EMU_PER_PT)),
                                          Emu(int(((x1 - x0) * 1.15 + 6) * k * EMU_PER_PT)), Emu(int((y1 - y0) * k * EMU_PER_PT)))
            tf = tb.text_frame
            tf.word_wrap = False
            tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
            run = tf.paragraphs[0].add_run()
            run.text = sp["text"]
            f = run.font
            f.size = Pt(max(1, sp["size"] * k))
            f.bold = bool(font_is_bold(sp["font"]))
            f.italic = bool(font_is_italic(sp["font"]))
            name = clean_font(sp["font"])
            if name:
                f.name = name
            f.color.rgb = RGBColor.from_string("%06X" % (sp["color"] & 0xFFFFFF))
    prs.save(str(dst))
    return dst


def pdf_to_pptx(src, dst, *, mode: str = "auto", pages=None, dpi: int = 150, **_):
    import pdfplumber
    import pypdfium2 as pdfium
    from pptx import Presentation
    from pptx.dml.color import RGBColor
    from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
    from pptx.util import Emu, Pt

    if mode not in ("auto", "editable", "exact", "image"):
        raise ValueError("mode must be 'auto', 'editable', 'exact' or 'image'")
    if mode in ("auto", "exact"):
        from .pdf_layout import looks_like_form
        if mode == "exact" or looks_like_form(src, pages):
            return _exact(src, dst, pages, dpi)
        mode = "editable"

    pdoc = pdfium.PdfDocument(str(src))
    with pdfplumber.open(str(src)) as pdf:
        idx = parse_pages(pages, len(pdf.pages))
        first = pdf.pages[idx[0]]
        W, H = float(first.width), float(first.height)
        prs = Presentation()
        prs.slide_width, prs.slide_height = Emu(int(W * EMU_PER_PT)), Emu(int(H * EMU_PER_PT))
        blank = prs.slide_layouts[6]
        scale_px = dpi / 72.0

        for i in idx:
            page = pdf.pages[i]
            pw, ph = float(page.width), float(page.height)
            s = min(W / pw, H / ph)
            ox, oy = (W - pw * s) / 2, (H - ph * s) / 2
            slide = prs.slides.add_slide(blank)

            def E(v):  # PDF points -> EMU
                return Emu(int(v * EMU_PER_PT))

            def render():
                return pdoc[i].render(scale=scale_px).to_pil().convert("RGB")

            def add_pil(img, x, y, w, h):
                buf = io.BytesIO()
                img.save(buf, format="PNG")
                buf.seek(0)
                slide.shapes.add_picture(buf, E(ox + x * s), E(oy + y * s), E(w * s), E(h * s))

            lines = []
            if mode == "editable":
                lines = page.extract_text_lines(return_chars=True, strip=True)

            if mode == "image" or not lines:
                add_pil(render(), 0, 0, pw, ph)
                continue

            # filled and outlined rectangles (table shading, banners, boxes, borders)
            for rect in [r for r in page.rects if r.get("fill") or r.get("stroke")][:600]:
                w, h = rect["x1"] - rect["x0"], rect["bottom"] - rect["top"]
                filled = bool(rect.get("fill"))
                hexc = color_hex(rect.get("non_stroking_color")) if filled else None
                if (w < 0.6 and h < 0.6) or (filled and hexc == "FFFFFF" and w * h > 0.8 * pw * ph and not rect.get("stroke")):
                    continue
                shp = slide.shapes.add_shape(
                    MSO_SHAPE.RECTANGLE, E(ox + rect["x0"] * s), E(oy + rect["top"] * s), E(max(w, 0.5) * s), E(max(h, 0.5) * s)
                )
                if filled:
                    shp.fill.solid()
                    shp.fill.fore_color.rgb = RGBColor.from_string(hexc)
                else:
                    shp.fill.background()
                if rect.get("stroke"):
                    shp.line.color.rgb = RGBColor.from_string(color_hex(rect.get("stroking_color")))
                    shp.line.width = Pt(max(0.25, (rect.get("linewidth") or 0.5) * s))
                else:
                    shp.line.fill.background()
                shp.shadow.inherit = False

            # ruled lines (underlines, separators, signature lines)
            for ln in page.lines[:800]:
                x0, y0, x1, y1 = ln["x0"], ln["top"], ln["x1"], ln["bottom"]
                conn = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, E(ox + x0 * s), E(oy + y0 * s), E(ox + x1 * s), E(oy + y1 * s))
                conn.line.color.rgb = RGBColor.from_string(color_hex(ln.get("stroking_color")))
                conn.line.width = Pt(max(0.25, (ln.get("linewidth") or 0.5) * s))

            # pictures
            if page.images:
                full = render()
                for im in page.images:
                    x0, y0, x1, y1 = max(im["x0"], 0), max(im["top"], 0), min(im["x1"], pw), min(im["bottom"], ph)
                    if x1 - x0 < 2 or y1 - y0 < 2:
                        continue
                    crop = full.crop((int(x0 * scale_px), int(y0 * scale_px), int(x1 * scale_px), int(y1 * scale_px)))
                    add_pil(crop, x0, y0, x1 - x0, y1 - y0)

            # text
            for line in lines:
                for seg in chars_to_segments(line["chars"]):
                    runs = group_runs(seg)
                    xs0 = min(c["x0"] for _, c in seg)
                    xs1 = max(c["x1"] for _, c in seg)
                    top = min(c["top"] for _, c in seg)
                    bottom = max(c["bottom"] for _, c in seg)
                    tb = slide.shapes.add_textbox(
                        E(ox + xs0 * s), E(oy + top * s), E((xs1 - xs0) * s * 1.06 + 4), E((bottom - top) * s)
                    )
                    tf = tb.text_frame
                    tf.word_wrap = False
                    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
                    para = tf.paragraphs[0]
                    for r in runs:
                        run = para.add_run()
                        run.text = r["text"]
                        f = run.font
                        f.size = Pt(max(1, r["size"] * s))
                        f.bold, f.italic = r["bold"], r["italic"]
                        if r["font"]:
                            f.name = r["font"]
                        f.color.rgb = RGBColor.from_string(r["color"])
    prs.save(str(dst))
    return dst
