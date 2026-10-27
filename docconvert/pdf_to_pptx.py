"""PDF -> PowerPoint.

Modes
-----
``editable`` (default)
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
from .utils import EMU_PER_PT, chars_to_segments, color_hex, group_runs, most_common_size, parse_pages


def pdf_to_pptx(src, dst, *, mode: str = "editable", pages=None, dpi: int = 150, **_):
    import pdfplumber
    import pypdfium2 as pdfium
    from pptx import Presentation
    from pptx.dml.color import RGBColor
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.util import Emu, Pt

    if mode not in ("editable", "image"):
        raise ValueError("mode must be 'editable' or 'image'")

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

            # background rectangles (table shading, banners, ...)
            for rect in [r for r in page.rects if r.get("fill")][:400]:
                hexc = color_hex(rect.get("non_stroking_color"))
                w, h = rect["x1"] - rect["x0"], rect["bottom"] - rect["top"]
                if w < 1 or h < 1 or (hexc == "FFFFFF" and w * h > 0.8 * pw * ph):
                    continue
                shp = slide.shapes.add_shape(
                    MSO_SHAPE.RECTANGLE, E(ox + rect["x0"] * s), E(oy + rect["top"] * s), E(w * s), E(h * s)
                )
                shp.fill.solid()
                shp.fill.fore_color.rgb = RGBColor.from_string(hexc)
                if rect.get("stroke"):
                    shp.line.color.rgb = RGBColor.from_string(color_hex(rect.get("stroking_color")))
                    shp.line.width = Pt(max(0.25, rect.get("linewidth", 0.5) * s))
                else:
                    shp.line.fill.background()
                shp.shadow.inherit = False

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
