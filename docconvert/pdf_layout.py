"""Page analysis shared by the exact-copy modes (PDF -> Word and PDF -> PowerPoint).

Pure pdfplumber + pypdfium2 (no PyMuPDF): finds the text that is really visible on a page (text hidden under
later text or white boxes is dropped, letters typed into comb boxes become one item each), and renders the page
*without* that text so lines, boxes, logos, check marks and rotated labels can ride along as a picture.
"""
from __future__ import annotations

import io
import re
from statistics import median

from .utils import parse_pages

_SYMBOLIC = re.compile(r"dingbat|symbol|wingding|webding|zapf", re.I)


def _rgb_int(c) -> int:
    try:
        if isinstance(c, (int, float)):
            c = (c, c, c)
        c = tuple(c)
        if len(c) == 1:
            c = c * 3
        if len(c) == 4:                                    # CMYK
            C, M, Y, K = c
            c = ((1 - C) * (1 - K), (1 - M) * (1 - K), (1 - Y) * (1 - K))
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


def looks_like_form(src, pages=None) -> bool:
    """Heuristic used by ``layout/mode="auto"``: fillable fields, comb boxes or heavy line art."""
    import pdfplumber
    with pdfplumber.open(str(src)) as pdf:
        for i in parse_pages(pages, len(pdf.pages))[:3]:
            page = pdf.pages[i]
            try:
                if any(str((a.get("data") or {}).get("Subtype")).endswith("Widget") for a in page.annots):
                    return True
            except Exception:
                pass
            if len(page.rects) + len(page.lines) + len(page.curves) >= 400:
                return True
            singles = sum(1 for it in page_items(page) if len(it["text"]) == 1)
            if singles >= 25:
                return True
    return False


def _runs(page):
    H = float(page.height)
    runs, cur = [], None
    for seq, ch in enumerate(page.chars):
        if not ch.get("upright", True) or _SYMBOLIC.search(ch["fontname"]):
            cur = None
            continue
        m = ch.get("matrix")
        oy = H - m[5] if m else ch["bottom"] - 0.2 * ch["size"]
        c = dict(c=ch["text"], bbox=(ch["x0"], ch["top"], ch["x1"], ch["bottom"]), oy=oy)
        key = (ch["fontname"], round(ch["size"] * 2), _rgb_int(ch.get("non_stroking_color")))
        if cur and cur["key"] == key:
            cur["chars"].append(c)
            cur["seq"] = seq
        else:
            cur = dict(key=key, font=ch["fontname"], size=ch["size"], color=key[2], flags=0, chars=[c], seq=seq)
            runs.append(cur)
    return runs


def page_items(page):
    """Visible-looking text of a pdfplumber page as items: text, bbox, font, size, color, flags."""
    items = []
    for run in _runs(page):
        size = run["size"]
        pieces, cur = [], []
        for c in run["chars"]:                              # split far-apart words and separate lines
            if cur:
                pv = cur[-1]
                gap = c["bbox"][0] - pv["bbox"][2]
                if abs(c["oy"] - pv["oy"]) > 0.4 * size or gap > 1.5 * size or gap < -2 * size:
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
                continue                                    # dashes and lone symbols stay in the picture
            oy = pc[0]["oy"]                                 # baseline, so placement does not depend on font metrics
            bb = (min(c["bbox"][0] for c in pc), oy - 0.905 * size,
                  max(c["bbox"][2] for c in pc), oy + 0.212 * size)
            items.append(dict(chars=pc, text=text, size=size, font=run["font"], flags=0, color=run["color"],
                              y=round(pc[0]["oy"], 1), bbox=bb, seq=run["seq"]))
    kept = []
    for it in sorted(items, key=lambda s: -s["seq"]):       # newest first: older overlapped text is hidden
        if not any(_overlap(it["bbox"], k["bbox"]) > 0.15 for k in kept):
            kept.append(it)
    items = sorted(kept, key=lambda s: (s["y"], s["bbox"][0]))
    merged = []
    for it in items:                                        # join neighbouring pieces of the same style
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
    out = []
    for s in merged:
        ch = s["chars"]
        gaps = [ch[k + 1]["bbox"][0] - ch[k]["bbox"][2] for k in range(len(ch) - 1) if ch[k]["c"] != " " and ch[k + 1]["c"] != " "]
        if len(ch) >= 3 and gaps and median(gaps) > 0.35 * s["size"]:      # comb boxes: one item per letter
            for c in ch:
                if c["c"].strip():
                    out.append(dict(s, text=c["c"], bbox=(c["bbox"][0], s["bbox"][1], c["bbox"][2], s["bbox"][3])))
        else:
            out.append(s)
    return out


def _bounds(obj):
    return obj.get_bounds() if hasattr(obj, "get_bounds") else obj.get_pos()


def _is_rotated(obj) -> bool:
    try:
        m = obj.get_matrix()
        a, b, c, d, e, f = m.get() if hasattr(m, "get") else tuple(m)
        return abs(b) > 0.01 or abs(c) > 0.01
    except Exception:
        return False


def render_page(src, index: int, dpi: int = 200):
    """-> (visible text items, PNG of the page without that text, (width_pt, height_pt))."""
    import pdfplumber
    import pypdfium2 as pdfium
    import pypdfium2.raw as raw
    from PIL import ImageChops

    with pdfplumber.open(str(src)) as pl:
        pp = pl.pages[index]
        spans = page_items(pp)
        size = (float(pp.width), float(pp.height))
    doc = pdfium.PdfDocument(str(src))
    try:
        pg = doc[index]
        H = pg.get_height()
        k = 144 / 72
        before = pg.render(scale=k).to_pil().convert("RGB")
        boxes = [s["bbox"] for s in spans]
        objs = list(pg.get_objects(filter=[raw.FPDF_PAGEOBJ_TEXT]))
        removed = []
        for obj in objs:
            if _is_rotated(obj):
                continue
            l, b, r, t = _bounds(obj)
            ox0, oy0, ox1, oy1 = l, H - t, r, H - b               # top-based box of the text object
            area = max((ox1 - ox0) * (oy1 - oy0), 1e-6)
            covered = 0.0
            for x0, y0, x1, y1 in boxes:                           # a run may span several words/items
                w, h = min(ox1, x1 + 1) - max(ox0, x0 - 1), min(oy1, y1 + 1) - max(oy0, y0 - 1)
                if w > 0 and h > 0:
                    covered += w * h
            if covered / area >= 0.5:
                pg.remove_obj(obj)
                removed.append(obj)
        pg.gen_content()
        for obj in objs:                                     # release children before their page/document
            try:
                obj.close()
            except Exception:
                pass
        del objs, removed
        after = pg.render(scale=k).to_pil().convert("RGB")
        keep = []
        for s in spans:                                      # drop text that is hidden under white boxes etc.
            x0, y0, x1, y1 = s["bbox"]
            ix, iy = min(1.5, 0.2 * (x1 - x0)), min(1.0, 0.2 * (y1 - y0))
            box = (max(int((x0 + ix) * k), 0), max(int((y0 + iy) * k), 0),
                   min(int((x1 - ix) * k) + 1, before.width), min(int((y1 - iy) * k) + 1, before.height))
            if box[2] <= box[0] or box[3] <= box[1]:
                continue
            if ImageChops.difference(before.crop(box), after.crop(box)).convert("L").getextrema()[1] >= 40:
                keep.append(s)
        buf = io.BytesIO()
        pg.render(scale=dpi / 72).to_pil().convert("RGB").save(buf, format="PNG")
        result = (keep, buf.getvalue(), size)
        pg.close()
        return result
    finally:
        doc.close()
