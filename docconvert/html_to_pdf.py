"""HTML -> PDF.

Engines (``engine=``)
---------------------
``weasyprint``   best CSS support, pure Python install (``pip install weasyprint``)
``chrome``       headless Chrome / Chromium / Edge (full browser rendering, JS included)
``libreoffice``  LibreOffice Writer/Web
``builtin``      pure-Python reportlab renderer: headings, paragraphs, bold/italic/underline,
                 links, lists, tables, preformatted text and local/data-URI images. CSS is ignored.
``auto``         first available of the above, in that order.
"""
from __future__ import annotations

import base64
import io
import os
import re
import shutil
import subprocess
import tempfile
from html.parser import HTMLParser
from pathlib import Path

from .errors import ConversionError, MissingDependencyError
from .builtin_fonts import esc as _esc, face, families
from .office_to_pdf import _libreoffice, find_soffice


def find_chrome() -> str | None:
    env = os.environ.get("DOCCONVERT_CHROME")
    if env and Path(env).exists():
        return env
    for name in ("google-chrome", "google-chrome-stable", "chromium", "chromium-browser", "chrome", "msedge"):
        p = shutil.which(name)
        if p:
            return p
    for p in ("/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
              r"C:\Program Files\Google\Chrome\Application\chrome.exe",
              r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"):
        if Path(p).exists():
            return p
    return None


def _has_weasyprint() -> bool:
    try:
        import weasyprint  # noqa: F401
        return True
    except Exception:
        return False


def html_to_pdf(src, dst, *, engine: str = "auto", timeout: int = 120, **_):
    if engine not in ("auto", "weasyprint", "chrome", "libreoffice", "builtin"):
        engine = "auto"
    order = [engine] if engine != "auto" else ["weasyprint", "chrome", "libreoffice", "builtin"]
    last = None
    for eng in order:
        try:
            if eng == "weasyprint":
                if not _has_weasyprint():
                    raise MissingDependencyError("pip install weasyprint")
                from weasyprint import HTML
                HTML(filename=str(src)).write_pdf(str(dst))
            elif eng == "chrome":
                exe = find_chrome()
                if not exe:
                    raise MissingDependencyError("Chrome/Chromium not found (or set DOCCONVERT_CHROME)")
                _chrome(exe, Path(src), Path(dst), timeout)
            elif eng == "libreoffice":
                exe = find_soffice()
                if not exe:
                    raise MissingDependencyError("LibreOffice not found")
                _libreoffice(exe, src, dst, timeout, filter_name="pdf:writer_web_pdf_Export", copy=False)
            else:
                _builtin(Path(src), Path(dst))
            if Path(dst).exists() and Path(dst).stat().st_size > 0:
                return dst
        except Exception as exc:  # try the next engine in auto mode
            last = exc
            if engine != "auto":
                raise
    raise ConversionError(f"HTML -> PDF failed: {last}")


def _chrome(exe, src: Path, dst: Path, timeout):
    with tempfile.TemporaryDirectory(prefix="docconvert_") as td:
        cmd = [exe, "--headless", "--disable-gpu", "--no-sandbox", f"--user-data-dir={td}",
               "--no-pdf-header-footer", f"--print-to-pdf={dst.resolve()}", src.resolve().as_uri()]
        try:
            res = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout)
        except subprocess.TimeoutExpired as exc:
            raise ConversionError("Chrome timed out") from exc
    if not dst.exists():
        raise ConversionError(f"Chrome produced no PDF: {res.stderr.strip()[:300]}")


# --------------------------------------------------------------------------- builtin
_NAMED = {"black": "000000", "white": "FFFFFF", "red": "FF0000", "green": "008000", "blue": "0000FF", "yellow": "FFFF00",
          "gray": "808080", "grey": "808080", "orange": "FFA500", "purple": "800080", "navy": "000080", "teal": "008080",
          "silver": "C0C0C0", "maroon": "800000", "olive": "808000", "lime": "00FF00", "aqua": "00FFFF", "fuchsia": "FF00FF",
          "pink": "FFC0CB", "brown": "A52A2A", "gold": "FFD700", "crimson": "DC143C", "tomato": "FF6347", "coral": "FF7F50",
          "darkblue": "00008B", "darkgreen": "006400", "darkred": "8B0000", "lightgray": "D3D3D3", "lightgrey": "D3D3D3",
          "lightblue": "ADD8E6", "lightgreen": "90EE90", "whitesmoke": "F5F5F5", "indigo": "4B0082", "violet": "EE82EE",
          "steelblue": "4682B4", "royalblue": "4169E1", "dodgerblue": "1E90FF", "slategray": "708090", "dimgray": "696969"}
_FONT_SIZES = {"xx-small": 6.75, "x-small": 7.5, "small": 9.75, "medium": 12, "large": 13.5, "x-large": 18, "xx-large": 24}
_LEGACY_SIZES = [7.5, 10, 12, 13.5, 18, 24, 36]
BASE_SIZE = 11.0


def css_color(v: str | None):
    if not v:
        return None
    v = v.strip().lower()
    m = re.search(r"rgba?\(\s*(\d+)[,\s]+(\d+)[,\s]+(\d+)", v)
    if m:
        return "%02X%02X%02X" % tuple(min(255, int(x)) for x in m.groups())
    v = v.split()[0] if v else ""
    if v in _NAMED:
        return _NAMED[v]
    m = re.fullmatch(r"#([0-9a-f]{3}|[0-9a-f]{6})", v)
    if m:
        h = m.group(1)
        return (h if len(h) == 6 else "".join(c * 2 for c in h)).upper()
    m = re.fullmatch(r"rgba?\(\s*(\d+)[,\s]+(\d+)[,\s]+(\d+)", v)
    if m:
        return "%02X%02X%02X" % tuple(min(255, int(x)) for x in m.groups())
    return None


def css_length(v: str, parent: float) -> float | None:
    v = (v or "").strip().lower()
    if v in _FONT_SIZES:
        return _FONT_SIZES[v]
    if v == "smaller":
        return parent * 0.83
    if v == "larger":
        return parent * 1.2
    m = re.fullmatch(r"([\d.]+)\s*(px|pt|em|rem|%|pc|in|cm|mm)?", v)
    if not m:
        return None
    n, unit = float(m.group(1)), m.group(2) or "px"
    return {"px": n * 0.75, "pt": n, "em": n * parent, "rem": n * BASE_SIZE, "%": parent * n / 100, "pc": n * 12,
            "in": n * 72, "cm": n * 28.35, "mm": n * 2.835}[unit]


def parse_decls(text: str) -> dict:
    out = {}
    for part in re.split(r";(?![^()]*\))", text or ""):
        if ":" in part:
            k, v = part.split(":", 1)
            out[k.strip().lower()] = re.sub(r"\s*!important", "", v.strip(), flags=re.I)
    return out


_UA = {
    "h1": dict(size=24, bold=True), "h2": dict(size=19, bold=True), "h3": dict(size=16, bold=True),
    "h4": dict(size=14, bold=True), "h5": dict(size=12, bold=True), "h6": dict(size=11, bold=True),
    "b": dict(bold=True), "strong": dict(bold=True), "th": dict(bold=True, align="center"),
    "i": dict(italic=True), "em": dict(italic=True), "cite": dict(italic=True), "dfn": dict(italic=True),
    "u": dict(underline=True), "ins": dict(underline=True), "s": dict(strike=True), "del": dict(strike=True),
    "strike": dict(strike=True), "code": dict(family="monospace"), "kbd": dict(family="monospace"),
    "samp": dict(family="monospace"), "tt": dict(family="monospace"), "pre": dict(family="monospace"),
    "mark": dict(bg="FFFF00"), "center": dict(align="center"), "blockquote": dict(italic=True),
    "small": dict(scale=0.83), "big": dict(scale=1.2),
}


def apply_decls(style: dict, d: dict, parent_size: float):
    if "color" in d and css_color(d["color"]):
        style["color"] = css_color(d["color"])
    for key in ("background-color", "background"):
        if key in d:
            c = next((css_color(t) for t in d[key].split() if css_color(t)), None)
            if c:
                style["bg"] = c
    if "font-size" in d:
        sz = css_length(d["font-size"], parent_size)
        if sz:
            style["size"] = sz
    if "font-weight" in d:
        w = d["font-weight"].lower()
        style["bold"] = w in ("bold", "bolder") or (w.isdigit() and int(w) >= 600)
    if "font-style" in d:
        style["italic"] = d["font-style"].lower() in ("italic", "oblique")
    td = d.get("text-decoration") or d.get("text-decoration-line")
    if td:
        style["underline"] = "underline" in td
        style["strike"] = "line-through" in td
    if "font-family" in d:
        style["family"] = d["font-family"].split(",")[0].strip(" \"'")
    if "text-align" in d:
        style["align"] = d["text-align"].lower()
    for key, val in d.items():
        if key == "border" or re.fullmatch(r"border-(top|bottom|left|right)", key):
            if not re.search(r"\bnone\b|^0", val.strip()):
                col = next((css_color(t) for t in val.split() if css_color(t)), "000000")
                style["border"] = col


class _Selector:
    def __init__(self, text):
        m = re.fullmatch(r"([a-zA-Z][\w-]*)?((?:\.[\w-]+)*)(#[\w-]+)?", text.strip())
        self.ok = bool(m) and bool(text.strip())
        if self.ok:
            self.tag = (m.group(1) or "").lower()
            self.classes = [c for c in (m.group(2) or "").split(".") if c]
            self.id = (m.group(3) or "")[1:]
        self.spec = (1 if self.id else 0, len(getattr(self, "classes", [])), 1 if getattr(self, "tag", "") else 0)

    def matches(self, tag, attrs):
        if not self.ok or (self.tag and self.tag != tag):
            return False
        cls = (attrs.get("class") or "").split()
        return all(c in cls for c in self.classes) and (not self.id or attrs.get("id") == self.id)


class _Builder(HTMLParser):
    SKIP = {"script", "noscript", "template", "head", "title"}
    BLOCK = {"p", "div", "section", "article", "header", "footer", "main", "nav", "aside", "blockquote", "center",
             "h1", "h2", "h3", "h4", "h5", "h6", "ul", "ol", "li", "pre", "hr", "figure", "figcaption", "tr", "body", "address"}
    VOID = {"br", "img", "hr", "input", "meta", "link", "area", "base", "col", "embed", "source", "track", "wbr"}

    def __init__(self, base: Path, avail):
        super().__init__(convert_charrefs=True)
        from reportlab.lib.styles import ParagraphStyle
        self.PS, self.base, self.avail = ParagraphStyle, base, avail
        self.flow, self.buf, self.kind, self.meta = [], "", None, {}
        self.skip, self.pre, self.quote = 0, 0, 0
        self.lists, self.bullet = [], None
        self.table = self.row = self.cell = None
        self.cell_meta, self.table_meta = {}, {}
        self.rules, self.in_style, self.style_buf = [], False, ""
        self.stack = [dict(size=BASE_SIZE, color="000000")]
        self.normal = ParagraphStyle("n", fontName=face(None), fontSize=BASE_SIZE, leading=BASE_SIZE * 1.3, spaceAfter=6)

    # ---------- styles
    def _compute(self, tag, attrs):
        parent = self.stack[-1]
        st = {k: parent[k] for k in ("color", "size", "bold", "italic", "underline", "strike", "family", "align") if k in parent}
        ua = dict(_UA.get(tag, {}))
        if tag == "a" and attrs.get("href"):
            ua.update(color="0645AD", underline=True)
        scale = ua.pop("scale", None)
        st.update({k: v for k, v in ua.items()})
        if scale:
            st["size"] = parent["size"] * scale
        if tag in ("h1", "h2", "h3", "h4", "h5", "h6") and "size" in ua:
            st["size"] = ua["size"] * BASE_SIZE / 11.0 if False else ua["size"]
        matched = sorted((r for r in self.rules if r[0].matches(tag, attrs)), key=lambda r: (r[0].spec, r[2]))
        for sel, decls, _ in matched:
            apply_decls(st, decls, parent["size"])
        # legacy attributes
        if attrs.get("align"):
            st["align"] = attrs["align"].lower()
        if tag == "font":
            if css_color(attrs.get("color")):
                st["color"] = css_color(attrs["color"])
            if attrs.get("size", "").strip("+-").isdigit():
                n = int(attrs["size"].strip("+-"))
                n = n if not attrs["size"].startswith(("+", "-")) else 3 + (n if attrs["size"][0] == "+" else -n)
                st["size"] = _LEGACY_SIZES[max(1, min(7, n)) - 1]
            if attrs.get("face"):
                st["family"] = attrs["face"].split(",")[0]
        if attrs.get("bgcolor") and css_color(attrs["bgcolor"]):
            st["bg"] = css_color(attrs["bgcolor"])
        if attrs.get("style"):
            apply_decls(st, parse_decls(attrs["style"]), parent["size"])
        if "bg" in ua:
            st["bg"] = ua["bg"]
        return st

    def _wrap(self, text):
        st = self.stack[-1]
        size = st.get("size", BASE_SIZE)
        attrs = f'name="{face(st.get("family"), bool(st.get("bold")), bool(st.get("italic")))}" size="{size:.1f}" color="#{st.get("color", "000000")}"'
        inline_bg = st.get("bg") if st.get("_inline") else None
        if inline_bg:
            attrs += f' backcolor="#{inline_bg}"'
        out = f"<font {attrs}>{text}</font>"
        if st.get("underline"):
            out = f"<u>{out}</u>"
        if st.get("strike"):
            out = f"<strike>{out}</strike>"
        return out

    # ---------- output helpers
    def _emit(self, text):
        if self.cell is not None:
            self.cell[0] += text
        else:
            self.buf += text

    def flush(self):
        from reportlab.lib import colors
        from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT, TA_RIGHT
        from reportlab.platypus import Paragraph
        text, self.buf = self.buf.strip(), ""
        kind, bullet, meta = self.kind, self.bullet, self.meta
        self.kind, self.bullet, self.meta = None, None, {}
        if not text.replace("&nbsp;", "").strip() or not re.sub(r"<[^>]+>", "", text).strip():
            return
        size = meta.get("size", BASE_SIZE)
        ps = self.PS("x", parent=self.normal, fontSize=size, leading=size * 1.3,
                     alignment={"center": TA_CENTER, "right": TA_RIGHT, "justify": TA_JUSTIFY}.get(meta.get("align"), TA_LEFT))
        if kind in ("h1", "h2", "h3", "h4", "h5", "h6"):
            ps.spaceBefore, ps.spaceAfter = size * 0.6, size * 0.4
        elif kind == "pre":
            text = text.replace(" ", "&nbsp;").replace("\n", "<br/>")
        if meta.get("bg"):
            ps.backColor = colors.HexColor("#" + meta["bg"])
        if self.lists:
            ps.leftIndent, ps.bulletIndent, ps.spaceAfter = 16 * len(self.lists), 16 * len(self.lists) - 12, 2
            ps.bulletFontName, ps.bulletFontSize = face(None), size
        if self.quote:
            ps.leftIndent += 18
        try:
            self.flow.append(Paragraph(text, ps, bulletText=bullet))
        except Exception:
            self.flow.append(Paragraph(re.sub(r"<[^>]+>", "", text), ps, bulletText=bullet))

    # ---------- parser callbacks
    def handle_starttag(self, tag, attrs):
        a = dict(attrs)
        if tag == "style":
            self.in_style, self.style_buf = True, ""
            return
        if tag in self.SKIP:
            self.skip += 1
            return
        if self.skip:
            return
        st = self._compute(tag, a)
        inline = tag not in self.BLOCK and tag not in ("table", "td", "th", "tr", "ul", "ol", "li")
        st["_inline"] = inline
        if tag not in self.VOID:
            self.stack.append(st)
        if tag == "table":
            self.flush()
            border = a.get("border", "0") not in ("0", "", None) or bool(st.get("border"))
            self.table, self.table_meta = [], dict(border=st.get("border") or ("000000" if border else None),
                                                   full=str(a.get("width", "")).endswith("%") or "100%" in a.get("style", ""))
        elif tag == "tr" and self.table is not None:
            self.row = []
        elif tag in ("td", "th") and self.table is not None:
            self.cell = [""]
            self.cell_meta = dict(head=tag == "th", bg=st.get("bg"), align=st.get("align"), size=st.get("size", BASE_SIZE),
                                  border=st.get("border"), span=int(a.get("colspan", 1) or 1) if str(a.get("colspan", "1")).isdigit() else 1)
        elif self.cell is not None:
            if tag in ("p", "div", "br", "li"):
                self.cell[0] += "<br/>"
            self._inline_tag(tag, a)
        elif tag in ("ul", "ol"):
            self.flush()
            self.lists.append([tag, 0])
        elif tag == "li":
            self.flush()
            if self.lists:
                self.lists[-1][1] += 1
                self.bullet = "•" if self.lists[-1][0] == "ul" else f"{self.lists[-1][1]}."
            self.kind, self.meta = "li", dict(align=st.get("align"), bg=st.get("bg"), size=st.get("size", BASE_SIZE))
        elif tag == "blockquote":
            self.flush()
            self.quote += 1
            self.kind, self.meta = tag, dict(align=st.get("align"), bg=st.get("bg"), size=st.get("size", BASE_SIZE))
        elif tag == "hr":
            from reportlab.platypus import HRFlowable
            self.flush()
            self.flow.append(HRFlowable(width="100%", color="#999999", spaceBefore=4, spaceAfter=6))
        elif tag in self.BLOCK:
            self.flush()
            self.kind, self.meta = tag, dict(align=st.get("align"), bg=st.get("bg"), size=st.get("size", BASE_SIZE))
            if tag == "pre":
                self.pre += 1
        else:
            self._inline_tag(tag, a)

    def _inline_tag(self, tag, a):
        if tag == "br":
            self._emit("<br/>")
        elif tag == "a" and a.get("href"):
            self._emit(f'<a href="{_esc(a["href"]).replace(chr(34), "&quot;")}">')
        elif tag == "img":
            self._image(a)
        elif tag == "sub":
            self._emit("<sub>")
        elif tag == "sup":
            self._emit("<super>")

    def _image(self, a):
        from reportlab.lib.utils import ImageReader
        from reportlab.platypus import Image as RImage
        src, data = a.get("src", ""), None
        try:
            if src.startswith("data:") and "base64," in src:
                data = base64.b64decode(src.split("base64,", 1)[1])
            elif not re.match(r"^[a-z]+://", src):
                p = self.base / src.split("?")[0].lstrip("/")
                data = p.read_bytes() if p.exists() else None
            if data is None:
                raise ValueError
            iw, ih = ImageReader(io.BytesIO(data)).getSize()
            w = float(re.sub(r"[^\d.]", "", a.get("width", "")) or iw * 0.75)
            w = min(w, self.avail)
            if self.cell is None:
                self.flush()
                self.flow.append(RImage(io.BytesIO(data), width=w, height=w * ih / iw))
        except Exception:
            if a.get("alt"):
                self._emit(f"[{_esc(a['alt'])}]")

    def handle_endtag(self, tag):
        if tag == "style":
            self.in_style = False
            self._parse_css(self.style_buf)
            return
        if tag in self.SKIP:
            self.skip = max(0, self.skip - 1)
            return
        if self.skip:
            return
        if tag not in self.VOID and len(self.stack) > 1:
            for i in range(len(self.stack) - 1, 0, -1):         # tolerate unclosed tags
                self.stack.pop()
                if i == len(self.stack) or True:
                    break
        if tag in ("td", "th") and self.cell is not None:
            self.row.append((self.cell[0], self.cell_meta))
            self.cell = None
        elif tag == "tr" and self.table is not None and self.row is not None:
            if self.row:
                self.table.append(self.row)
            self.row = None
        elif tag == "table" and self.table is not None:
            self._table()
            self.table = None
        elif self.cell is not None:
            if tag == "a":
                self.cell[0] += "</a>"
            elif tag in ("sub", "sup"):
                self.cell[0] += "</sub>" if tag == "sub" else "</super>"
        elif tag in ("ul", "ol"):
            self.flush()
            if self.lists:
                self.lists.pop()
        elif tag == "blockquote":
            self.flush()
            self.quote = max(0, self.quote - 1)
        elif tag == "pre":
            self.flush()
            self.pre = max(0, self.pre - 1)
        elif tag in self.BLOCK:
            self.flush()
        elif tag == "a":
            self._emit("</a>")
        elif tag in ("sub", "sup"):
            self._emit("</sub>" if tag == "sub" else "</super>")

    def handle_data(self, data):
        if self.in_style:
            self.style_buf += data
            return
        if self.skip:
            return
        if not self.pre:
            data = re.sub(r"\s+", " ", data)
            if not data.strip() and not (self.cell and self.cell[0]) and not self.buf:
                return
        if data:
            self._emit(self._wrap(_esc(data)))

    def _parse_css(self, css):
        css = re.sub(r"/\*.*?\*/", "", css, flags=re.S)
        for m in re.finditer(r"([^{}@]+)\{([^{}]*)\}", css):
            decls = parse_decls(m.group(2))
            for sel in m.group(1).split(","):
                s = _Selector(sel)
                if s.ok:
                    self.rules.append((s, decls, len(self.rules)))

    def _table(self):
        from reportlab.lib import colors
        from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
        from reportlab.pdfbase.pdfmetrics import stringWidth
        from reportlab.platypus import Paragraph, Spacer
        from reportlab.platypus import Table as RTable, TableStyle
        if not self.table:
            return
        n = max(sum(m["span"] for _, m in r) for r in self.table)
        data, cmds, natural = [], [], [10.0] * n
        border = self.table_meta.get("border")
        for ri, row in enumerate(self.table):
            cells, col = [], 0
            for txt, m in row:
                txt = txt.strip()
                while txt.startswith("<br/>"):
                    txt = txt[5:]
                size = max(m.get("size") or BASE_SIZE, 6)
                ps = self.PS("c", parent=self.normal, fontSize=size, leading=size * 1.25, spaceAfter=0,
                             alignment={"center": TA_CENTER, "right": TA_RIGHT}.get(m.get("align"), TA_LEFT))
                try:
                    cells.append(Paragraph(txt, ps))
                except Exception:
                    cells.append(Paragraph(re.sub(r"<[^>]+>", "", txt), ps))
                plain = re.sub(r"<[^>]+>", "", txt.replace("<br/>", "\n"))
                w = max((stringWidth(line, face(None, m["head"]), size) for line in plain.split("\n")), default=0) + 12
                if m["span"] == 1 and col < n:
                    natural[col] = max(natural[col], min(w, 260))
                if m.get("bg"):
                    cmds.append(("BACKGROUND", (col, ri), (col + m["span"] - 1, ri), colors.HexColor("#" + m["bg"])))
                elif m["head"] and not border:
                    pass
                if m.get("border") and not border:
                    cmds.append(("BOX", (col, ri), (col + m["span"] - 1, ri), 0.5, colors.HexColor("#" + m["border"])))
                if m["span"] > 1:
                    cmds.append(("SPAN", (col, ri), (col + m["span"] - 1, ri)))
                cells += [""] * (m["span"] - 1)
                col += m["span"]
            cells += [""] * (n - len(cells))
            data.append(cells)
        total = sum(natural)
        if total > self.avail or self.table_meta.get("full"):
            natural = [w * self.avail / total for w in natural]
        t = RTable(data, colWidths=natural, repeatRows=0, hAlign="LEFT")
        style = [("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 5), ("RIGHTPADDING", (0, 0), (-1, -1), 5),
                 ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3)]
        if border:
            style.append(("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#" + border)))
        t.setStyle(TableStyle(style + cmds))
        self.flow += [t, Spacer(1, 8)]


def _builtin(src: Path, dst: Path):
    from reportlab.lib.pagesizes import A4
    from reportlab.platypus import SimpleDocTemplate, Spacer

    raw = src.read_bytes()
    try:
        html = raw.decode("utf-8")
    except UnicodeDecodeError:
        html = raw.decode("latin-1")
    families()
    margin = 42
    b = _Builder(src.parent, A4[0] - 2 * margin)
    b.feed(html)
    b.close()
    b.flush()
    story = b.flow or [Spacer(1, 1)]
    SimpleDocTemplate(str(dst), pagesize=A4, leftMargin=margin, rightMargin=margin,
                      topMargin=margin, bottomMargin=margin).build(story)
