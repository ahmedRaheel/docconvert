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
from .office_to_pdf import _esc, _fonts, _libreoffice, find_soffice


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
class _Builder(HTMLParser):
    SKIP = {"script", "style", "head", "title", "noscript", "template"}
    BLOCK = {"p", "div", "section", "article", "header", "footer", "main", "nav", "aside", "blockquote",
             "h1", "h2", "h3", "h4", "h5", "h6", "ul", "ol", "li", "pre", "hr", "figure", "figcaption", "tr"}
    HEAD = {"h1": 24, "h2": 19, "h3": 16, "h4": 14, "h5": 12, "h6": 11}

    def __init__(self, base: Path, fonts, avail):
        super().__init__(convert_charrefs=True)
        from reportlab.lib.styles import ParagraphStyle
        self.PS, self.base, self.fonts, self.avail = ParagraphStyle, base, fonts, avail
        self.flow, self.buf, self.kind = [], "", None
        self.skip, self.pre = 0, 0
        self.lists, self.bullet = [], None
        self.table = self.row = self.cell = None  # cell = list holding the cell's buffer
        self.cell_head = False
        self.quote = 0
        self.normal = ParagraphStyle("n", fontName=fonts[0], fontSize=10.5, leading=14, spaceAfter=6)

    # -- output helpers
    def _emit(self, text):
        if self.cell is not None:
            self.cell[0] += text
        else:
            self.buf += text

    def flush(self):
        from reportlab.platypus import Paragraph, Preformatted
        text, self.buf = self.buf.strip(), ""
        kind, bullet = self.kind, self.bullet
        self.kind, self.bullet = None, None
        if not text:
            return
        ps = self.PS("x", parent=self.normal)
        if kind in self.HEAD:
            ps.fontName, ps.fontSize = self.fonts[1], self.HEAD[kind]
            ps.leading, ps.spaceBefore, ps.spaceAfter = ps.fontSize * 1.25, 8, 6
        elif kind == "pre":
            ps.fontName, ps.fontSize, ps.leading = "Courier", 9, 11
            text = text.replace(" ", "&nbsp;").replace("\n", "<br/>")
        if self.lists:
            ps.leftIndent, ps.bulletIndent, ps.spaceAfter = 16 * len(self.lists), 16 * len(self.lists) - 12, 2
        if self.quote:
            ps.leftIndent += 18
            ps.fontName = self.fonts[2]
        try:
            self.flow.append(Paragraph(text, ps, bulletText=bullet))
        except Exception:  # malformed inline markup -> plain text
            self.flow.append(Paragraph(re.sub(r"<[^>]+>", "", text), ps, bulletText=bullet))

    # -- parser callbacks
    def handle_starttag(self, tag, attrs):
        a = dict(attrs)
        if tag in self.SKIP:
            self.skip += 1
            return
        if self.skip:
            return
        if tag == "table":
            self.flush()
            self.table = []
        elif tag == "tr" and self.table is not None:
            self.row = []
        elif tag in ("td", "th") and self.table is not None:
            self.cell, self.cell_head = [""], tag == "th"
        elif self.cell is not None:
            if tag in ("p", "div", "br", "li"):
                self.cell[0] += "<br/>"
            self._inline(tag, a)
        elif tag in ("ul", "ol"):
            self.flush()
            self.lists.append([tag, 0])
        elif tag == "li":
            self.flush()
            if self.lists:
                self.lists[-1][1] += 1
                self.bullet = "•" if self.lists[-1][0] == "ul" else f"{self.lists[-1][1]}."
        elif tag == "blockquote":
            self.flush()
            self.quote += 1
        elif tag == "hr":
            from reportlab.platypus import HRFlowable
            self.flush()
            self.flow.append(HRFlowable(width="100%", color="#999999", spaceBefore=4, spaceAfter=6))
        elif tag in self.BLOCK:
            self.flush()
            self.kind = tag
            if tag == "pre":
                self.pre += 1
        else:
            self._inline(tag, a)

    def _inline(self, tag, a):
        m = {"b": "<b>", "strong": "<b>", "i": "<i>", "em": "<i>", "cite": "<i>", "u": "<u>", "ins": "<u>",
             "s": "<strike>", "del": "<strike>", "strike": "<strike>", "sub": "<sub>", "sup": "<super>",
             "code": '<font face="Courier">', "kbd": '<font face="Courier">', "tt": '<font face="Courier">',
             "br": "<br/>", "mark": "<b>"}
        if tag == "a" and a.get("href"):
            self._emit(f'<a href="{_esc(a["href"]).replace(chr(34), "&quot;")}" color="#0645AD">')
        elif tag == "img":
            self._image(a)
        elif tag in m:
            self._emit(m[tag])

    def _image(self, a):
        from reportlab.lib.utils import ImageReader
        from reportlab.platypus import Image as RImage
        src, data = a.get("src", ""), None
        try:
            if src.startswith("data:") and "base64," in src:
                data = base64.b64decode(src.split("base64,", 1)[1])
            elif not re.match(r"^[a-z]+://", src):
                p = (self.base / src.split("?")[0].lstrip("/") if src.startswith("/") else self.base / src)
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
        if tag in self.SKIP:
            self.skip = max(0, self.skip - 1)
            return
        if self.skip:
            return
        if tag in ("td", "th") and self.cell is not None:
            self.row.append((self.cell[0], self.cell_head))
            self.cell = None
        elif tag == "tr" and self.table is not None and self.row is not None:
            if self.row:
                self.table.append(self.row)
            self.row = None
        elif tag == "table" and self.table is not None:
            self._table()
            self.table = None
        elif self.cell is not None:
            self._close_inline(tag)
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
        else:
            self._close_inline(tag)

    def _close_inline(self, tag):
        m = {"b": "</b>", "strong": "</b>", "i": "</i>", "em": "</i>", "cite": "</i>", "u": "</u>", "ins": "</u>",
             "s": "</strike>", "del": "</strike>", "strike": "</strike>", "sub": "</sub>", "sup": "</super>",
             "code": "</font>", "kbd": "</font>", "tt": "</font>", "a": "</a>", "mark": "</b>"}
        if tag in m:
            self._emit(m[tag])

    def handle_data(self, data):
        if self.skip:
            return
        if not self.pre:
            data = re.sub(r"\s+", " ", data)
            if not data.strip() and not (self.cell or self.buf):
                return
        self._emit(_esc(data))

    def _table(self):
        from reportlab.lib import colors
        from reportlab.platypus import Paragraph, Spacer
        from reportlab.platypus import Table as RTable, TableStyle
        if not self.table:
            return
        n = max(len(r) for r in self.table)
        cs = self.PS("c", parent=self.normal, fontSize=9, leading=11, spaceAfter=0)
        data, head_rows = [], []
        for ri, row in enumerate(self.table):
            cells = []
            for txt, head in row:
                txt = txt.strip().removeprefix("<br/>")
                try:
                    cells.append(Paragraph(f"<b>{txt}</b>" if head and txt else txt, cs))
                except Exception:
                    cells.append(Paragraph(re.sub(r"<[^>]+>", "", txt), cs))
            cells += [""] * (n - len(cells))
            data.append(cells)
            if any(h for _, h in row):
                head_rows.append(ri)
        style = [("GRID", (0, 0), (-1, -1), 0.5, colors.grey), ("VALIGN", (0, 0), (-1, -1), "TOP")]
        style += [("BACKGROUND", (0, r), (-1, r), colors.HexColor("#EEEEEE")) for r in head_rows]
        t = RTable(data, repeatRows=1 if head_rows else 0)
        t.setStyle(TableStyle(style))
        self.flow += [t, Spacer(1, 8)]


def _builtin(src: Path, dst: Path):
    from reportlab.lib.pagesizes import A4
    from reportlab.platypus import SimpleDocTemplate, Spacer

    raw = src.read_bytes()
    try:
        html = raw.decode("utf-8")
    except UnicodeDecodeError:
        html = raw.decode("latin-1")
    fonts = _fonts()
    margin = 42
    b = _Builder(src.parent, fonts, A4[0] - 2 * margin)
    b.feed(html)
    b.close()
    b.flush()
    story = b.flow or [Spacer(1, 1)]
    SimpleDocTemplate(str(dst), pagesize=A4, leftMargin=margin, rightMargin=margin,
                      topMargin=margin, bottomMargin=margin).build(story)
