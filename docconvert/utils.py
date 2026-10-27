"""Small helpers shared by the converters."""
from __future__ import annotations

import re
from collections import Counter

EMU_PER_PT = 12700


def parse_pages(spec, n: int) -> list[int]:
    """Turn ``None`` / ``3`` / ``[1, 2]`` / ``"1-3,5,8-"`` (1-based) into 0-based indices."""
    if spec is None:
        return list(range(n))
    if isinstance(spec, int):
        spec = [spec]
    if isinstance(spec, str):
        out: list[int] = []
        for part in spec.split(","):
            part = part.strip()
            if not part:
                continue
            if "-" in part:
                a, b = part.split("-", 1)
                out.extend(range(int(a) if a else 1, (int(b) if b else n) + 1))
            else:
                out.append(int(part))
        spec = out
    idx = sorted({p - 1 for p in spec if 1 <= p <= n})
    if not idx:
        raise ValueError(f"No valid pages selected (document has {n} pages)")
    return idx


_NUM = re.compile(r"(\()?(-)?([$€£])?(\d{1,3}(?:,\d{3})+|\d+)(\.\d+)?(\))?(%)?")


def coerce_cell(value):
    """Convert text such as ``"1,234.50"`` or ``"12%"`` to ``(number, number_format)``.

    Anything that is not clearly a number (including IDs with leading zeros) is
    returned unchanged as text.
    """
    if value is None:
        return None, None
    s = str(value).strip()
    if not s:
        return None, None
    if re.fullmatch(r"0\d+", s):
        return s, None
    m = _NUM.fullmatch(s)
    if not m:
        return s, None
    paren, neg, _cur, intp, frac, rparen, pct = m.groups()
    if bool(paren) != bool(rparen):
        return s, None
    num = float(intp.replace(",", "") + (frac or ""))
    if neg or paren:
        num = -num
    fmt = None
    if pct:
        num /= 100
        fmt = "0.00%" if frac else "0%"
    elif "," in intp:
        fmt = "#,##0.00" if frac else "#,##0"
    if not frac and not pct:
        num = int(num)
    return num, fmt


def color_hex(c) -> str:
    """PDF colour (gray / RGB / CMYK tuple) -> ``RRGGBB``."""
    try:
        if c is None:
            return "000000"
        if isinstance(c, (int, float)):
            c = (c,)
        c = tuple(float(v) for v in c)
        if len(c) == 1:
            r = g = b = c[0]
        elif len(c) == 3:
            r, g, b = c
        elif len(c) == 4:
            C, M, Y, K = c
            r, g, b = (1 - C) * (1 - K), (1 - M) * (1 - K), (1 - Y) * (1 - K)
        else:
            return "000000"
        return "%02X%02X%02X" % tuple(max(0, min(255, round(v * 255))) for v in (r, g, b))
    except Exception:
        return "000000"


_FONT_MAP = {
    "helvetica": "Arial", "arial": "Arial", "times": "Times New Roman",
    "courier": "Courier New", "calibri": "Calibri", "cambria": "Cambria",
    "verdana": "Verdana", "georgia": "Georgia", "tahoma": "Tahoma",
    "symbol": "Symbol", "liberation": "Arial", "dejavu": "DejaVu Sans",
}


def clean_font(name: str | None) -> str | None:
    """``ABCDEF+Arial-BoldMT`` -> ``Arial`` (best effort)."""
    if not name:
        return None
    name = re.sub(r"^[A-Z]{6}\+", "", name)
    base = re.split(r"[-,]", name)[0]
    low = base.lower()
    for key, val in _FONT_MAP.items():
        if low.startswith(key):
            return val
    return re.sub(r"(?<=[a-z])(?=[A-Z])", " ", base) or None


def font_is_bold(name: str | None) -> bool:
    return bool(name) and bool(re.search(r"bold|black|heavy|semibold|demi", name, re.I))


def font_is_italic(name: str | None) -> bool:
    return bool(name) and bool(re.search(r"italic|oblique", name, re.I))


def most_common_size(sizes) -> float:
    c = Counter(round(s * 2) / 2 for s in sizes if s)
    return c.most_common(1)[0][0] if c else 11.0


def chars_to_segments(chars, gap_factor: float = 2.0):
    """Split a line's characters into visually separate segments (columns).

    Returns a list of segments; each segment is a list of ``(text, char_dict)``
    pairs with inter-word spaces inserted where the PDF has gaps but no space glyph.
    """
    chars = sorted(chars, key=lambda c: c["x0"])
    segments: list[list[tuple[str, dict]]] = []
    cur: list[tuple[str, dict]] = []
    prev = None
    for ch in chars:
        text = ch.get("text", "")
        if not text:
            continue
        if prev is not None:
            size = max(ch.get("size", 10), 1)
            gap = ch["x0"] - prev["x1"]
            if gap > size * gap_factor:
                if cur:
                    segments.append(cur)
                cur = []
            elif gap > size * 0.18 and text != " " and prev.get("text") != " ":
                cur.append((" ", prev))
        if text == " " and not cur:
            prev = ch
            continue
        cur.append((text, ch))
        prev = ch
    if cur:
        segments.append(cur)
    # strip trailing spaces
    for seg in segments:
        while seg and seg[-1][0] == " ":
            seg.pop()
    return [s for s in segments if s]


def group_runs(segment):
    """Group ``(text, char)`` pairs into runs with identical styling."""
    runs: list[dict] = []
    for text, ch in segment:
        key = (
            clean_font(ch.get("fontname")),
            font_is_bold(ch.get("fontname")),
            font_is_italic(ch.get("fontname")),
            round(ch.get("size", 11) * 2) / 2,
            color_hex(ch.get("non_stroking_color")),
        )
        if runs and runs[-1]["key"] == key:
            runs[-1]["text"] += text
        else:
            runs.append({"key": key, "text": text})
    out = []
    for r in runs:
        font, bold, italic, size, color = r["key"]
        out.append(dict(text=r["text"], font=font, bold=bold, italic=italic, size=size, color=color))
    return out
