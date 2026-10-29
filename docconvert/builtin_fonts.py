"""Fonts and colours shared by the pure-Python renderers.

Word/Excel/PowerPoint name fonts we usually do not have (Calibri, Arial, Times New Roman...).
We map each requested name to the closest installed *family* (sans / serif / mono) and register the
regular/bold/italic/bold-italic faces with reportlab, so bold, italic and size survive even when the
exact typeface cannot.
"""
from __future__ import annotations

import re
from pathlib import Path

_SPEC = {
    "sans": [
        ("/usr/share/fonts/truetype/liberation/", ["LiberationSans-Regular.ttf", "LiberationSans-Bold.ttf", "LiberationSans-Italic.ttf", "LiberationSans-BoldItalic.ttf"]),
        ("/usr/share/fonts/truetype/dejavu/", ["DejaVuSans.ttf", "DejaVuSans-Bold.ttf", "DejaVuSans-Oblique.ttf", "DejaVuSans-BoldOblique.ttf"]),
        (r"C:\Windows\Fonts\\", ["arial.ttf", "arialbd.ttf", "ariali.ttf", "arialbi.ttf"]),
        ("/Library/Fonts/", ["Arial.ttf", "Arial Bold.ttf", "Arial Italic.ttf", "Arial Bold Italic.ttf"]),
    ],
    "serif": [
        ("/usr/share/fonts/truetype/liberation/", ["LiberationSerif-Regular.ttf", "LiberationSerif-Bold.ttf", "LiberationSerif-Italic.ttf", "LiberationSerif-BoldItalic.ttf"]),
        ("/usr/share/fonts/truetype/dejavu/", ["DejaVuSerif.ttf", "DejaVuSerif-Bold.ttf", "DejaVuSerif-Italic.ttf", "DejaVuSerif-BoldItalic.ttf"]),
        (r"C:\Windows\Fonts\\", ["times.ttf", "timesbd.ttf", "timesi.ttf", "timesbi.ttf"]),
        ("/Library/Fonts/", ["Times New Roman.ttf", "Times New Roman Bold.ttf", "Times New Roman Italic.ttf", "Times New Roman Bold Italic.ttf"]),
    ],
    "mono": [
        ("/usr/share/fonts/truetype/liberation/", ["LiberationMono-Regular.ttf", "LiberationMono-Bold.ttf", "LiberationMono-Italic.ttf", "LiberationMono-BoldItalic.ttf"]),
        ("/usr/share/fonts/truetype/dejavu/", ["DejaVuSansMono.ttf", "DejaVuSansMono-Bold.ttf", "DejaVuSansMono-Oblique.ttf", "DejaVuSansMono-BoldOblique.ttf"]),
        (r"C:\Windows\Fonts\\", ["cour.ttf", "courbd.ttf", "couri.ttf", "courbi.ttf"]),
        ("/Library/Fonts/", ["Courier New.ttf", "Courier New Bold.ttf", "Courier New Italic.ttf", "Courier New Bold Italic.ttf"]),
    ],
}
_BASE14 = {
    "sans": ("Helvetica", "Helvetica-Bold", "Helvetica-Oblique", "Helvetica-BoldOblique"),
    "serif": ("Times-Roman", "Times-Bold", "Times-Italic", "Times-BoldItalic"),
    "mono": ("Courier", "Courier-Bold", "Courier-Oblique", "Courier-BoldOblique"),
}
_cache: dict | None = None


def families() -> dict:
    """``{"sans": (family, regular, bold, italic, bolditalic), ...}`` registered with reportlab."""
    global _cache
    if _cache:
        return _cache
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont

    out = {}
    for fam, options in _SPEC.items():
        faces = None
        for folder, files in options:
            paths = [Path(folder) / f for f in files]
            if all(p.exists() for p in paths):
                names = [f"PW-{fam}-{i}" for i in range(4)]
                for n, p in zip(names, paths):
                    if n not in pdfmetrics.getRegisteredFontNames():
                        pdfmetrics.registerFont(TTFont(n, str(p)))
                faces = names
                break
        if faces is None:
            faces = list(_BASE14[fam])
        family = f"PW-{fam}"
        pdfmetrics.registerFontFamily(family, normal=faces[0], bold=faces[1], italic=faces[2], boldItalic=faces[3])
        out[fam] = (family, *faces)
    _cache = out
    return out


_SERIF = re.compile(r"times|georgia|cambria|garamond|palatino|book ?antiqua|minion|serif(?!.*sans)|baskerville|bookman|century|constantia", re.I)
_MONO = re.compile(r"courier|consolas|mono|menlo|monaco|lucida console|source code|fixedsys|typewriter", re.I)


def family_kind(name: str | None) -> str:
    if not name:
        return "sans"
    if _MONO.search(name):
        return "mono"
    if _SERIF.search(name) and not re.search(r"sans", name, re.I):
        return "serif"
    return "sans"


def face(name: str | None, bold: bool = False, italic: bool = False) -> str:
    fam = families()[family_kind(name)]
    return fam[1 + (1 if bold else 0) + (2 if italic else 0)]


# Office default theme (2013+) used when a colour is given as a theme slot
THEME = {"dk1": "000000", "lt1": "FFFFFF", "dk2": "44546A", "lt2": "E7E6E6", "accent1": "4472C4", "accent2": "ED7D31",
         "accent3": "A5A5A5", "accent4": "FFC000", "accent5": "5B9BD5", "accent6": "70AD47", "hlink": "0563C1",
         "folHlink": "954F72", "tx1": "000000", "bg1": "FFFFFF", "tx2": "44546A", "bg2": "E7E6E6"}
THEME_INDEX = ["lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"]


def tint(hexcolor: str, amount: float) -> str:
    """Excel/Office tint: >0 lightens toward white, <0 darkens toward black."""
    r, g, b = (int(hexcolor[i:i + 2], 16) for i in (0, 2, 4))
    f = (lambda c: c + (255 - c) * amount) if amount > 0 else (lambda c: c * (1 + amount))
    return "%02X%02X%02X" % tuple(max(0, min(255, round(f(c)))) for c in (r, g, b))


def esc(s: str) -> str:
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
