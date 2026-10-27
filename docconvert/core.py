"""Format detection, routing and the public ``convert`` API."""
from __future__ import annotations

import shutil
import tempfile
from pathlib import Path
from typing import Iterable

from .errors import ConversionError, UnsupportedConversionError

OFFICE = {"docx", "xlsx", "pptx"}
LEGACY = {"doc": "docx", "xls": "xlsx", "ppt": "pptx", "odt": "docx", "ods": "xlsx", "odp": "pptx",
          "rtf": "docx", "xlsm": "xlsx"}
TARGETS = {"pdf", "docx", "xlsx", "pptx"}


def detect_format(path) -> str:
    ext = Path(path).suffix.lower().lstrip(".")
    if ext == "htm":
        ext = "html"
    if ext in ("pdf", "html") or ext in OFFICE or ext in LEGACY:
        return ext
    raise UnsupportedConversionError(f"Unsupported input format: .{ext or '?'}")


def supported_conversions() -> list[tuple[str, str]]:
    srcs = ["pdf", "docx", "xlsx", "pptx"]
    return [(s, t) for s in srcs for t in srcs if s != t] + [("html", t) for t in srcs]


def available_engines() -> dict:
    """Report which optional engines are usable on this machine."""
    from .office_to_pdf import find_soffice

    try:
        import pdf2docx  # noqa: F401
        p2d = True
    except Exception:
        p2d = False
    from .html_to_pdf import _has_weasyprint, find_chrome

    return {"pdf2docx": p2d, "libreoffice": find_soffice(), "weasyprint": _has_weasyprint(),
            "chrome": find_chrome(), "builtin": True}


def _pdf_to(target: str):
    if target == "docx":
        from .pdf_to_docx import pdf_to_docx as fn
    elif target == "xlsx":
        from .pdf_to_xlsx import pdf_to_xlsx as fn
    else:
        from .pdf_to_pptx import pdf_to_pptx as fn
    return fn


def convert(src, dst=None, *, to: str | None = None, engine: str = "auto", overwrite: bool = True, **options) -> Path:
    """Convert ``src`` to another format and return the output path.

    Any of PDF, Word, Excel, PowerPoint can be converted to any other of them, and HTML
    can be converted to PDF (or, via PDF, to Word/Excel/PowerPoint). Conversions that
    have no direct path go through PDF.

    Parameters
    ----------
    src, dst : paths. ``dst`` may be omitted if ``to`` is given (output goes next to ``src``).
    to       : target format ("pdf", "docx", "xlsx", "pptx"); inferred from ``dst`` if omitted.
    engine   : "auto" | "libreoffice" | "pdf2docx" | "weasyprint" | "chrome" | "builtin" (engines that do not apply to a step are ignored).
    options  : format specific, e.g. ``pages="1-3"``, ``mode="image"`` (pptx), ``sheet_mode="table"`` (xlsx).
    """
    src = Path(src)
    if not src.is_file():
        raise FileNotFoundError(src)
    sfmt = detect_format(src)
    if sfmt == "htm":
        sfmt = "html"
    if dst is not None:
        dst = Path(dst)
        to = (to or dst.suffix.lstrip(".")).lower()
    if not to:
        raise ValueError("Specify `to` or a destination path with an extension")
    to = to.lower().lstrip(".")
    if to not in TARGETS:
        raise UnsupportedConversionError(f"Unsupported target format: {to}")
    if dst is None:
        dst = src.with_suffix("." + to)
    if dst.exists() and not overwrite:
        raise FileExistsError(dst)
    dst.parent.mkdir(parents=True, exist_ok=True)

    canon = LEGACY.get(sfmt, sfmt)
    if sfmt == to:
        raise UnsupportedConversionError("Source and target formats are the same")

    from .html_to_pdf import html_to_pdf
    from .office_to_pdf import office_to_pdf

    to_pdf = html_to_pdf if sfmt == "html" else office_to_pdf

    try:
        if sfmt == "pdf":
            _pdf_to(to)(src, dst, engine=engine, **options)
        elif to == "pdf":
            to_pdf(src, dst, engine=engine, **options)
        else:  # office -> office via PDF
            with tempfile.TemporaryDirectory(prefix="docconvert_") as td:
                mid = Path(td) / "mid.pdf"
                pdf_opts = {k: v for k, v in options.items() if k in ("timeout",)}
                to_pdf(src, mid, engine=engine, **pdf_opts)
                rest = {k: v for k, v in options.items() if k not in ("timeout",)}
                _pdf_to(to)(mid, dst, engine=engine, **rest)
    except (ConversionError, UnsupportedConversionError):
        raise
    except ImportError:
        raise
    except Exception as exc:
        if dst.exists() and dst.stat().st_size == 0:
            dst.unlink()
        raise ConversionError(f"{sfmt} -> {to} failed: {exc}") from exc
    return dst


def convert_many(paths: Iterable, to: str, outdir=None, **kwargs) -> list[Path]:
    """Convert several files; returns the output paths."""
    outs = []
    for p in paths:
        p = Path(p)
        out = (Path(outdir) / p.name if outdir else p).with_suffix("." + to)
        outs.append(convert(p, out, to=to, **kwargs))
    return outs
