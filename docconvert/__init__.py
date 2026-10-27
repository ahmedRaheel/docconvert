"""docconvert - convert between PDF, Word, Excel and PowerPoint, preserving source formatting.

>>> import docconvert
>>> docconvert.convert("report.pdf", "report.docx")
>>> docconvert.convert("deck.pptx", to="pdf")
>>> docconvert.convert("page.html", "page.pdf")
>>> docconvert.convert("sheet.xlsx", "sheet.pptx")   # office -> office goes through PDF
"""
from .core import available_engines, convert, convert_many, detect_format, supported_conversions
from .errors import ConversionError, DocConvertError, MissingDependencyError, UnsupportedConversionError

# Convenience wrappers ------------------------------------------------------
def pdf_to_word(src, dst=None, **kw):       return convert(src, dst, to="docx", **kw)
def pdf_to_excel(src, dst=None, **kw):      return convert(src, dst, to="xlsx", **kw)
def pdf_to_powerpoint(src, dst=None, **kw): return convert(src, dst, to="pptx", **kw)
def word_to_pdf(src, dst=None, **kw):       return convert(src, dst, to="pdf", **kw)
def excel_to_pdf(src, dst=None, **kw):      return convert(src, dst, to="pdf", **kw)
def powerpoint_to_pdf(src, dst=None, **kw): return convert(src, dst, to="pdf", **kw)
def html_to_pdf(src, dst=None, **kw):       return convert(src, dst, to="pdf", **kw)


def html_string_to_pdf(html: str, dst, **kw):
    """Render an HTML string (relative links resolve against the current directory)."""
    import tempfile
    from pathlib import Path
    with tempfile.TemporaryDirectory() as td:
        f = Path(td) / "page.html"
        f.write_text(html, encoding="utf-8")
        return convert(f, dst, to="pdf", **kw)


__all__ = [
    "convert", "convert_many", "detect_format", "supported_conversions", "available_engines",
    "pdf_to_word", "pdf_to_excel", "pdf_to_powerpoint", "word_to_pdf", "excel_to_pdf", "powerpoint_to_pdf", "html_to_pdf", "html_string_to_pdf",
    "DocConvertError", "ConversionError", "UnsupportedConversionError", "MissingDependencyError",
]
__version__ = "0.1.0"
