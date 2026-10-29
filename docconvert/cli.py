from __future__ import annotations

import argparse
import sys

from . import __version__, available_engines, convert
from .errors import DocConvertError


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(prog="docconvert", description="Convert between PDF, Word, Excel and PowerPoint.")
    ap.add_argument("input", nargs="?")
    ap.add_argument("-t", "--to", choices=["pdf", "docx", "xlsx", "pptx"], help="target format")
    ap.add_argument("-o", "--output", help="output file (default: next to input)")
    ap.add_argument("-e", "--engine", default="auto", choices=["auto", "libreoffice", "pdf2docx", "weasyprint", "chrome", "builtin"])
    ap.add_argument("-p", "--pages", help='pages for PDF input, e.g. "1-3,5"')
    ap.add_argument("--mode", choices=["auto", "exact", "editable", "image"], help="PDF->PPTX mode")
    ap.add_argument("--sheet-mode", choices=["page", "table"], help="PDF->XLSX sheet layout")
    ap.add_argument("--layout", choices=["auto", "fixed", "flow"], help="PDF->Word: exact copy (forms) or flowing text")
    ap.add_argument("--list-engines", action="store_true")
    ap.add_argument("--version", action="version", version=__version__)
    a = ap.parse_args(argv)

    if a.list_engines:
        for k, v in available_engines().items():
            print(f"{k:12} {v if v else 'not available'}")
        return 0
    if not a.input or not (a.to or a.output):
        ap.error("input and --to/--output are required")
    opts = {k: v for k, v in dict(pages=a.pages, mode=a.mode, sheet_mode=a.sheet_mode, layout=a.layout).items() if v}
    try:
        out = convert(a.input, a.output, to=a.to, engine=a.engine, **opts)
    except (DocConvertError, FileNotFoundError, FileExistsError, ValueError) as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 1
    print(out)
    return 0


if __name__ == "__main__":
    sys.exit(main())
