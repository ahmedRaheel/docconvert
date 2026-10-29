# docconvert

Convert between **PDF, Word (.docx), Excel (.xlsx) and PowerPoint (.pptx)** in any direction, plus **HTML → PDF** (and HTML → Word/Excel/PowerPoint via PDF), keeping the source formatting as far as each format allows.

```bash
pip install .            # core (pure Python)
pip install ".[fidelity]"  # + pdf2docx for best PDF->Word layout
```

```python
import docconvert as dc

dc.convert("report.pdf", "report.docx")            # PDF -> Word
dc.convert("tables.pdf", "tables.xlsx", sheet_mode="table")
dc.convert("slides.pdf", "slides.pptx", mode="editable")   # or mode="image"
dc.convert("deck.pptx", to="pdf")                  # Office -> PDF
dc.convert("page.html", "page.pdf")                # HTML -> PDF
dc.html_string_to_pdf("<h1>Hi</h1>", "hi.pdf")      # HTML string -> PDF
dc.convert("sheet.xlsx", "sheet.docx")             # Office -> Office (via PDF)
dc.convert("big.pdf", "part.docx", pages="1-3,7")
```

```bash
docconvert report.pdf -t docx
docconvert deck.pptx -o deck.pdf --engine builtin
docconvert --list-engines
```

## What is preserved

| Direction | Engine | Preserves |
|---|---|---|
| PDF → Word | `layout="fixed"` for forms (exact copy: text pinned in place + line art as background), else `pdf2docx` if installed, else `builtin` | layout, fonts, tables, images / headings, bullets, bold/italic/colour, tables, images |
| PDF → Excel | `builtin` | tables → real cells (numbers, %, thousands separators), styled headers; text-only pages split into columns |
| PDF → PowerPoint | `builtin` | one slide per page; positioned text boxes with font/size/colour, shaded rectangles, pictures. Scanned pages fall back to a full-slide image |
| Word/Excel/PowerPoint → PDF | `libreoffice` (if installed) else `builtin` | LibreOffice: full fidelity. Builtin: text styles, tables, pictures, fills, slide positions |
| HTML → PDF | `weasyprint` → `chrome` → `libreoffice` → `builtin` (first available) | weasyprint/chrome: full CSS. Builtin: headings, text styles, links, lists, tables, images (CSS ignored) |
| Office/HTML → Office | via PDF | content and basic structure |

Engines: `auto` (default), `libreoffice`, `pdf2docx`, `weasyprint`, `chrome`, `builtin`. Set `DOCCONVERT_SOFFICE` / `DOCCONVERT_CHROME` for custom install paths. For the best HTML results: `pip install weasyprint`.

## Limits

* Scanned PDFs have no text layer — run OCR first (e.g. `ocrmypdf`) for Word/Excel output.
* Charts, SmartArt and floating objects need the LibreOffice engine for Office → PDF.
* Office ↔ Office conversions are approximate by nature (PDF is a fixed-layout pivot).
* The builtin PDF fallback for non-Latin text uses DejaVu/Liberation/Arial when installed.

## Tests

`pip install ".[test]" && pytest` — exercises all 12 PDF/Word/Excel/PowerPoint pairs and HTML conversions.

## Web interface

A Bootstrap-based web UI lives in `web/` and ships with a Dockerfile and `docker-compose.yml`.
See **DEPLOY.md** for running it on a server next to an existing app (port, reverse-proxy and sub-path setup).

```bash
docker compose up -d --build     # http://localhost:8081
```

### PDF forms to Word

`layout="auto"` (default) detects forms (comb boxes, form fields, heavy line art) and converts them as an
**exact copy**: every box, line, logo and rotated label is kept in a background picture, and all normal text
becomes editable text frames at the original coordinates. Force it with `layout="fixed"`, or use
`layout="flow"` for reflowing Word text. Needs nothing beyond the core install (no PyMuPDF).

## Formatting is copied, not flattened

Every direction reads the real formatting from the source and writes it to the target:

| Source -> target | What is copied |
|---|---|
| Word -> PDF | run font family (sans/serif/mono), size, bold, italic, underline, strike, colour, highlight, sub/superscript; paragraph alignment, indents, spacing, line spacing; bullets/numbering; table shading, borders, merged cells; headers/footers; style inheritance |
| Excel -> PDF | cell font/size/colour/bold/italic/underline/strike, fills, borders, alignment, wrap, indent, merged cells, number/percent/currency/date formats, column widths, row heights, hidden rows/columns, fit-to-width |
| PowerPoint -> PDF | formatting resolved through layout/master/theme (size, bold, italic, underline, colour incl. theme tints, alignment, bullets, spacing, insets, autofit scale, anchors), shape fills/outlines/geometry/rotation, groups, cropped pictures, tables, backgrounds |
| HTML -> PDF | `<style>` rules (tag, .class, #id) and `style=""`: colour, background, font size/weight/style/family, underline/strike, text-align, borders; table cell fills/alignment |
| PDF -> Word | size, colour, bold, italic, font per run (flow mode); exact copy mode for forms |
| PDF -> Excel | per-cell font, size, colour, bold, italic (mixed formatting as rich text), fill, borders, alignment, merged cells, column widths, row heights |
| PDF -> PowerPoint | text boxes with font, size, colour, bold, italic; fills, outlines and ruled lines as shapes; exact copy mode for forms |

With LibreOffice installed, Office/HTML -> PDF uses it instead (full fidelity). The pure-Python renderers above
are the fallback and are the ones to rely on for slim servers.
Fonts that are not installed (Calibri, Arial...) are mapped to the closest installed family, so size, weight and
style survive even when the exact typeface cannot.
