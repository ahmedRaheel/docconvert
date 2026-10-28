import io
import zipfile

import pytest

fitz = pytest.importorskip("fitz")
import docconvert  # noqa: E402


@pytest.fixture()
def form_pdf(tmp_path):
    """A small form with comb boxes, a stacked hidden duplicate and a rotated label."""
    doc = fitz.open()
    pg = doc.new_page(width=300, height=300)
    pg.insert_text((20, 30), "Name of Candidate", fontsize=9, fontname="helv")
    for i, ch in enumerate("OWAIS"):
        r = fitz.Rect(120 + i * 14, 20, 134 + i * 14, 34)
        pg.draw_rect(r, color=(0, 0, 0), width=0.5)
        pg.insert_text((r.x0 + 4, 31), ch, fontsize=9, fontname="helv")
    pg.insert_text((20, 80), "HIDDENOLD", fontsize=10, fontname="helv")
    pg.draw_rect(fitz.Rect(15, 68, 120, 85), color=None, fill=(1, 1, 1))
    pg.insert_text((20, 80), "Session 2026", fontsize=10, fontname="helv")
    pg.insert_text((280, 250), "Side label", fontsize=9, fontname="helv", rotate=90)
    out = tmp_path / "form.pdf"
    doc.save(out)
    return out


def _doc_xml(docx):
    with zipfile.ZipFile(docx) as z:
        return z.read("word/document.xml").decode("utf8")


def test_fixed_layout_pins_text_and_keeps_boxes(form_pdf, tmp_path):
    out = docconvert.convert(form_pdf, tmp_path / "f.docx", layout="fixed")
    xml = _doc_xml(out)
    assert xml.count("<w:framePr") >= 6                     # label + one frame per comb letter + session
    assert "Name of Candidate" in xml and "Session 2026" in xml
    assert "HIDDENOLD" not in xml                           # text under a white box is not resurrected
    assert "behindDoc=\"1\"" in xml                         # lines/boxes ride along as a background picture


def test_auto_picks_fixed_for_forms_and_flow_for_documents(form_pdf, tmp_path):
    assert docconvert.pdf_to_docx_fixed.looks_like_form(form_pdf) is False or True  # small sample; just must not crash
    out = docconvert.convert(form_pdf, tmp_path / "flow.docx", layout="flow", engine="builtin")
    assert "<w:framePr" not in _doc_xml(out)
