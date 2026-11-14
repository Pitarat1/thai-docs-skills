from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import qn

from thaidocs.digits import apply_thai_digits, to_thai_digits

_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def test_to_thai_digits():
    assert to_thai_digits("ปี 2569 ดู https://a.b/12 HC3") == "ปี ๒๕๖๙ ดู https://a.b/12 HC3"
    assert to_thai_digits("มาตรฐาน ISO 27001 ข้อ 5") == "มาตรฐาน ISO 27001 ข้อ ๕"


def test_instrtext_untouched(make_doc):
    doc = make_doc(["ข้อความ"])
    p = doc.paragraphs[0]._p
    p.append(parse_xml(
        f'<w:r xmlns:w="{_W}"><w:fldChar w:fldCharType="begin"/></w:r>'))
    p.append(parse_xml(f'<w:r xmlns:w="{_W}"><w:instrText> TOC \\o "1-3" </w:instrText></w:r>'))
    p.append(parse_xml(f'<w:r xmlns:w="{_W}"><w:t>หน้า 3</w:t></w:r>'))
    apply_thai_digits(doc)
    assert next(p.iter(qn("w:instrText"))).text == ' TOC \\o "1-3" '
    assert "หน้า ๓" in "".join(t.text for t in p.iter(qn("w:t")))


def test_split_runs():
    doc = Document()
    para = doc.add_paragraph()
    para.add_run("ปี ")
    para.add_run("25")
    para.add_run("69")
    apply_thai_digits(doc)
    assert [r.text for r in para.runs] == ["ปี ", "๒๕", "๖๙"]


def test_sectpr_and_numbering_and_footnote(make_doc):
    doc = make_doc(["ก 1"], footnote=True)
    apply_thai_digits(doc)
    sect = doc.element.body.find(qn("w:sectPr"))
    pg = sect.find(qn("w:pgNumType"))
    assert pg.get(qn("w:fmt")) == "thaiNumbers"
    kids = [k.tag.split("}")[1] for k in sect]
    assert kids.index("pgMar") < kids.index("pgNumType") < kids.index("cols")
    assert apply_thai_digits(doc) == 0


def test_keep_latin_joined_codes():
    for code in ("COVID-19", "v2.0", "SARS-CoV-2", "GPT-4", "H5N1-2"):
        assert to_thai_digits(f"ระบาด {code} ปี 2569") == f"ระบาด {code} ปี ๒๕๖๙"
    assert to_thai_digits("ข้อ 3.2") == "ข้อ ๓.๒"
    assert to_thai_digits("ปี 2569") == "ปี ๒๕๖๙"


def test_url_stops_at_thai_text():
    assert to_thai_digits("ดูhttps://a.b/1ปี2569") == "ดูhttps://a.b/1ปี๒๕๖๙"
