from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.opc.part import Part

from thaidocs.check import check
from thaidocs.layout import finalize

CUSTOM_CT = "application/vnd.openxmlformats-officedocument.custom-properties+xml"
CUSTOM_RT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties"
_CUSTOM = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" '
    'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
    '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="thaidocs.wrapped">'
    "<vt:lpwstr>1</vt:lpwstr></property></Properties>"
)


def codes(issues):
    return {i.code for i in issues}


def test_check_finds_no_cs(make_doc):
    assert "NO_CS" in codes(check(make_doc(["สวัสดีชาวโลก"])))


def test_check_clean_after_finalize(make_doc):
    doc = make_doc(["บทนำ", ("เนื้อหาไทย 25", None)], table=True, header="หัวกระดาษ", footnote=True)
    finalize(doc)
    assert codes(check(doc)) <= {"WRAPPED"}


def test_check_other_codes(make_doc):
    doc = make_doc(["ก ข 2569"])
    p = doc.paragraphs[0]
    p.add_run("x").add_break()
    p.add_run("y")
    got = {i.code: i for i in check(doc, thai_digits=True)}
    assert {"BAD_CHARS", "MANUAL_BREAK", "ARABIC_DIGITS"} <= set(got)
    assert got["ARABIC_DIGITS"].count == 4
    assert "ARABIC_DIGITS" not in codes(check(doc))


def test_check_wrapped_ruling_a(make_doc):
    doc = make_doc(["สวัสดีชาวโลก"])
    part = Part(PackURI("/docProps/custom.xml"), CUSTOM_CT, _CUSTOM.encode("utf-8"), doc.part.package)
    doc.part.package.relate_to(part, CUSTOM_RT)
    found = codes(check(doc))
    assert "NO_CS" not in found
    assert "WRAPPED" in found


def test_check_finds_no_lang(make_doc):
    from docx.oxml.ns import qn
    doc = make_doc(["สวัสดีชาวโลก"])
    finalize(doc)
    assert "NO_LANG" not in codes(check(doc))
    rpr = doc.paragraphs[0].runs[0]._r.find(qn("w:rPr"))
    rpr.remove(rpr.find(qn("w:lang")))
    got = {i.code: i for i in check(doc)}
    assert "NO_LANG" in got and got["NO_LANG"].count == 1
    assert "NO_CS" not in got


def test_check_no_lang_not_reported_when_wrapped(make_doc):
    from docx.oxml.ns import qn
    doc = make_doc(["สวัสดีชาวโลก"])
    finalize(doc)
    rpr = doc.paragraphs[0].runs[0]._r.find(qn("w:rPr"))
    rpr.remove(rpr.find(qn("w:lang")))
    part = Part(PackURI("/docProps/custom.xml"), CUSTOM_CT, _CUSTOM.encode("utf-8"), doc.part.package)
    doc.part.package.relate_to(part, CUSTOM_RT)
    assert "NO_LANG" not in codes(check(doc))
