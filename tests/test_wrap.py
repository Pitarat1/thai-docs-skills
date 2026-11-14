"""เทส wrap / unwrap: แยกคำละ run ไม่ติดป้ายไทย + ZWSP ระหว่างคำไทยที่ติดกัน"""
from __future__ import annotations

import copy

import pytest
from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import parse_xml
from docx.oxml.ns import qn
from lxml import etree

from thaidocs.check import check
from thaidocs.layout import THAI, finalize
from thaidocs.ooxml import iter_content_parts
from thaidocs.wordcut import Segmenter
from thaidocs.wrap import ZWSP, is_wrapped, unwrap, wrap

_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


@pytest.fixture(scope="module")
def seg():
    return Segmenter()


def rtext(r) -> str:
    return "".join(t.text or "" for t in r.iter(qn("w:t")))


def ptext(p) -> str:
    return "".join(t.text or "" for t in p.iter(qn("w:t")))


def all_runs(doc):
    for root in iter_content_parts(doc):
        yield from root.iter(qn("w:r"))


def all_paras(doc):
    for root in iter_content_parts(doc):
        yield from root.iter(qn("w:p"))


def texts(doc) -> list[str]:
    return [ptext(p) for p in all_paras(doc)]


def xml_all(doc) -> list[bytes]:
    return [etree.tostring(root) for root in iter_content_parts(doc)]


def has_cs(r) -> bool:
    rpr = r.find(qn("w:rPr"))
    return rpr is not None and rpr.find(qn("w:cs")) is not None


def lang_bidi(r):
    lang = r.find(f"{qn('w:rPr')}/{qn('w:lang')}")
    return None if lang is None else lang.get(qn("w:bidi"))


def test_wrap_marks_runs(make_doc, seg, tmp_path):
    src = "โรงพยาบาลอำเภอวังน้ำเย็นจังหวัดสระแก้ว"
    doc = make_doc([src])
    finalize(doc)
    assert wrap(doc, seg) == 1
    assert is_wrapped(doc)
    p = doc.element.body.find(qn("w:p"))
    runs = p.findall(qn("w:r"))
    thai = [r for r in runs if THAI.search(rtext(r))]
    assert thai and not any(has_cs(r) for r in thai)
    assert all(lang_bidi(r) == "en-US" for r in thai)
    for r in thai:  # ฟอนต์ 4 ช่องเท่ากัน, sz == szCs
        rf = r.find(f"{qn('w:rPr')}/{qn('w:rFonts')}")
        assert len({rf.get(qn(a)) for a in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia")}) == 1
        rpr = r.find(qn("w:rPr"))
        sz, szcs = rpr.find(qn("w:sz")), rpr.find(qn("w:szCs"))
        assert (sz is None) == (szcs is None)
        assert sz is None or sz.get(qn("w:val")) == szcs.get(qn("w:val"))
    assert ptext(p).replace(ZWSP, "") == src
    words = [rtext(r) for r in runs]
    i = words.index("อำเภอ" + ZWSP)
    assert words[i + 1].startswith("วังน้ำเย็น")
    assert not words[-1].endswith(ZWSP)
    # check รู้ว่า wrap แล้ว: ไม่รายงาน NO_CS
    codes = {x.code for x in check(doc)}
    assert "WRAPPED" in codes and "NO_CS" not in codes
    # บันทึกแล้วเปิดใหม่ด้วย python-docx ได้ และยังรู้ว่า wrap แล้ว
    path = tmp_path / "wrapped.docx"
    doc.save(path)
    again = Document(str(path))
    assert is_wrapped(again)
    assert again.paragraphs[0].text.replace(ZWSP, "") == src


def test_wrap_idempotent(make_doc, seg):
    doc = make_doc(["ระบบ HIS ของโรงพยาบาลอำเภอวังน้ำเย็น", "ประชุมคณะกรรมการ\tวันที่ 5"],
                   table=True, header="หัวกระดาษภาษาไทย", footnote=True)
    finalize(doc)
    wrap(doc, seg)
    once = xml_all(doc)
    wrap(doc, seg)
    assert xml_all(doc) == once
    assert not any(ZWSP * 2 in t for t in texts(doc))


def test_roundtrip_text_equal(make_doc, seg):
    doc = make_doc(["สำนักงานสาธารณสุขจังหวัดสระแก้ว", "ระบบ HIS ของโรงพยาบาล", "Plain English"],
                   table=True, header="หัวกระดาษภาษาไทย", footnote=True)
    finalize(doc)
    before = texts(doc)
    body_before = etree.tostring(doc.element.body)
    wrap(doc, seg)
    assert unwrap(doc) > 0
    assert texts(doc) == before
    assert not is_wrapped(doc)
    for r in all_runs(doc):
        if THAI.search(rtext(r)):
            assert has_cs(r) and lang_bidi(r) == "th-TH"
    # ย่อหน้าธรรมดาที่ finalize แล้ว: unwrap คืน XML เดิมทุกตัวอักษร
    assert etree.tostring(doc.element.body) == body_before
    # unwrap แล้ว wrap ใหม่ได้ผลเท่า wrap ครั้งแรก
    wrap(doc, seg)
    once = xml_all(doc)
    unwrap(doc)
    wrap(doc, seg)
    assert xml_all(doc) == once


def test_wrap_keeps_formatting(make_doc, seg):
    doc = make_doc([""])
    p = doc.paragraphs[0]
    p.add_run("ข้อความปกติ")
    p.add_run("ตัวหนามาก").bold = True
    p.add_run("และ")
    p.add_run("ยกกำลังสอง").font.superscript = True
    finalize(doc)
    wrap(doc, seg)
    runs = p._p.findall(qn("w:r"))
    bold = [r for r in runs if rtext(r).replace(ZWSP, "") in ("ตัว", "หนา", "มาก", "ตัวหนา", "ตัวหนามาก")]
    sup = [r for r in runs if rtext(r).replace(ZWSP, "") in ("ยกกำลัง", "สอง", "ยกกำลังสอง")]
    assert bold and all(r.find(f"{qn('w:rPr')}/{qn('w:b')}") is not None for r in bold)
    assert sup and all(
        r.find(f"{qn('w:rPr')}/{qn('w:vertAlign')}").get(qn("w:val")) == "superscript" for r in sup)
    plain = [r for r in runs if rtext(r).replace(ZWSP, "") in ("ข้อความ", "ปกติ", "และ")]
    assert plain and all(r.find(f"{qn('w:rPr')}/{qn('w:b')}") is None for r in plain)
    assert ptext(p._p).replace(ZWSP, "") == "ข้อความปกติตัวหนามากและยกกำลังสอง"


def _field_link_para(doc, rid):
    xml = (
        f'<w:p xmlns:w="{_W}" xmlns:r="{_R}">'
        "<w:r><w:t xml:space=\"preserve\">หน้าที่ </w:t></w:r>"
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
        '<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>'
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
        "<w:r><w:t>หนึ่งร้อย</w:t></w:r>"
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
        "<w:r><w:t>ดูที่</w:t></w:r>"
        f'<w:hyperlink r:id="{rid}"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr>'
        "<w:t>เว็บไซต์โรงพยาบาลวังน้ำเย็น</w:t></w:r></w:hyperlink>"
        "<w:r><w:t>สำหรับประชาชน</w:t></w:r>"
        "</w:p>"
    )
    p = parse_xml(xml)
    doc.element.body.insert(0, p)
    return p


def test_wrap_keeps_fields_and_links(make_doc, seg):
    doc = make_doc(["ย่อหน้าอื่น"])
    rid = doc.part.relate_to("https://example.go.th/", RT.HYPERLINK, is_external=True)
    p = _field_link_para(doc, rid)
    field_before = [copy.deepcopy(r) for r in p.findall(qn("w:r"))[1:6]]
    wrap(doc, seg)
    instr = list(p.iter(qn("w:instrText")))
    assert [i.text for i in instr] == [" PAGE "]
    runs = p.findall(qn("w:r"))
    # run ของ field (begin … end) ไม่ถูกแตะเลย
    begin = next(i for i, r in enumerate(runs) if r.find(qn("w:fldChar")) is not None)
    def c14n(r):
        return etree.tostring(r, method="c14n", exclusive=True)

    assert [c14n(r) for r in runs[begin:begin + 5]] == [c14n(r) for r in field_before]
    links = p.findall(qn("w:hyperlink"))
    assert len(links) == 1 and links[0].get(qn("r:id")) == rid
    assert ptext(links[0]).replace(ZWSP, "") == "เว็บไซต์โรงพยาบาลวังน้ำเย็น"
    assert len(links[0].findall(qn("w:r"))) > 1  # แยกคำในลิงก์ด้วย
    # คำสุดท้ายในลิงก์ติดกับ "สำหรับ" (ไทย-ไทย) → ZWSP ท้ายลิงก์
    assert ptext(links[0]).endswith(ZWSP)
    assert ptext(p).replace(ZWSP, "") == "หน้าที่ หนึ่งร้อยดูที่เว็บไซต์โรงพยาบาลวังน้ำเย็นสำหรับประชาชน"
    # ไม่มี ZWSP ข้าม field: "ดูที่" นำหน้าด้วย field end ไม่ต้องมี ZWSP ก่อน
    assert all(ZWSP not in rtext(r) for r in runs[begin:begin + 5])
    # ลิงก์ยังใช้ได้หลังบันทึก
    assert doc.part.rels[rid].target_ref == "https://example.go.th/"


def test_wrap_mixed_script(make_doc, seg):
    doc = make_doc(["ระบบ HIS ของโรงพยาบาล"])
    wrap(doc, seg)
    t = doc.paragraphs[0].text
    assert t.replace(ZWSP, "") == "ระบบ HIS ของโรงพยาบาล"
    for i, ch in enumerate(t):
        if ch == ZWSP:
            assert THAI.match(t[i - 1]) and THAI.match(t[i + 1])
    assert "ของ" + ZWSP + "โรงพยาบาล" in t


def test_wrap_nbsp_no_zwsp(make_doc, seg):
    doc = make_doc(["นายสมชาย ใจดีมาประชุม"])
    wrap(doc, seg)
    t = doc.paragraphs[0].text
    assert " " + ZWSP not in t and ZWSP + " " not in t
    assert t.replace(ZWSP, "") == "นายสมชาย ใจดีมาประชุม"


def test_wrap_tab_is_separator(make_doc, seg):
    doc = make_doc(["ชื่อโรงพยาบาล\tวังน้ำเย็นอำเภอ"])
    wrap(doc, seg)
    p = doc.paragraphs[0]._p
    assert len(list(p.iter(qn("w:tab")))) == 1
    t = ptext(p)
    assert "โรงพยาบาล" in t and "โรงพยาบาล" + ZWSP not in t  # ไม่มี ZWSP ข้ามแท็บ
    seq = "".join("\t" if e.tag == qn("w:tab") else (e.text or "") for e in p.iter(qn("w:t"), qn("w:tab")))
    assert seq.replace(ZWSP, "") == "ชื่อโรงพยาบาล\tวังน้ำเย็นอำเภอ"


def test_wrap_tables_headers(make_doc, seg):
    doc = make_doc(["เนื้อหา"], table=True, header="หัวกระดาษภาษาไทยสำหรับทดสอบ")
    finalize(doc)
    wrap(doc, seg)
    cells = [c.text for c in doc.tables[0].rows[0].cells]
    assert all(ZWSP in c for c in cells)
    head = doc.sections[0].header.paragraphs[0]
    assert ZWSP in head.text
    assert not any(has_cs(r) for r in head._p.iter(qn("w:r")) if THAI.search(rtext(r)))
    for cell in doc.tables[0].rows[0].cells:
        assert not any(has_cs(r) for r in cell._tc.iter(qn("w:r")) if THAI.search(rtext(r)))


def test_wrap_footnotes(make_doc, seg):
    doc = make_doc(["ข้อความมีเชิงอรรถ"], footnote=True)
    finalize(doc)
    wrap(doc, seg)
    roots = list(iter_content_parts(doc))
    fn = next(r for r in roots if etree.QName(r).localname == "footnotes")
    assert ZWSP in ptext(fn)
    assert len(list(fn.iter(qn("w:footnoteRef")))) == 1
    assert len(list(doc.element.body.iter(qn("w:footnoteReference")))) == 1


def test_wrap_leaves_empty_and_english(make_doc, seg):
    doc = make_doc(["", "Hello world, plain English 2026.", "ภาษาไทยหนึ่งย่อหน้า"])
    ps = doc.element.body.findall(qn("w:p"))
    before = [etree.tostring(p) for p in ps[:2]]
    assert wrap(doc, seg) == 1
    assert [etree.tostring(p) for p in ps[:2]] == before


def test_wrap_property_part(make_doc, seg, tmp_path):
    import zipfile

    from thaidocs.ooxml import get_custom_property, set_custom_property

    doc = make_doc(["ภาษาไทยขนาดใหญ่"])
    doc.paragraphs[0].runs[0].font.size = 254000  # 20 pt: มีแค่ sz → wrap เติม szCs
    set_custom_property(doc, "other", "x")
    wrap(doc, seg)
    rpr = doc.paragraphs[0]._p.find(f"{qn('w:r')}/{qn('w:rPr')}")
    assert rpr.find(qn("w:szCs")).get(qn("w:val")) == "40"
    path = tmp_path / "p.docx"
    doc.save(path)
    with zipfile.ZipFile(path) as z:
        assert "application/vnd.openxmlformats-officedocument.custom-properties+xml" in z.read(
            "[Content_Types].xml").decode()
        xml = z.read("docProps/custom.xml").decode()
    assert 'pid="2" name="other"' in xml and 'pid="3" name="thaidocs.wrapped"' in xml
    again = Document(str(path))
    unwrap(again)
    assert not is_wrapped(again) and get_custom_property(again, "other") == "x"
