"""ตัวสร้าง .docx: หนังสือโต้ตอบ (บันทึกข้อความ หนังสือภายนอก หนังสือประทับตรา)

ค่าที่วัดอ่านจากไฟล์ที่บันทึกแล้วเปิดกลับ เทียบกับ LayoutSpec (คลาดเคลื่อน ≤ 0.02 ซม. / 0.1 pt)
"""
from __future__ import annotations

import re
import shutil
from dataclasses import replace

import pytest
from docx.enum.text import WD_COLOR_INDEX, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

from thaidocs.check import check
from thaidocs.digits import KEEP
from thaidocs.official.profile import PLACEHOLDER_EMBLEM
from thaidocs.official.render import DocxRenderer
from thaidocs.official.schema import load_example, parse_doc
from thaidocs.official.style import DEFAULT_SPEC
from thaidocs.official.textutil import PLACEHOLDER_RE

S = DEFAULT_SPEC
LETTERS = ["memo", "external_letter", "stamp_letter"]
CM_TOL = 0.02
PT_TOL = 0.1
EMU_PER_CM = 360000


def cm(length) -> float:
    return (length or 0) / EMU_PER_CM


def pt(length) -> float:
    return (length or 0) / 12700


def approx_cm(length, want):
    return abs(cm(length) - want) <= CM_TOL


def approx_pt(length, want):
    return abs(pt(length) - want) <= PT_TOL


def body_runs(doc):
    """(run element, text) ของทุก run ที่มี w:t ในเนื้อหาและหัว/ท้ายกระดาษ"""
    sec = doc.sections[0]
    roots = [doc.element.body]
    for part in (sec.header, sec.footer, sec.first_page_header, sec.first_page_footer):
        if not part.is_linked_to_previous:
            roots.append(part._element)
    for root in roots:
        for r in root.iter(qn("w:r")):
            ts = list(r.iter(qn("w:t")))
            if ts:
                yield r, "".join(t.text or "" for t in ts)


def all_text(doc) -> str:
    return "\n".join(t for _, t in body_runs(doc))


def pictures(doc):
    """ความสูง (ซม.) ของรูปทุกรูปในเนื้อเอกสาร (inline หรือ anchor)"""
    out = []
    for tag in ("wp:inline", "wp:anchor"):
        for el in doc.element.body.iter(qn(tag)):
            out.append(int(el.find(qn("wp:extent")).get("cy")) / EMU_PER_CM)
    return out


def jc(p):
    el = p._p.find(f"{qn('w:pPr')}/{qn('w:jc')}")
    return el.get(qn("w:val")) if el is not None else None


def is_bold(r):
    rpr = r.find(qn("w:rPr"))
    b = rpr.find(qn("w:b")) if rpr is not None else None
    return b is not None and b.get(qn("w:val"), "1") in ("1", "true", "on")


def size_pt(r):
    sz = r.find(f"{qn('w:rPr')}/{qn('w:sz')}")
    return int(sz.get(qn("w:val"))) / 2 if sz is not None else None


def color(r):
    c = r.find(f"{qn('w:rPr')}/{qn('w:color')}")
    return c.get(qn("w:val")) if c is not None else None


# ------------------------------------------------------------------ ทั้ง 3 ชนิด
@pytest.mark.parametrize("doc_type", LETTERS)
def test_page_setup(render_example, doc_type):
    doc = render_example(doc_type)
    sec = doc.sections[0]
    assert approx_cm(sec.page_width, 21.0) and approx_cm(sec.page_height, 29.7)
    assert approx_cm(sec.left_margin, 3.0)
    assert approx_cm(sec.right_margin, 2.0)
    assert approx_cm(sec.top_margin, 1.5)
    assert approx_cm(sec.bottom_margin, 2.0)


@pytest.mark.parametrize("doc_type", LETTERS)
def test_every_run_uses_font(render_example, doc_type):
    doc = render_example(doc_type)
    n = 0
    for r, _ in body_runs(doc):
        rf = r.find(f"{qn('w:rPr')}/{qn('w:rFonts')}")
        assert rf is not None
        for a in ("w:ascii", "w:hAnsi", "w:cs"):
            assert rf.get(qn(a)) == "TH Sarabun New"
        n += 1
    assert n > 10


@pytest.mark.parametrize("doc_type", LETTERS)
def test_core_check_passes(render_example, doc_type):
    assert check(render_example(doc_type)) == []


@pytest.mark.parametrize("doc_type", LETTERS)
def test_page_number_field_and_first_page(render_example, doc_type):
    doc = render_example(doc_type)
    sec = doc.sections[0]
    assert sec.different_first_page_header_footer
    instr = "".join(e.text or "" for e in sec.header._element.iter(qn("w:instrText")))
    assert "PAGE" in instr


# ------------------------------------------------------------------ บันทึกข้อความ
def test_memo_title_and_labels(render_example, para):
    doc = render_example("memo")
    p = para(doc, "บันทึกข้อความ")
    assert p is not None
    title_runs = [r for r in p._p.iter(qn("w:r")) if "บันทึกข้อความ" in "".join(t.text or "" for t in r.iter(qn("w:t")))]
    assert title_runs and all(is_bold(r) and abs(size_pt(r) - 29) <= PT_TOL for r in title_runs)
    pf = p.paragraph_format
    assert pf.line_spacing_rule == WD_LINE_SPACING.EXACTLY
    assert approx_pt(pf.line_spacing, 35)

    labels = {"ส่วนราชการ", "ที่", "วันที่", "เรื่อง"}
    found = set()
    for r, text in body_runs(doc):
        if text in labels:
            assert is_bold(r) and abs(size_pt(r) - 20) <= PT_TOL, text
            found.add(text)
    assert found == labels


def test_memo_emblem_height(render_example):
    hs = pictures(render_example("memo"))
    assert len(hs) == 1 and abs(hs[0] - 1.5) <= CM_TOL


# ------------------------------------------------------------------ หนังสือภายนอก
def test_external_emblem_height(render_example):
    hs = pictures(render_example("external_letter"))
    assert len(hs) == 1 and abs(hs[0] - 3.0) <= CM_TOL


def test_external_body_paragraphs(render_example, para):
    doc = render_example("external_letter")
    body = load_example("external_letter")["body"]
    texts = [*body["cause"], *body["purpose"], body["conclusion"]]
    for t in texts:
        p = para(doc, t[:12])
        assert p is not None, t
        assert jc(p) == "thaiDistribute"
        pf = p.paragraph_format
        assert approx_cm(pf.first_line_indent, 2.5)
        assert approx_pt(pf.space_before, 6)


def test_external_closing_signature_and_date(render_example, para):
    doc = render_example("external_letter")
    ps = doc.paragraphs
    closing = para(doc, "ขอแสดงความนับถือ")
    assert closing is not None and approx_pt(closing.paragraph_format.space_before, 12)
    i = next(k for k, p in enumerate(ps) if p._p is closing._p)
    assert [p.text for p in ps[i + 1:i + 4]] == ["", "", ""]
    assert ps[i + 4].text.strip() == "(นายสมมติ ตัวอย่าง)"
    date = para(doc, "๑๔\xa0พฤศจิกายน")  # protect_breaks: วันที่กับเดือนไม่แยกบรรทัด
    assert date is not None and approx_cm(date.paragraph_format.left_indent, 8.0)


# ------------------------------------------------------------------ หนังสือประทับตรา
def test_stamp_to_and_no_closing(render_example, para):
    doc = render_example("stamp_letter")
    to = para(doc, "ถึง")
    assert to is not None and approx_pt(to.paragraph_format.space_before, 8)
    assert para(doc, "ขอแสดงความนับถือ") is None


# ------------------------------------------------------------------ [ระบุ…]
def test_placeholders_highlighted_only(render_example):
    doc = render_example("external_letter", phone="[ระบุ]",
                         **{"body.cause": ["ตามที่ประชุมเมื่อวันที่ [ระบุวันที่] ได้มีมติให้ดำเนินการ"]})
    seen = 0
    for r, text in body_runs(doc):
        hl = r.find(f"{qn('w:rPr')}/{qn('w:highlight')}")
        if PLACEHOLDER_RE.fullmatch(text):
            assert hl is not None and hl.get(qn("w:val")) == "yellow", text
            seen += 1
        else:
            assert hl is None, text
    assert seen >= 2


# ------------------------------------------------------------------ เลขไทย / อารบิก
# เลขที่ไม่ซ้ำกับเลขไทยในตัวอย่าง จึงตรวจได้ว่าเลขไทยที่พบไม่ได้แปลงมาจากร่าง
ARABIC_OVERRIDES = {
    "doc_no": "ตย 0009/987",
    "body.cause": ["ด้วยมีผู้เข้าร่วมจำนวน 47 คน ดูรายละเอียดที่ https://example.org/a1"],
}


def arabic_overrides(doc_type):
    """ชนิดหนังสือประทับตราเก็บข้อความใน paragraphs แทน body.cause"""
    if doc_type == "stamp_letter":
        return {"doc_no": ARABIC_OVERRIDES["doc_no"], "paragraphs": ARABIC_OVERRIDES["body.cause"]}
    return ARABIC_OVERRIDES


@pytest.mark.parametrize("doc_type", LETTERS)
def test_thai_digits_by_default(render_example, doc_type):
    text = all_text(render_example(doc_type, **arabic_overrides(doc_type)))
    assert "๐๐๐๙/๙๘๗" in text and "๔๗" in text
    assert "https://example.org/a1" in text
    assert not re.search("[0-9]", KEEP.sub("", text))


@pytest.mark.parametrize("doc_type", LETTERS)
def test_arabic_digits_when_disabled(render_example, doc_type):
    doc = render_example(doc_type, spec=replace(S, thai_numerals=False), **arabic_overrides(doc_type))
    text = all_text(doc)
    assert "0009/987" in text and re.search(r"จำนวน\s47\s", text)
    assert "๐๐๐๙/๙๘๗" not in text and "๔๗" not in text
    instr = "".join(e.text or "" for e in doc.sections[0].header._element.iter(qn("w:instrText")))
    assert "ThaiArabic" not in instr


# ------------------------------------------------------------------ ชั้นความเร็ว
@pytest.mark.parametrize("doc_type", LETTERS)
def test_speed_red_bold_large(render_example, doc_type):
    doc = render_example(doc_type, speed="ด่วนที่สุด")
    runs = [r for r, t in body_runs(doc) if t == "ด่วนที่สุด"]
    assert runs
    for r in runs:
        assert is_bold(r) and size_pt(r) >= 32 and color(r) == "FF0000"


# ------------------------------------------------------------------ ส่วนที่ไม่มีข้อมูล
def test_external_empty_optional_parts(render_example, para):
    doc = render_example("external_letter", references=[], attachments=[], copies_to=[], fax=None,
                         **{"signer.name": ""})
    for start in ("อ้างถึง", "สิ่งที่ส่งมาด้วย", "สำเนาส่ง", "โทรสาร"):
        assert para(doc, start) is None, start
    assert "()" not in all_text(doc)


@pytest.mark.parametrize("doc_type", ["memo", "stamp_letter"])
def test_empty_fax_and_signer_name(render_example, para, doc_type):
    over = {"fax": None} if doc_type == "stamp_letter" else {"signer.name": ""}
    doc = render_example(doc_type, **over)
    assert para(doc, "โทรสาร") is None
    assert "()" not in all_text(doc)


def test_omit_signature(render_example, para):
    doc = render_example("memo", spec=replace(S, omit_signature=("memo",)))
    assert para(doc, "(นายสมมติ") is None
    assert para(doc, "ผู้อำนวยการกองตัวอย่าง") is None
    # ชนิดอื่นยังพิมพ์ผู้ลงนาม
    doc = render_example("external_letter", spec=replace(S, omit_signature=("memo",)))
    assert para(doc, "(นายสมมติ") is not None


# ------------------------------------------------------------------ ตรา
def test_placeholder_emblem_flag(render_example, tmp_path):
    doc = render_example("external_letter")
    assert doc.renderer.used_placeholder_emblem is True
    own = tmp_path / "emblem.png"
    shutil.copyfile(PLACEHOLDER_EMBLEM, own)
    doc = render_example("external_letter", emblem=own)
    assert doc.renderer.used_placeholder_emblem is False
    assert len(pictures(doc)) == 1


def test_render_returns_bytes():
    data = DocxRenderer().render(parse_doc(load_example("memo")))
    assert data[:2] == b"PK"
