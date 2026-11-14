"""ตัวสร้าง .docx: ชนิดที่เหลือ (คำสั่ง ระเบียบ ข้อบังคับ ประกาศ แถลงการณ์ ข่าว หนังสือรับรอง รายงานการประชุม)"""
from __future__ import annotations

import re

import pytest
from docx.oxml.ns import qn

from thaidocs.check import check
from thaidocs.digits import KEEP
from thaidocs.official.schema import load_example
from .test_render_letters import all_text, approx_cm, is_bold, jc

OTHERS = ["order", "regulation", "bylaw", "announcement", "statement", "news", "certificate", "minutes"]
ALL_TYPES = ["memo", "external_letter", "stamp_letter", *OTHERS]


def norm(text: str) -> str:
    """ช่องว่างไม่ตัดบรรทัด (protect_breaks) และช่องว่างซ้ำ → ช่องว่างเดียว"""
    return re.sub(r"[ \xa0]+", " ", text)


def texts(doc):
    return [norm(p.text).lstrip("\t ") for p in doc.paragraphs]


def has(doc, startswith: str) -> bool:
    return any(t.startswith(norm(startswith)) for t in texts(doc))


def flat_text(doc) -> str:
    return norm(all_text(doc))


def run_bold_of(p):
    runs = [r for r in p._p.iter(qn("w:r")) if "".join(t.text or "" for t in r.iter(qn("w:t")))]
    return runs and all(is_bold(r) for r in runs)


# ------------------------------------------------------------------ ทุกชนิด
@pytest.mark.parametrize("doc_type", ALL_TYPES)
def test_every_type_renders_and_passes_check(render_example, doc_type):
    doc = render_example(doc_type)
    assert check(doc) == []
    sec = doc.sections[0]
    assert approx_cm(sec.page_width, 21.0) and approx_cm(sec.page_height, 29.7)
    assert approx_cm(sec.left_margin, 3.0) and approx_cm(sec.right_margin, 2.0)
    assert approx_cm(sec.bottom_margin, 2.0)
    text = all_text(doc)
    assert text.strip()
    assert not re.search("[0-9]", KEEP.sub("", text))


# ------------------------------------------------------------------ คำสั่ง
def test_order_head_and_date(render_example, para):
    doc = render_example("order")
    head = para(doc, "คำสั่ง")
    assert head is not None and jc(head) == "center" and run_bold_of(head)
    # บรรทัดแรกที่มีข้อความคือหัวคำสั่ง (ก่อนหน้านั้นมีแต่ย่อหน้าตรา)
    assert next(p for p in doc.paragraphs if p.text.strip())._p is head._p
    assert has(doc, "ที่ ๑๒๓/๒๕๖๙")
    assert has(doc, "สั่ง") 
    assert "สั่ง ณ วันที่" in flat_text(doc)


def test_order_members_in_order_and_subgroup_numbers(render_example):
    doc = render_example("order")
    data = load_example("order")
    expected = []
    for c in data["committees"]:
        expected += [m["name"] for m in c.get("members", [])]
        for g in c.get("subgroups", []):
            expected += [m["name"] for m in g["members"]]
    cells = [c.text for t in doc.tables for row in t.rows for c in row.cells]
    flat = "\n".join(cells)
    pos = [flat.index(n) for n in expected]
    assert pos == sorted(pos) and len(set(pos)) == len(pos)
    assert "๒.๑" in flat and "๒.๒" in flat
    assert any(t.startswith("๒.๑") for t in texts(doc))
    assert "๑.๑" in flat


# ------------------------------------------------------------------ ระเบียบ / ข้อบังคับ
@pytest.mark.parametrize("doc_type,word", [("regulation", "ระเบียบ"), ("bylaw", "ข้อบังคับ")])
def test_regulation_articles_and_chapters(render_example, para, doc_type, word):
    doc = render_example(doc_type)
    data = load_example(doc_type)
    assert has(doc, word) and has(doc, f"ว่าด้วย{data['title']}")
    ts = texts(doc)
    idx = []
    for n, art in enumerate(data["articles"], 1):
        i = next(k for k, t in enumerate(ts) if t.startswith(f"ข้อ {_thai(n)}"))
        idx.append(i)
        if art["chapter"]:
            assert ts[i - 1] == art["chapter"]
            ch = doc.paragraphs[i - 1]
            assert jc(ch) == "center"
    assert idx == sorted(idx)
    other = "ข้อบังคับ" if doc_type == "regulation" else "ระเบียบ"
    assert not any(t.startswith(other) for t in ts)


def _thai(n: int) -> str:
    return str(n).translate(str.maketrans("0123456789", "๐๑๒๓๔๕๖๗๘๙"))


# ------------------------------------------------------------------ ประกาศ
def test_announcement_word_and_date(render_example, para):
    doc = render_example("announcement")
    assert has(doc, "ประกาศ") and "ประกาศ ณ วันที่" in flat_text(doc)
    doc = render_example("announcement", word="แจ้งความ")
    assert has(doc, "แจ้งความ") and not has(doc, "ประกาศ")
    assert "แจ้งความ ณ วันที่" in flat_text(doc)


# ------------------------------------------------------------------ แถลงการณ์ / ข่าว
@pytest.mark.parametrize("doc_type,word", [("statement", "แถลงการณ์"), ("news", "ข่าว")])
def test_statement_news_no_signer(render_example, para, doc_type, word):
    doc = render_example(doc_type, number="๓")
    assert has(doc, word)
    assert has(doc, "ฉบับที่ ๓")
    assert not has(doc, "(")
    assert "()" not in all_text(doc)
    other = "ข่าว" if doc_type == "statement" else "แถลงการณ์"
    assert not has(doc, other)


# ------------------------------------------------------------------ หนังสือรับรอง
def test_certificate_parts(render_example, para):
    doc = render_example("certificate")
    assert has(doc, "เลขที่")
    assert "ให้ไว้ ณ วันที่" in flat_text(doc)
    assert has(doc, "ลายมือชื่อผู้ได้รับการรับรอง")
    assert has(doc, "(นายตัวอย่าง สมมติ)")
    doc = render_example("certificate", photo_holder=None)
    assert not has(doc, "ลายมือชื่อผู้ได้รับการรับรอง")


# ------------------------------------------------------------------ รายงานการประชุม
def test_minutes_parts(render_example, para):
    doc = render_example("minutes")
    data = load_example("minutes")
    for label in ("ผู้มาประชุม", "ผู้ไม่มาประชุม", "เริ่มประชุมเวลา", "เลิกประชุมเวลา", "ผู้จดรายงานการประชุม"):
        assert has(doc, label), label
    ts = texts(doc)
    for item in data["agenda"]:
        i = next(k for k, t in enumerate(ts) if item["title"] in t)
        assert ts[i].startswith("ระเบียบวาระที่")
        assert any(t.startswith("มติที่ประชุม") and item["resolution"] in t for t in ts[i:])
    assert sum(t.startswith("มติที่ประชุม") for t in ts) == len(data["agenda"])
    assert "น. น." not in flat_text(doc)


# ------------------------------------------------------------------ เอกสารยาว
def test_order_large_committee_and_long_text(render_example):
    members = [{"name": f"สมาชิกลำดับที่{n}", "position": "ตำแหน่งตัวอย่าง", "role": "กรรมการ"} for n in range(1, 61)]
    members[-1]["name"] = "นายคนสุดท้าย ตัวอย่าง"
    paragraphs = ["ข้อความสั่งการตัวอย่างที่ยาวพอสมควร " * 12] * 18      # รวม > 6,000 อักขระ
    assert sum(map(len, paragraphs)) >= 6000
    doc = render_example("order", paragraphs=paragraphs,
                         committees=[{"title": "คณะทำงานขนาดใหญ่", "members": members}])
    cells = "\n".join(c.text for t in doc.tables for row in t.rows for c in row.cells)
    assert "นายคนสุดท้าย ตัวอย่าง" in cells and "๑.๖๐" in cells
    assert check(doc) == []


def test_external_long_subject_complete(render_example):
    subject = "การขอความอนุเคราะห์ข้อมูลประกอบการพิจารณาเรื่องตัวอย่าง" * 4
    subject = subject[:200]
    assert len(subject) == 200
    doc = render_example("external_letter", subject=subject)
    assert subject in all_text(doc).replace("\xa0", "").replace("​", "").replace("⁠", "")
