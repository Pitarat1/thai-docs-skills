"""fixture ร่วมของเทส

make_doc(paragraphs, *, table=False, header=None, footnote=False) -> docx.Document
  - paragraphs: list ของ str หรือ (text, style_name) เช่น ("บทนำ", "Heading 1")
  - table=True      เพิ่มตาราง 1x2 ที่มีข้อความไทยท้ายเอกสาร
  - header="..."    ใส่ข้อความในหัวกระดาษของ section แรก
  - footnote=True   เพิ่ม part เชิงอรรถ (footnotes.xml: separator id -1/0 + เชิงอรรถจริง id 1
                    ที่มีข้อความไทย) และ run footnoteReference ในย่อหน้าท้ายสุดของเนื้อหา
  - คืน Document ที่ "บันทึกแล้วเปิดใหม่" จากไฟล์จริง (tmp_path/"doc.docx") เพื่อให้ part
    ทุกตัว (รวมเชิงอรรถ ซึ่งโหลดกลับมาเป็น Part ธรรมดา) ผ่านเส้นทางเดียวกับไฟล์จริง
    path ของไฟล์อยู่ที่ doc.saved_path
"""
from __future__ import annotations

import pytest
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.opc.part import Part
from docx.oxml import parse_xml

FOOTNOTES_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"
_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
FOOTNOTE_TEXT = "เชิงอรรถภาษาไทยสำหรับทดสอบ"
TABLE_TEXTS = ("เซลล์ตารางภาษาไทย", "ข้อมูลในตาราง")

_FOOTNOTES_XML = (
    f'<w:footnotes xmlns:w="{_W}">'
    '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>'
    '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>'
    '<w:footnote w:id="1"><w:p><w:r><w:footnoteRef/></w:r>'
    f"<w:r><w:t>{FOOTNOTE_TEXT}</w:t></w:r></w:p></w:footnote>"
    "</w:footnotes>"
)


def _add_footnote(doc) -> None:
    part = Part(
        PackURI("/word/footnotes.xml"), FOOTNOTES_CT,
        ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + _FOOTNOTES_XML).encode("utf-8"),
        doc.part.package,
    )
    doc.part.relate_to(part, RT.FOOTNOTES)
    ref = parse_xml(f'<w:r xmlns:w="{_W}"><w:footnoteReference w:id="1"/></w:r>')
    body_ps = doc.element.body.findall(f"{{{_W}}}p")
    body_ps[-1].append(ref)


@pytest.fixture
def make_doc(tmp_path):
    def _make(paragraphs, *, table=False, header=None, footnote=False):
        doc = Document()
        for item in paragraphs:
            text, style = item if isinstance(item, tuple) else (item, None)
            if style and style not in [s.name for s in doc.styles]:
                doc.styles.add_style(style, WD_STYLE_TYPE.PARAGRAPH)  # เช่น "TOC 1" ที่ template ไม่มี
            doc.add_paragraph(text, style=style)
        if table:
            t = doc.add_table(rows=1, cols=2)
            for cell, text in zip(t.rows[0].cells, TABLE_TEXTS):
                cell.text = text
        if header is not None:
            doc.sections[0].header.paragraphs[0].text = header
        if footnote:
            if not paragraphs:
                doc.add_paragraph("ข้อความมีเชิงอรรถ")
            _add_footnote(doc)
        path = tmp_path / "doc.docx"
        doc.save(path)
        out = Document(str(path))
        out.saved_path = path
        return out

    return _make
