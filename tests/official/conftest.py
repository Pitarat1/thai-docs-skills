"""fixture ร่วมของเทสหนังสือราชการ

render_example(doc_type, *, spec=None, emblem=None, **overrides) -> docx.Document
  - โหลด examples/<doc_type>.json แล้วแก้ field ตาม overrides ก่อนตรวจโครงและสร้าง .docx
  - key ของ overrides ใช้จุดชี้ field ซ้อนได้ เช่น
        render_example("memo", **{"signer.name": "", "body.cause": ["ด้วย..."]})
    ค่า None ของ key ระดับบนสุดใดๆ ก็ส่งผ่านตามนั้น (เช่น fax=None)
  - spec: LayoutSpec ที่ใช้ (None = DEFAULT_SPEC); emblem: path ไฟล์ตรา (None = ไม่ส่ง)
  - ไม่ผ่าน fill ยกเว้นข้อความส่งสำเนาของหนังสือประทับตราที่มี forwarded แต่ไม่มี paragraphs
    (ตัวอย่าง stamp_letter เป็นแบบนี้ และคำสั่ง render เติมให้ผ่าน fill เสมอ)
  - คืน Document ที่บันทึกลงไฟล์จริงแล้วเปิดกลับด้วย python-docx
    doc.saved_path = path ไฟล์ .docx, doc.renderer = DocxRenderer ที่ใช้สร้าง
    (อ่าน doc.renderer.used_placeholder_emblem ได้)

para(doc, startswith) -> Paragraph | None
  - ย่อหน้าแรกในเนื้อเอกสาร (doc.paragraphs) ที่ข้อความ (ตัดแท็บ/ช่องว่างนำหน้า) ขึ้นต้นตาม startswith
"""
from __future__ import annotations

import copy

import pytest
from docx import Document

from thaidocs.official.fill import forward_paragraphs
from thaidocs.official.render import DocxRenderer
from thaidocs.official.schema import DocType, load_example, parse_doc
from thaidocs.official.style import DEFAULT_SPEC


def _set_path(data: dict, dotted: str, value) -> None:
    keys = dotted.split(".")
    cur = data
    for k in keys[:-1]:
        cur = cur.setdefault(k, {})
    cur[keys[-1]] = value


@pytest.fixture
def render_example(tmp_path):
    counter = {"n": 0}

    def _render(doc_type: str, *, spec=None, emblem=None, **overrides):
        data = copy.deepcopy(load_example(doc_type))
        for key, value in overrides.items():
            _set_path(data, key, value)
        model = parse_doc(data)
        if model.doc_type == DocType.STAMP_LETTER and model.forwarded and not model.paragraphs:
            model.paragraphs = forward_paragraphs(model.issuer, model.forwarded)
        renderer = DocxRenderer(spec or DEFAULT_SPEC, emblem=emblem)
        counter["n"] += 1
        path = tmp_path / f"{doc_type}-{counter['n']}.docx"
        renderer.render_to(model, path)
        out = Document(str(path))
        out.saved_path = path
        out.renderer = renderer
        return out

    return _render


def _para(doc, startswith: str):
    for p in doc.paragraphs:
        if p.text.lstrip("\t ").startswith(startswith):
            return p
    return None


@pytest.fixture
def para():
    return _para
