"""เลขไทย: แปลงเลขอารบิกเป็น ๐–๙ ใน w:t (ต่อย่อหน้า ความยาวเดิม) + เลขข้อ/เลขหน้า

ที่มา: skill "thai-docx-layout-th" หมวด 3 และ 9 ข้อ "เลขไทย" ปรับให้ทำงานบน docx.Document
กติกา: ห้ามแตะ w:instrText; คง URL/อีเมล/DOI และเลขที่ติดอักษรละติน (HC3, ISO 27001)
"""
from __future__ import annotations

import re

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from .ooxml import iter_content_parts, local

_TH = str.maketrans("0123456789", "๐๑๒๓๔๕๖๗๘๙")
# ส่วนที่คงเลขอารบิก
_NT = r"[^\s\u0e00-\u0e7f]"  # อักขระที่ไม่ใช่ช่องว่างและไม่ใช่อักษรไทย (URL หยุดที่อักษรไทย)
KEEP = re.compile(
    rf"https?://{_NT}+|www\.{_NT}+|{_NT}+@{_NT}+\.{_NT}+|doi:{_NT}+|10\.\d{{4,}}/{_NT}+"
    r"|ICC\(1\)|(?:ISO|IEC)(?:/IEC)?\s*\d[\d:.\-]*"
    r"|[A-Za-z][A-Za-z0-9]*(?:[-.:/][A-Za-z0-9]+)*"  # รหัสที่ขึ้นต้นด้วยอักษรละติน เช่น COVID-19, v2.0
)
_ARABIC = re.compile("[0-9]")

# CT_SectPr: element ที่อยู่ก่อน pgNumType
_SECT_BEFORE_PGNUM = {
    "headerReference", "footerReference", "footnotePr", "endnotePr", "type", "pgSz", "pgMar",
    "paperSrc", "pgBorders", "lnNumType",
}


def to_thai_digits(text: str) -> str:
    out, last = [], 0
    for m in KEEP.finditer(text):
        out.append(text[last:m.start()].translate(_TH))
        out.append(m.group())
        last = m.end()
    out.append(text[last:].translate(_TH))
    return "".join(out)


def apply_thai_digits(doc) -> int:
    """คืนจำนวน w:t ที่เปลี่ยนข้อความ (ไม่รวมเลขข้อ/เลขหน้า)"""
    n = 0
    for root in iter_content_parts(doc):
        for p in root.iter(qn("w:p")):
            ts = list(p.iter(qn("w:t")))  # ไม่รวม instrText
            full = "".join(t.text or "" for t in ts)
            if not _ARABIC.search(full):
                continue
            new, pos = to_thai_digits(full), 0
            for t in ts:
                k = len(t.text or "")
                if k and new[pos:pos + k] != t.text:
                    t.text = new[pos:pos + k]
                    n += 1
                pos += k
    numbering = doc.part.numbering_part.element if _has_numbering(doc) else None
    if numbering is not None:
        for f in numbering.iter(qn("w:numFmt")):
            if f.get(qn("w:val")) == "decimal":
                f.set(qn("w:val"), "thaiNumbers")
    for root in iter_content_parts(doc):
        for s in root.iter(qn("w:sectPr")):
            pg = s.find(qn("w:pgNumType"))
            if pg is None:
                pg = OxmlElement("w:pgNumType")
                pos = max([i + 1 for i, k in enumerate(s)
                           if isinstance(k.tag, str) and local(k) in _SECT_BEFORE_PGNUM] or [0])
                s.insert(pos, pg)
            pg.set(qn("w:fmt"), "thaiNumbers")
    return n


def _has_numbering(doc) -> bool:
    from docx.opc.constants import RELATIONSHIP_TYPE as RT
    return any(r.reltype == RT.NUMBERING for r in doc.part.rels.values())
