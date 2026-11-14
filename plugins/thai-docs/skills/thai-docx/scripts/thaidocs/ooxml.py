"""ตัวช่วย OOXML: ลำดับ element ตาม schema, การเข้าถึงทุก part ของเอกสาร

ที่มา: ส่วนที่ถอดจาก SARABUN.AI (Provincial Waterworks Authority / การประปาส่วนภูมิภาค,
Apache-2.0, engine/sarabun/render/thai_layout.py) และจาก skill
"thai-docx-layout-th" หมวด 9 (RPR_ORDER / PPR_ORDER / reorder)
Copyright 2026 การประปาส่วนภูมิภาค — ใช้ภายใต้ Apache License 2.0 ดัดแปลงให้ทำงานบน python-docx
ดู NOTICE ของ repo

API ที่ task อื่นใช้:
  RPR_ORDER, PPR_ORDER      ลำดับลูกของ w:rPr / w:pPr ตาม schema (ตัวที่ไม่รู้จักไปท้ายสุด)
  reorder(el, order)        เรียงลูกของ el ใหม่ตาม order (เสถียร, ไม่ลบ/เพิ่ม element)
  child(parent, tag)        หา/สร้างลูก (สร้างแล้วต่อท้าย: เรียก reorder เองหลังแก้เสร็จ)
  iter_parts(doc)           root (lxml) ของทุก part ที่มี: document, header*, footer*, footnotes,
                            endnotes, styles, numbering
  iter_content_parts(doc)   เฉพาะ part "เนื้อหา" ที่มีย่อหน้า: document, header*, footer*,
                            footnotes, endnotes (ไม่รวม styles/numbering)
  get_custom_property / set_custom_property / remove_custom_property
                            custom property ข้อความใน docProps/custom.xml (สร้าง/ลบ part ให้เอง)
  is_wrapped(doc)           thaidocs.wrapped == "1" (ใช้ร่วมกันโดย check และ wrap)
"""
from __future__ import annotations

from collections.abc import Iterator

from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.opc.part import XmlPart
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn
from lxml import etree

RPR_ORDER = [
    # CT_ParaRPr (rPr ของเครื่องหมายย่อหน้า) มี ins/del/moveFrom/moveTo นำหน้า
    "ins", "del", "moveFrom", "moveTo",
    "rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike", "outline",
    "shadow", "emboss", "imprint", "noProof", "snapToGrid", "vanish", "webHidden", "color", "spacing", "w",
    "kern", "position", "sz", "szCs", "highlight", "u", "effect", "bdr", "shd", "fitText", "vertAlign",
    "rtl", "cs", "em", "lang", "eastAsianLayout", "specVanish", "oMath", "rPrChange",
]
PPR_ORDER = [
    "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr",
    "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap",
    "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid",
    "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc", "textDirection",
    "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange",
]

_CONTENT_CT = {
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.macroEnabled.main+xml",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml",
}
_STYLE_CT = {
    "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
}


def local(el) -> str:
    return etree.QName(el).localname


def reorder(el, order: list[str]) -> None:
    """เรียงลูกของ el ตาม order (เสถียร: ตัวที่ไม่อยู่ใน order ไปท้ายตามลำดับเดิม)"""
    kids = list(el)
    key = [order.index(local(k)) if isinstance(k.tag, str) and local(k) in order else len(order) for k in kids]
    ordered = [k for _, _, k in sorted(zip(key, range(len(kids)), kids), key=lambda t: (t[0], t[1]))]
    if ordered == kids:
        return
    for k in kids:
        el.remove(k)
    for k in ordered:
        el.append(k)


def child(parent, tag: str):
    """คืนลูกชื่อ tag (เช่น "w:sz") ถ้าไม่มีก็สร้างต่อท้าย"""
    el = parent.find(qn(tag))
    if el is None:
        el = OxmlElement(tag)
        parent.append(el)
    return el


def _upgrade(rel) -> XmlPart:
    """part ที่ python-docx ไม่รู้จัก (เช่น footnotes) โหลดมาเป็น Part ธรรมดา (blob ดิบ):
    แปลงเป็น XmlPart ในที่เดิม เพื่อให้แก้ element ได้และถูกเขียนกลับตอน save"""
    part = rel.target_part
    if isinstance(part, XmlPart):
        return part
    new = XmlPart(part.partname, part.content_type, parse_xml(part.blob), part.package)
    for r in part.rels.values():  # คง relationship ของ part นั้น (เช่น ภาพในหัวกระดาษ)
        new.rels._rels[r.rId] = r
    rel._target = new
    return new


def _walk(doc, cts: set[str]) -> Iterator:
    main = doc.part
    seen: set[str] = set()
    for rel in list(main.rels.values()):
        if rel.is_external or rel.target_part.content_type not in cts:
            continue
        name = str(rel.target_part.partname)
        if name in seen:
            continue
        seen.add(name)
        yield _upgrade(rel)._element


def iter_content_parts(doc) -> Iterator:
    """root ของ part ที่มีย่อหน้า: document, header*/footer* (ทุก section ทุกชนิด รวมที่ไม่ได้ผูก
    ผ่าน doc.sections), footnotes, endnotes — document มาก่อนเสมอ"""
    yield doc.element
    yield from _walk(doc, _CONTENT_CT - {doc.part.content_type})


def iter_parts(doc) -> Iterator:
    """root ของทุก part ที่ใช้ปรับฟอนต์/ภาษา: iter_content_parts + styles + numbering"""
    yield from iter_content_parts(doc)
    yield from _walk(doc, _STYLE_CT)


# ---------- custom property (docProps/custom.xml) ----------
CUSTOM_CT = "application/vnd.openxmlformats-officedocument.custom-properties+xml"
CUSTOM_RT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties"
_CP = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
_VT = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
_FMTID = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"
WRAPPED_PROP = "thaidocs.wrapped"


def _custom_rel(doc):
    pkg = doc.part.package
    for rel in pkg.rels.values():
        if not rel.is_external and rel.reltype == CUSTOM_RT:
            return rel
    return None


def get_custom_property(doc, name: str) -> str | None:
    """ค่าข้อความของ custom property ชื่อ name (None = ไม่มี)"""
    rel = _custom_rel(doc)
    if rel is None:
        return None
    part = rel.target_part
    if isinstance(part, XmlPart):
        root = part._element
    else:
        try:
            root = etree.fromstring(part.blob)
        except etree.XMLSyntaxError:
            return None
    for prop in root:
        if prop.get("name") == name:
            return "".join(prop.itertext()).strip()
    return None


def set_custom_property(doc, name: str, value: str) -> None:
    """ตั้ง custom property แบบข้อความ (vt:lpwstr) สร้าง part docProps/custom.xml ถ้ายังไม่มี"""
    rel = _custom_rel(doc)
    if rel is None:
        pkg = doc.part.package
        root = parse_xml(f'<Properties xmlns="{_CP}" xmlns:vt="{_VT}"/>')
        part = XmlPart(PackURI("/docProps/custom.xml"), CUSTOM_CT, root, pkg)
        pkg.relate_to(part, CUSTOM_RT)
    else:
        root = _upgrade(rel)._element
    for prop in root:
        if prop.get("name") == name:
            for k in list(prop):
                prop.remove(k)
            break
    else:
        pids = [int(p.get("pid")) for p in root if (p.get("pid") or "").isdigit()]
        prop = etree.SubElement(root, f"{{{_CP}}}property")
        prop.set("fmtid", _FMTID)
        prop.set("pid", str(max([1, *pids]) + 1))
        prop.set("name", name)
    etree.SubElement(prop, f"{{{_VT}}}lpwstr").text = value


def remove_custom_property(doc, name: str) -> bool:
    """ลบ custom property; ถ้าไม่เหลือ property ใดเลยจะลบ part custom.xml ออกด้วย คืน True ถ้าลบ"""
    rel = _custom_rel(doc)
    if rel is None:
        return False
    root = _upgrade(rel)._element
    hit = [p for p in root if p.get("name") == name]
    for p in hit:
        root.remove(p)
    if len(root) == 0:
        del doc.part.package.rels[rel.rId]
    return bool(hit)


def is_wrapped(doc) -> bool:
    """custom property thaidocs.wrapped == "1" (เอกสารผ่าน wrap แล้ว)"""
    return get_custom_property(doc, WRAPPED_PROP) == "1"
