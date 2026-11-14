"""wrap / unwrap: ให้ Word for Mac ขึ้นบรรทัดภาษาไทยตรงขอบคำ

ผลทดลอง (spec หัวข้อ "ผลทดลองที่ใช้ออกแบบ", ไฟล์ v4_latin_zw): ปล่อย Word ตัดเองขาดกลางคำ 12.8%
แยกคำละ run ที่ "ไม่ติดป้ายไทย" (ไม่มี w:cs, lang bidi=en-US) + ZWSP ระหว่างคำไทยที่ติดกัน เหลือ 2.0%

wrap ทำต่อย่อหน้าที่มีอักษรไทย ในทุก part เนื้อหา (เนื้อความ ตาราง หัว/ท้ายกระดาษ เชิงอรรถ อ้างอิงท้ายเรื่อง):
  1. ลบ ZWSP เดิม (ทำซ้ำได้ผลเท่าเดิม)
  2. ต่อข้อความ run ทั้งย่อหน้า (run ลูกตรงของ w:p และใน hyperlink/smartTag/ins/customXml)
     ตัวกั้น (นับเป็นช่องว่าง ไม่แตะ): run ของ field ตั้งแต่ fldChar begin ถึง end (รวม instrText และผลของ field),
     fldSimple, sdt, สมการ และ element อื่นใน run ที่ไม่ใช่ w:t (tab, br, cr, sym, drawing, object,
     footnoteReference …) — run ที่มีทั้งข้อความและตัวกั้นถูกแยก run ที่ตัวกั้นก่อน
  3. seg.boundaries() ครั้งเดียวต่อย่อหน้า แล้วแยก run ที่จุดตัด (คัดลอก rPr และ attribute ของ run)
     ชิ้นที่ไม่มีอักษรไทยซึ่งอยู่ติดกันใน run เดียวกันรวมเป็น run เดียว (เช่น " HIS ") และคง rPr เดิม
     ข้อความไม่ข้ามขอบ hyperlink: run ในลิงก์อยู่ในลิงก์เดิม คำที่คร่อมขอบลิงก์ถูกแยกที่ขอบ
  4. run ที่มีอักษรไทย: ลบ w:cs, lang val=en-US bidi=en-US (คง eastAsia), rFonts 4 ช่องเท่ากัน,
     sz=szCs, ตัวหนา/เอียงที่มีแค่ bCs/iCs ได้ b/i ด้วย (ไม่มี w:cs แล้ว Word ใช้ b/i)
  5. ต่อท้าย ZWSP ให้ run ที่จบคำไทยเมื่อคำถัดไป (ในย่อหน้าเดียวกัน แม้คนละ run/ลิงก์) ขึ้นต้นด้วยอักษรไทย
     ไม่ใส่ติดละติน ตัวเลข ช่องว่าง ช่องว่างไม่ตัดบรรทัด และไม่ใส่ข้าม field/ตัวกั้น
  6. ตั้ง custom property thaidocs.wrapped=1
  w:lastRenderedPageBreak (แคชตำแหน่งขึ้นหน้าของ Word) ในย่อหน้าที่ wrap ถูกลบทิ้ง Word คำนวณใหม่เอง

unwrap: ลบ ZWSP, คืน w:cs + lang bidi=th-TH ให้ run ไทย (layout.thai_runs), รวม run ติดกันที่ attribute
และ rPr เหมือนกัน (เทียบ XML แบบ c14n) ซึ่งมีลูกเป็น w:t อย่างเดียว, ลบ property
"""
from __future__ import annotations

import copy
from dataclasses import dataclass, field

from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from lxml import etree

from .layout import FONT_ATTRS, THAI, thai_runs
from .ooxml import (
    RPR_ORDER,
    WRAPPED_PROP,
    child,
    is_wrapped,
    iter_content_parts,
    local,
    remove_custom_property,
    reorder,
    set_custom_property,
)
from .wordcut import _is_thai

__all__ = ["ZWSP", "is_wrapped", "wrap", "unwrap"]

ZWSP = "​"
_XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
# กล่องที่ run ข้างในต้องอยู่ในกล่องเดิม แต่ข้อความต่อกับย่อหน้าได้
_CONTAINERS = {"hyperlink", "smartTag", "ins", "customXml"}
# ไม่มีข้อความที่แสดง ไม่กั้นคำ (ข้ามไป)
_TRANSPARENT = {
    "pPr", "bookmarkStart", "bookmarkEnd", "proofErr", "commentRangeStart", "commentRangeEnd",
    "permStart", "permEnd", "del", "moveFrom", "moveFromRangeStart", "moveFromRangeEnd",
    "moveToRangeStart", "moveToRangeEnd",
}
_SEP = "\n"  # แทนตัวกั้นในข้อความที่ส่งให้ตัวตัดคำ


@dataclass
class _Chunk:
    """ส่วนของ run เดิม: ข้อความ (w:t ติดกัน) หรือ element อื่น (ตัวกั้นในบรรทัด)"""
    run: object
    elems: list
    text: str | None = None  # None = ตัวกั้น
    start: int = 0


@dataclass
class _Para:
    items: list = field(default_factory=list)  # _Chunk หรือ None (ตัวกั้นที่ไม่แตะ)
    runs: list = field(default_factory=list)   # run เดิมที่จะถูกแทน (ตามลำดับ)


def _collect(container, out: _Para, state: list[int]) -> None:
    """ไล่ลูกของ w:p / กล่อง ตามลำดับเอกสาร state[0] = ความลึกของ field"""
    for el in container:
        if not isinstance(el.tag, str):
            continue
        name = local(el)
        if name == "r":
            fld = el.findall(qn("w:fldChar"))
            if fld or state[0] > 0 or el.find(qn("w:instrText")) is not None:
                for f in fld:
                    kind = f.get(qn("w:fldCharType"))
                    if kind == "begin":
                        state[0] += 1
                    elif kind == "end":
                        state[0] = max(0, state[0] - 1)
                out.items.append(None)
                continue
            n_items = len(out.items)
            group: list = []
            for k in el:
                if not isinstance(k.tag, str) or local(k) in ("rPr", "lastRenderedPageBreak"):
                    continue
                is_t = local(k) == "t"
                if group and (local(group[-1]) == "t") == is_t:
                    group.append(k)
                    continue
                if group:
                    out.items.append(_mk_chunk(el, group))
                group = [k]
            if group:
                out.items.append(_mk_chunk(el, group))
            if len(out.items) > n_items:  # run ว่าง (มีแค่ rPr) ไม่แตะ
                out.runs.append(el)
        elif name in _CONTAINERS:
            _collect(el, out, state)
        elif name in _TRANSPARENT:
            continue
        else:
            out.items.append(None)


def _mk_chunk(run, elems) -> _Chunk:
    if local(elems[0]) == "t":
        return _Chunk(run, elems, "".join(t.text or "" for t in elems).replace(ZWSP, ""))
    return _Chunk(run, elems)


def _shell(run):
    """run ใหม่ที่มี attribute และ rPr (สำเนา) เหมือน run เดิม"""
    new = OxmlElement("w:r")
    for k, v in run.attrib.items():
        new.set(k, v)
    rpr = run.find(qn("w:rPr"))
    if rpr is not None:
        new.append(copy.deepcopy(rpr))
    return new


def _t(text: str):
    t = OxmlElement("w:t")
    if text != text.strip():
        t.set(_XML_SPACE, "preserve")
    t.text = text
    return t


def _latinize(run) -> None:
    """run ไทยแบบไม่ติดป้ายไทย: ลบ w:cs, lang en-US/en-US, rFonts 4 ช่องเท่ากัน, sz=szCs, bCs/iCs → b/i"""
    rpr = run.find(qn("w:rPr"))
    if rpr is None:
        rpr = OxmlElement("w:rPr")
        run.insert(0, rpr)
    for cs in rpr.findall(qn("w:cs")):
        rpr.remove(cs)
    lang = child(rpr, "w:lang")
    lang.set(qn("w:val"), "en-US")
    lang.set(qn("w:bidi"), "en-US")
    rf = rpr.find(qn("w:rFonts"))
    if rf is not None:
        font = next((rf.get(qn(a)) for a in FONT_ATTRS if rf.get(qn(a))), None)
        if font:
            for a in FONT_ATTRS:
                rf.set(qn(a), font)
    sz, szcs = rpr.find(qn("w:sz")), rpr.find(qn("w:szCs"))
    if szcs is not None and szcs.get(qn("w:val")):
        child(rpr, "w:sz").set(qn("w:val"), szcs.get(qn("w:val")))
    elif sz is not None and sz.get(qn("w:val")):
        child(rpr, "w:szCs").set(qn("w:val"), sz.get(qn("w:val")))
    for cs_tag, tag in (("w:bCs", "w:b"), ("w:iCs", "w:i")):
        src = rpr.find(qn(cs_tag))
        if src is not None and rpr.find(qn(tag)) is None:
            dst = OxmlElement(tag)
            if src.get(qn("w:val")) is not None:
                dst.set(qn("w:val"), src.get(qn("w:val")))
            rpr.append(dst)
    reorder(rpr, RPR_ORDER)


def _wrap_para(p, seg) -> bool:
    info = _Para()
    _collect(p, info, [0])
    parts, pos = [], 0
    for it in info.items:
        s = it.text if isinstance(it, _Chunk) and it.text is not None else _SEP
        if isinstance(it, _Chunk):
            it.start = pos
        parts.append(s)
        pos += len(s)
    text = "".join(parts)
    if not THAI.search(text):
        return False
    before = etree.tostring(p)
    bs = seg.boundaries(text)
    cuts = sorted(bs)

    def zw(end: int) -> bool:
        return end in bs and _is_thai(text[end - 1]) and _is_thai(text[end])

    new_runs: dict[int, list] = {}
    for it in info.items:
        if not isinstance(it, _Chunk):
            continue
        out = new_runs.setdefault(id(it.run), [])
        if it.text is None:
            r = _shell(it.run)
            for e in it.elems:
                r.append(e)
            out.append(r)
            continue
        a, b = it.start, it.start + len(it.text)
        edges = [a, *(c for c in cuts if a < c < b), b]
        pieces: list[list[int]] = []
        for x, y in zip(edges, edges[1:]):
            if pieces and not THAI.search(text[x:y]) and not THAI.search(text[pieces[-1][0]:pieces[-1][1]]):
                pieces[-1][1] = y
            else:
                pieces.append([x, y])
        if not pieces:  # w:t ว่าง: คงไว้หนึ่ง run
            pieces = [[a, b]]
        for x, y in pieces:
            r = _shell(it.run)
            word = text[x:y]
            r.append(_t(word + (ZWSP if y < len(text) and zw(y) else "")))
            if THAI.search(word):
                _latinize(r)
            out.append(r)
    for run in info.runs:
        parent = run.getparent()
        idx = parent.index(run)
        for i, r in enumerate(new_runs.get(id(run), [])):
            parent.insert(idx + i, r)
        parent.remove(run)
    return etree.tostring(p) != before


def wrap(doc, seg) -> int:
    """แยกคำละ run ไม่ติดป้ายไทย + ZWSP ระหว่างคำไทย ในทุก part เนื้อหา ตั้ง thaidocs.wrapped=1
    คืนจำนวนย่อหน้าที่แก้ (0 เมื่อทำซ้ำกับเอกสารที่ wrap แล้ว)"""
    n = 0
    for root in iter_content_parts(doc):
        for p in list(root.iter(qn("w:p"))):
            n += _wrap_para(p, seg)
    set_custom_property(doc, WRAPPED_PROP, "1")
    return n


def _rpr_key(run) -> bytes:
    rpr = run.find(qn("w:rPr"))
    if rpr is None or (len(rpr) == 0 and not rpr.attrib):
        return b""
    return etree.tostring(rpr, method="c14n")


def _text_only(run) -> bool:
    kids = [k for k in run if isinstance(k.tag, str) and local(k) != "rPr"]
    return bool(kids) and all(local(k) == "t" for k in kids)


def _merge_runs(parent) -> None:
    prev = None
    for el in list(parent):
        if not isinstance(el.tag, str) or local(el) != "r" or not _text_only(el):
            prev = None
            continue
        if (prev is not None and el.getprevious() is prev and dict(el.attrib) == dict(prev.attrib)
                and _rpr_key(el) == _rpr_key(prev)):
            ts = prev.findall(qn("w:t")) + el.findall(qn("w:t"))
            text = "".join(t.text or "" for t in ts)
            for t in ts[1:]:
                t.getparent().remove(t)
            ts[0].text = text
            if text != text.strip():
                ts[0].set(_XML_SPACE, "preserve")
            elif ts[0].get(_XML_SPACE) is not None:
                del ts[0].attrib[_XML_SPACE]
            parent.remove(el)
            continue
        prev = el


def unwrap(doc) -> int:
    """คืนสภาพก่อน wrap: ลบ ZWSP, คืน w:cs + lang bidi=th-TH ให้ run ไทย, รวม run ที่ rPr เหมือนกัน,
    ลบ property thaidocs.wrapped คืนจำนวนย่อหน้าที่แก้"""
    roots = list(iter_content_parts(doc))
    paras = [p for root in roots for p in root.iter(qn("w:p"))]
    before = [etree.tostring(p) for p in paras]
    for root in roots:
        for t in root.iter(qn("w:t")):
            if t.text and ZWSP in t.text:
                t.text = t.text.replace(ZWSP, "")
    thai_runs(doc)
    for root in roots:
        for el in list(root.iter()):
            if isinstance(el.tag, str) and (el.tag == qn("w:p") or local(el) in _CONTAINERS):
                _merge_runs(el)
    remove_custom_property(doc, WRAPPED_PROP)
    return sum(etree.tostring(p) != b for p, b in zip(paras, before))
