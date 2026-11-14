"""เติมส่วนตายตัวของร่างหนังสือราชการ (deterministic ไม่เรียก AI)

fill(doc, agency) คืนสำเนาที่:
  - แก้ถ้อยคำที่พบบ่อย: "เรื่อง"/"เรียน" ซ้ำหน้าค่า, วงเล็บรอบชื่อผู้ลงนาม, ย่อหน้าที่มี \\n,
    รายการ "วันที่: … / เวลา: …" -> ความเรียง
  - เติมข้อมูลหน่วยงานเฉพาะช่องที่ยังเป็นค่าเริ่มต้นหรือ [ระบุ…] (ค่าที่ร่างระบุแล้วไม่ถูกทับ)
  - หนังสือประทับตรา: เขียนข้อความส่งสำเนาหนังสือต้นเรื่องให้เมื่อร่างไม่มี paragraphs
ไม่แก้วัตถุที่รับเข้ามา
"""
from __future__ import annotations

import re
from typing import Optional

from ..digits import to_thai_digits
from . import salutations
from .profile import AgencyInfo
from .schema import (AnyDoc, CertificateDoc, ExternalLetterDoc, ForwardedLetter, MemoDoc, MinutesDoc,
                     StampLetterDoc)
from .textutil import is_placeholder, merge_list_lines, split_paragraphs, time_range

_SUBJECT_PREFIX = re.compile(r"^เรื่อง(?:\s*[:：]\s*|\s+)")
_TO_PREFIX = re.compile(r"^(?:กราบเรียน|เรียน)(?:\s*[:：]\s*|\s+)")
_SIGNER_EDGE = re.compile(r"^[\s(（]+|[\s)）]+$")
# คำขึ้นต้นภาคเหตุติดกับคำถัดไป: "ด้วย สำนักงาน…" -> "ด้วยสำนักงาน…" ("ในการนี้ …" ไม่แตะ)
_OPENING_SPACE = re.compile(r"^(ด้วย|เนื่องด้วย|เนื่องจาก|ตามที่)\s+")
# renderer ต่อ " น." หลังเวลาเริ่ม/เลิกประชุมเอง
_TIME_UNIT = re.compile(r"(?:\s*(?:น\.|นาฬิกา)|(?<=[0-9๐-๙])\s*น)\s*$")

_DEFAULT_SALUTATION = "เรียน"
_DEFAULT_CLOSING = "ขอแสดงความนับถือ"


def forward_paragraphs(issuer: str, letters: list[ForwardedLetter]) -> list[str]:
    """ข้อความตายตัวของหนังสือประทับตราที่ส่งสำเนาหนังสือต้นเรื่องเพื่อทราบ
    ๑–๒ ฉบับ: ประโยคเดียว เชื่อมด้วย "และสำเนา"; ๓ ฉบับขึ้นไป: "จำนวน n ฉบับ ดังนี้" แล้วแยกข้อ"""
    def ref(x: ForwardedLetter) -> str:
        speed = f" {x.speed}" if x.speed else ""
        return f"หนังสือ{x.agency}{speed} ที่ {x.doc_no} ลงวันที่ {x.date} เรื่อง {x.subject}"

    if len(letters) <= 2:
        refs = " และสำเนา".join(ref(x) for x in letters)
        return [f"{issuer} ขอส่งสำเนา{refs}", "จึงเรียนมาเพื่อทราบ"]
    items = [f"{to_thai_digits(str(i))}. {ref(x)}" for i, x in enumerate(letters, 1)]
    return [f"{issuer} ขอส่งสำเนาหนังสือ จำนวน {to_thai_digits(str(len(letters)))} ฉบับ ดังนี้",
            *items, "จึงเรียนมาเพื่อทราบ"]


def _prose(items: list[str], *, join_opening: bool = False) -> list[str]:
    out = merge_list_lines(split_paragraphs(items))
    if join_opening:
        out = [_OPENING_SPACE.sub(r"\1", p) for p in out]
    return out


def _fix_prose(doc: AnyDoc) -> None:
    body = getattr(doc, "body", None)
    if body is not None and hasattr(body, "cause"):
        body.cause = _prose(body.cause, join_opening=True)
        body.purpose = _prose(body.purpose)
        body.conclusion = time_range(body.conclusion)
    elif isinstance(body, list):                      # หนังสือรับรอง
        doc.body = _prose(body)
    if isinstance(getattr(doc, "paragraphs", None), list):
        doc.paragraphs = _prose(doc.paragraphs, join_opening=True)
    for sec in getattr(doc, "sections", None) or []:
        sec.paragraphs = _prose(sec.paragraphs)
    for c in getattr(doc, "committees", None) or []:
        c.paragraphs = _prose(c.paragraphs)
        c.duties = _prose(c.duties)
        for g in c.subgroups:
            g.duties = _prose(g.duties)
    for item in getattr(doc, "agenda", None) or []:
        item.discussion = _prose(item.discussion)


def _unset(value, default: str = "") -> bool:
    """ช่องที่ยังไม่ได้กรอก: ค่าเท่าค่าเริ่มต้นของโมเดล ว่าง หรือเป็น [ระบุ…]"""
    return value == default or is_placeholder(value)


def _unset_list(values: list[str]) -> bool:
    return not values or all(is_placeholder(v) for v in values)


def _fill_agency(doc: AnyDoc, a: AgencyInfo) -> None:
    if isinstance(doc, (ExternalLetterDoc, StampLetterDoc)):
        if _unset(doc.owner_unit) and a.department:
            doc.owner_unit = a.department
        if is_placeholder(doc.phone) and a.phone:
            doc.phone = a.phone
        if not doc.fax and a.fax:
            doc.fax = a.fax
        if not doc.email and a.email:
            doc.email = a.email
    if isinstance(doc, (ExternalLetterDoc, CertificateDoc)) and _unset_list(doc.agency) and a.address:
        doc.agency = list(a.address)
    if hasattr(doc, "issuer") and is_placeholder(doc.issuer) and a.org:
        doc.issuer = a.org
    if isinstance(doc, MemoDoc) and is_placeholder(doc.department) and a.department:
        doc.department = f"{a.department} โทร. {a.phone}" if a.phone else a.department
    signer = getattr(doc, "signer", None)
    if signer is not None:
        if is_placeholder(signer.name) and a.signer_name:
            signer.name = a.signer_name
        if _unset_list(signer.position) and a.signer_position:
            signer.position = list(a.signer_position)


def fill(doc: AnyDoc, agency: Optional[AgencyInfo] = None) -> AnyDoc:
    """คืนสำเนาที่เติมส่วนตายตัวแล้ว (ไม่แก้ doc เดิม)"""
    doc = doc.model_copy(deep=True)
    _fix_prose(doc)
    if isinstance(getattr(doc, "subject", None), str):
        doc.subject = _SUBJECT_PREFIX.sub("", doc.subject.strip()).strip()
    if isinstance(getattr(doc, "to", None), str):
        doc.to = _TO_PREFIX.sub("", doc.to.strip()).strip()
    signer = getattr(doc, "signer", None)
    if signer is not None:
        signer.name = _SIGNER_EDGE.sub("", signer.name)
    if isinstance(doc, MinutesDoc):
        doc.recorder.name = _SIGNER_EDGE.sub("", doc.recorder.name)
        doc.start_time = _TIME_UNIT.sub("", doc.start_time)
        doc.end_time = _TIME_UNIT.sub("", doc.end_time)
    if agency is not None:
        _fill_agency(doc, agency)
    if isinstance(doc, ExternalLetterDoc) and doc.salutation == _DEFAULT_SALUTATION \
            and doc.closing == _DEFAULT_CLOSING:
        doc.salutation, doc.closing = salutations.suggest(doc.to)
    if isinstance(doc, StampLetterDoc) and doc.forwarded and not doc.paragraphs:
        doc.paragraphs = forward_paragraphs(doc.issuer, doc.forwarded)
    return doc
