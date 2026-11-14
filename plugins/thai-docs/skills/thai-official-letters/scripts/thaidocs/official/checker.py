"""ตรวจหนังสือราชการ (rule-based, deterministic ไม่เรียก AI)

check_document(doc, spec) ตรวจโครงสร้าง/ถ้อยคำของร่างก่อนสร้างไฟล์ ข้อที่อ้างระเบียบสำนักนายกรัฐมนตรี
                          ว่าด้วยงานสารบรรณ พ.ศ. ๒๕๒๖ และที่แก้ไขเพิ่มเติมมี `ref` เป็นเลขข้อ
check_docx(path, spec)    ตรวจไฟล์ .docx: ขอบ ฟอนต์ ขนาด ตรา หัวบันทึกข้อความ ช่อง [ระบุ…] ที่เหลือ
                          (ref "ว ๒๐๑๙" = แนวทางการพิมพ์หนังสือราชการภาษาไทยด้วยคอมพิวเตอร์ ของสำนักนายกรัฐมนตรี)

ระดับ: error = ขัดกับระเบียบ/แนวทางที่อ้างใน ref; warning/info = ข้อสังเกต
ธรรมเนียมการเขียนที่ไม่มีระเบียบรองรับเป็น warning/info เท่านั้น และไม่มี ref

ALL_CODES = รหัสทุกตัวที่โมดูลนี้ออกได้ (เทสตรวจว่าตรงกับซอร์ส; เอกสารผู้ใช้ใช้ตัวนี้เป็นรายการอ้างอิง)

ดัดแปลงจากโครงการต้นทางภายใต้ Apache License 2.0 ดู NOTICE ของ repo
"""
from __future__ import annotations

import re
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Literal

from . import salutations
from .fill import forward_paragraphs
from .schema import (DOC_TYPE_REF, AnnouncementDoc, AnyDoc, CertificateDoc, DocType, ExternalLetterDoc,
                     MemoDoc, MinutesDoc, OrderDoc, RegulationDoc, StampLetterDoc)
from .style import DEFAULT_SPEC, NATIONAL_FONTS, LayoutSpec
from .textutil import (LIST_LINE_RE, THAI_MONTHS, find_placeholders, has_arabic_digits, strip_kept,
                       to_arabic_digits)

Level = Literal["error", "warning", "info"]

ALL_CODES: tuple[str, ...] = (
    # ร่าง (check_document)
    "PLACEHOLDER", "INFORMAL_WORD", "DATE_FORMAT", "DATE_BE", "DATE_DAY_BLANK", "SUBJECT_LONG",
    "SIGNER_PAREN", "DOC_NO", "MEMO_PROPOSAL", "MEMO_PHONE", "BODY_CAUSE", "BODY_PURPOSE",
    "BODY_CONCLUSION", "BODY_OPENING", "SALUTATION_PAIR", "SALUTATION_RANK", "REF_LATEST", "OWNER_UNIT",
    "STAMP_NOTIFY_ONLY", "STAMP_CONCLUSION", "STAMP_SCOPE", "STAMP_FORWARDED_IGNORED", "ORDER_NO", "ORDER_EFFECTIVE",
    "REG_ARTICLES", "REG_NAME", "REG_WORD", "REG_EFFECTIVE", "REG_KEEPER", "REG_EDITION", "BYLAW_AUTHORITY",
    "ANN_BODY", "CERT_OPENING", "MIN_AGENDA", "MIN_RESOLUTION",
    # ถ้อยคำ
    "OPENING_NEW_NAN", "OPENING_REF_NAN", "OPENING_REF_PARTS", "CONCLUSION_RANK", "PARA_LONG", "WORDY",
    # คุณภาพข้อความ (ทั้งร่างและไฟล์)
    "TEXT_PUA", "TEXT_SARA_AM", "TEXT_DUPLICATE", "TEXT_LIST_LINE", "TEXT_MIXED_DIGITS", "TEXT_BRACKET",
    # ไฟล์ .docx (check_docx)
    "MARGIN", "PAPER", "FONT", "FONT_SIZE", "ARABIC_DIGITS", "GARUDA", "MEMO_TITLE", "MEMO_LABEL",
)

_CUSTOM = " (ธรรมเนียมการเขียน ไม่ใช่ข้อกำหนดของระเบียบ)"


@dataclass
class Finding:
    level: Level
    code: str
    message: str
    where: str = ""
    ref: str = ""          # เลขข้อของระเบียบ / แนวทางที่อ้าง (ว่างเมื่อเป็นธรรมเนียมการเขียน)

    def dict(self):
        return asdict(self)


# ------------------------------------------------------------------ ข้อของระเบียบแยกตามชนิดหนังสือ
_DATE_REF = {
    DocType.EXTERNAL_LETTER: "ข้อ ๑๑.๓", DocType.MEMO: "ข้อ ๑๒.๓", DocType.STAMP_LETTER: "ข้อ ๑๔.๖",
    DocType.ORDER: "ข้อ ๑๖.๕", DocType.REGULATION: "ข้อ ๑๗.๗", DocType.ANNOUNCEMENT: "ข้อ ๒๐.๔",
    DocType.STATEMENT: "ข้อ ๒๑.๖", DocType.CERTIFICATE: "ข้อ ๒๔.๔",
}
_SUBJECT_REF = {
    DocType.EXTERNAL_LETTER: "ข้อ ๑๑.๔", DocType.MEMO: "ข้อ ๑๒.๔", DocType.ORDER: "ข้อ ๑๖.๓",
    DocType.ANNOUNCEMENT: "ข้อ ๒๐.๒", DocType.STATEMENT: "ข้อ ๒๑.๒",
}
_DOC_NO_REF = {DocType.EXTERNAL_LETTER: "ข้อ ๑๑.๑", DocType.MEMO: "ข้อ ๑๒.๒", DocType.STAMP_LETTER: "ข้อ ๑๔.๑"}
_BODY_REF = {DocType.EXTERNAL_LETTER: "ข้อ ๑๑.๘", DocType.MEMO: "ข้อ ๑๒.๖"}
_GUIDE = "ว ๒๐๑๙"


def _ref(table: dict, doc) -> str:
    return table.get(doc.doc_type, DOC_TYPE_REF[doc.doc_type])


_NUM = "[0-9๐-๙]"
DATE_RE = re.compile(rf"^\s*({_NUM}{{1,2}})?\s*({'|'.join(THAI_MONTHS)})\s+(?:พ\.ศ\.\s*)?({_NUM}{{4}})\s*$")
YEAR_NO_RE = re.compile(rf"^.+/\s*{_NUM}{{4}}$")           # เลขที่/ปีพุทธศักราช

# ---- ถ้อยคำ (ธรรมเนียมการเขียนหนังสือราชการ ไม่ใช่ข้อบังคับ)
_NEW_OPENINGS = ("ด้วย", "เนื่องจาก", "เนื่องด้วย")      # เรื่องใหม่: ไม่มี "นั้น" ท้ายตอน
_REF_OPENINGS = ("ตาม", "อนุสนธิ")                      # เรื่องเดิม: ต้องมี "นั้น" + อย่างน้อย ๒ ตอนก่อน "จึง"
_WORDY = ("กระทำการ", "ทำการ")                          # ฟุ่มเฟือยเมื่อนำหน้ากริยาอื่น
_STAMP_ACTION = ("ขอให้", "โปรดดำเนินการ", "ภายในวันที่", "ขอความร่วมมือ", "โปรดส่ง", "เข้าร่วม", "โปรดพิจารณา")
_PARA_MAX_CHARS = 900                                    # ≈ ๑๐ บรรทัด ที่ ๑๖ พอยท์

INFORMAL = {"ครับ": "ภาษาพูด", "ค่ะ": "ภาษาพูด", "นะคะ": "ภาษาพูด", "เยอะ": "ใช้ 'จำนวนมาก'",
            "ได้มั้ย": "ภาษาพูด", "อยากจะ": "ใช้ 'มีความประสงค์จะ'", "ช่วย": "ใช้ 'ขอความอนุเคราะห์'",
            "เลยขอ": "ภาษาพูด"}


def _check_writing(f: list[Finding], body, salutation: str = "เรียน"):
    """หลักการใช้คำขึ้นต้นเนื้อเรื่องและภาคสรุป (ธรรมเนียมการเขียน)"""
    first = (body.cause[0] if body.cause else "").strip()
    if first.startswith(_NEW_OPENINGS) and first.endswith("นั้น"):
        f.append(Finding("warning", "OPENING_NEW_NAN",
                         "ขึ้นต้น 'ด้วย/เนื่องจาก' (เรื่องใหม่) ไม่ต้องลงท้ายตอนด้วย 'นั้น'" + _CUSTOM, "body.cause[0]"))
    if first.startswith(_REF_OPENINGS):
        if not first.endswith("นั้น"):
            f.append(Finding("warning", "OPENING_REF_NAN",
                             "ขึ้นต้น 'ตาม/ตามที่/อนุสนธิ' (อ้างเรื่องเดิม) ต้องลงท้ายตอนแรกด้วย 'นั้น'" + _CUSTOM,
                             "body.cause[0]"))
        if len(body.cause) + len(body.purpose) < 2:
            f.append(Finding("warning", "OPENING_REF_PARTS",
                             "อ้างเรื่องเดิมแล้วควรมีเรื่องต่อเนื่องอีกอย่างน้อย ๑ ตอนก่อนย่อหน้า 'จึง…' "
                             "(เช่น 'ในการนี้ …')" + _CUSTOM, "body"))
    if salutation == "กราบเรียน" and body.conclusion.strip().startswith("จึงเรียน"):
        f.append(Finding("warning", "CONCLUSION_RANK",
                         "ผู้รับใช้ 'กราบเรียน' ภาคสรุปควรเป็น 'จึงกราบเรียนมาเพื่อโปรด…'" + _CUSTOM, "body.conclusion"))
    for i, t in enumerate(body.paragraphs()):
        if len(t) > _PARA_MAX_CHARS:
            f.append(Finding("info", "PARA_LONG", "ตอนหนึ่งไม่ควรเกิน ๑๐ บรรทัด — แยกประเด็นเป็นย่อหน้าใหม่" + _CUSTOM,
                             f"ย่อหน้า {i + 1}"))
        for w in _WORDY:
            if w in t:
                f.append(Finding("info", "WORDY", f"'{w}' นำหน้ากริยามักฟุ่มเฟือย ตัดออกได้" + _CUSTOM, f"ย่อหน้า {i + 1}"))
                break


# ---- คุณภาพข้อความ (พบบ่อยในหนังสือที่คัดลอกจาก PDF/ไฟล์เก่า หรือแก้ซ้ำหลายรอบ)
_PUA_RE = re.compile("[-]")                 # อักษรไทยตำแหน่งพิเศษของฟอนต์เก่า (วรรณยุกต์/สระลอย)
_SPLIT_AM_RE = re.compile("ํา")                # "ํ" + "า" แทน "ำ" -> ค้นคำไม่เจอ
_DUP_RE = re.compile(r"([ก-ฮเ-ไ][ก-๏]{3,})\1")               # คำซ้ำติดกัน เช่น "แต่งตั้งแต่งตั้ง"
_MIXED_DIGIT_RE = re.compile(r"[0-9][๐-๙]|[๐-๙][0-9]")     # เลขไทยปนอารบิกในตัวเลขเดียว เช่น "๐๐3๓"
_PAIRS = (("(", ")"), ("[", "]"), ("“", "”"))
_LIST_MARK_RE = re.compile(r"^\s*[0-9๐-๙]+\)")            # "1) ..." เลขข้อย่อย ไม่นับเป็นวงเล็บ


def text_quality(text: str, where: str = "") -> list[Finding]:
    """ตรวจข้อความที่มักผิดจากการคัดลอก/แก้ไข — ใช้ได้ทั้งร่างและไฟล์ .docx"""
    f: list[Finding] = []
    if not text:
        return f
    if _PUA_RE.search(text):
        f.append(Finding("warning", "TEXT_PUA", "พบอักษรไทยรหัสพิเศษของฟอนต์เก่า (วรรณยุกต์/สระอาจหายเมื่อเปลี่ยนฟอนต์) "
                         "— พิมพ์คำนั้นใหม่ด้วยฟอนต์แห่งชาติ", where))
    if _SPLIT_AM_RE.search(text):
        f.append(Finding("warning", "TEXT_SARA_AM", "พบ 'ํ' + 'า' แทนสระ 'ำ' — ค้นหาคำไม่เจอ ให้พิมพ์ 'ำ' ใหม่", where))
    for m in _DUP_RE.finditer(text):
        f.append(Finding("warning", "TEXT_DUPLICATE", f"พบคำซ้ำติดกัน '{m.group(0)}'", where))
    if LIST_LINE_RE.match(text):
        f.append(Finding("warning", "TEXT_LIST_LINE",
                         "แจกแจงแบบ 'หัวข้อ: ค่า' — วัน เวลา สถานที่ ควรเขียนต่อท้ายประโยคในย่อหน้าเดียว "
                         "เช่น 'ในวันที่ … เวลา … ณ …'" + _CUSTOM, where))
    if _MIXED_DIGIT_RE.search(strip_kept(text)):          # รหัส/URL ที่แกนคงเลขอารบิกไว้ ไม่นับ
        f.append(Finding("warning", "TEXT_MIXED_DIGITS", "พบเลขไทยปนเลขอารบิกในตัวเลขเดียวกัน", where))
    body = _LIST_MARK_RE.sub("", text)
    for a, b in _PAIRS:
        if body.count(a) != body.count(b):
            f.append(Finding("warning", "TEXT_BRACKET", f"เครื่องหมาย {a} {b} ไม่ครบคู่", where))
    return f


def _texts(doc: AnyDoc) -> dict[str, str]:
    out: dict[str, str] = {}

    def walk(prefix, v):
        if isinstance(v, str):
            out[prefix] = v
        elif isinstance(v, list):
            for i, x in enumerate(v):
                walk(f"{prefix}[{i}]", x)
        elif isinstance(v, dict):
            for k, x in v.items():
                walk(f"{prefix}.{k}" if prefix else k, x)
    walk("", doc.model_dump(mode="json"))
    return out


def _check_date(f: list[Finding], value: str, ref: str, where: str = "date"):
    if not value or find_placeholders(value):
        return
    m = DATE_RE.match(value)
    if not m:
        f.append(Finding("error", "DATE_FORMAT", "วันที่ต้องเป็น 'ตัวเลขวันที่ ชื่อเต็มของเดือน ตัวเลขปีพุทธศักราช' "
                         "เช่น ๓ ตุลาคม ๒๕๖๙", where, ref))
    elif int(to_arabic_digits(m.group(3))) < 2400:
        f.append(Finding("error", "DATE_BE", "ปีต้องเป็นพุทธศักราช", where, ref))
    elif not m.group(1):
        f.append(Finding("info", "DATE_DAY_BLANK", "เว้นวันที่ไว้ให้ลงเมื่อลงนาม/ออกเลข", where, ref))


def check_document(doc: AnyDoc, spec: LayoutSpec = DEFAULT_SPEC) -> list[Finding]:
    f: list[Finding] = []
    texts = _texts(doc)
    if doc.doc_type.value in spec.omit_signature:          # ผู้ลงนามไม่อยู่ในเอกสาร — ไม่ตรวจช่องผู้ลงนาม
        texts = {k: v for k, v in texts.items() if not k.startswith("signer")}

    for where, t in texts.items():
        f += text_quality(t, where)
        for ph in find_placeholders(t):
            f.append(Finding("warning", "PLACEHOLDER", f"ยังไม่ได้กรอก {ph}", where))
        for w, why in INFORMAL.items():
            if re.search(rf"(?<![ก-๙]){w}(?![ก-๙])", t):
                f.append(Finding("warning", "INFORMAL_WORD", f"พบคำ '{w}' — {why}" + _CUSTOM, where))

    if not isinstance(doc, MinutesDoc):
        _check_date(f, getattr(doc, "date", ""), _ref(_DATE_REF, doc))

    subj = getattr(doc, "subject", "") or ""
    if len(subj) > 120:
        f.append(Finding("warning", "SUBJECT_LONG", "เรื่องให้ลงเรื่องย่อที่เป็นใจความสั้นที่สุด", "subject",
                         _ref(_SUBJECT_REF, doc)))

    signer = getattr(doc, "signer", None)
    if signer and signer.name.strip().startswith("(") and doc.doc_type.value not in spec.omit_signature:
        f.append(Finding("info", "SIGNER_PAREN", "ไม่ต้องใส่วงเล็บ ระบบพิมพ์ชื่อเต็มในวงเล็บให้", "signer.name", "ข้อ ๑๑.๑๐"))

    # ---- หนังสือโต้ตอบ: ที่ = รหัส/เลขทะเบียนส่ง (ข้อ ๑๑.๑, ๑๒.๒, ๑๔.๑) — ช่อง [ระบุ…] รายงานเป็น PLACEHOLDER แล้ว
    if isinstance(doc, (MemoDoc, ExternalLetterDoc, StampLetterDoc)) and "/" not in doc.doc_no \
            and not find_placeholders(doc.doc_no):
        f.append(Finding("error", "DOC_NO", "ที่ ต้องเป็น 'รหัสตัวพยัญชนะและเลขประจำของเจ้าของเรื่อง/เลขทะเบียนหนังสือส่ง'",
                         "doc_no", _ref(_DOC_NO_REF, doc)))

    body = getattr(doc, "body", None)
    if isinstance(doc, MemoDoc) and body is None:
        if not any("ข้อเสนอ" in x.heading for x in doc.sections):
            f.append(Finding("info", "MEMO_PROPOSAL", "บันทึกเสนอเรื่องแบบมีหัวข้อ ควรจบด้วยหัวข้อ 'ข้อเสนอ'" + _CUSTOM,
                             "sections"))
    elif isinstance(doc, (MemoDoc, ExternalLetterDoc)):
        ref = _ref(_BODY_REF, doc)
        if not body.cause:
            f.append(Finding("error", "BODY_CAUSE", "ไม่มีภาคเหตุ (ข้อความควรเริ่มด้วยที่มาหรือเหตุผลของเรื่อง)",
                             "body.cause", ref))
        if not body.purpose:
            f.append(Finding("error", "BODY_PURPOSE", "ไม่มีภาคความประสงค์ (สิ่งที่ต้องการให้ผู้รับดำเนินการ)",
                             "body.purpose", ref))
        if not body.conclusion.strip().startswith(("จึง", "ทั้งนี้", "อนึ่ง")):
            f.append(Finding("info", "BODY_CONCLUSION", "ภาคสรุปมักขึ้นต้น 'จึงเรียนมาเพื่อ...'" + _CUSTOM,
                             "body.conclusion"))
        if body.cause and not body.cause[0].startswith(("ด้วย", "ตาม", "เนื่องด้วย", "เนื่องจาก", "อนุสนธิ")):
            f.append(Finding("info", "BODY_OPENING", "ภาคเหตุมักขึ้นต้น 'ด้วย' / 'ตาม' / 'เนื่องด้วย'" + _CUSTOM,
                             "body.cause[0]"))
        _check_writing(f, body, doc.salutation)

    if isinstance(doc, ExternalLetterDoc):
        known = doc.salutation in salutations.VALID_PAIRS
        if not salutations.pair_ok(doc.salutation, doc.closing):
            if known:
                f.append(Finding("error", "SALUTATION_PAIR",
                                 f"คำขึ้นต้น '{doc.salutation}' ไม่คู่กับคำลงท้าย '{doc.closing}'", "closing",
                                 "ข้อ ๑๑.๕, ๑๑.๙ / ภาคผนวก ๒"))
            else:
                f.append(Finding("warning", "SALUTATION_PAIR",
                                 f"คำขึ้นต้น '{doc.salutation}' ไม่อยู่ในตารางของเครื่องมือ จึงตรวจคู่กับคำลงท้าย "
                                 f"'{doc.closing}' ไม่ได้ — ตรวจกับภาคผนวก ๒", "salutation", "ภาคผนวก ๒"))
        exp_sal, _ = salutations.suggest(doc.to)
        if known and exp_sal != doc.salutation:
            f.append(Finding("warning", "SALUTATION_RANK", f"ผู้รับ '{doc.to}' ควรใช้คำขึ้นต้น '{exp_sal}'",
                             "salutation", "ภาคผนวก ๒"))
        if len(doc.references) > 1:
            f.append(Finding("info", "REF_LATEST", "อ้างถึงให้อ้างหนังสือที่ติดต่อกันมาครั้งหลังสุดเพียงฉบับเดียว "
                             "เว้นแต่มีสาระที่ต้องพิจารณาต่อเนื่อง", "references", "ข้อ ๑๑.๖"))
        if not doc.owner_unit or not doc.phone:
            f.append(Finding("error", "OWNER_UNIT", "ต้องมีส่วนราชการเจ้าของเรื่องและโทร.", "owner_unit",
                             "ข้อ ๑๑.๑๒–๑๑.๑๓"))

    if isinstance(doc, MemoDoc) and "โทร" not in doc.department:
        f.append(Finding("info", "MEMO_PHONE", "ส่วนราชการ ให้ลงส่วนราชการเจ้าของเรื่องพร้อมหมายเลขโทรศัพท์ (ถ้ามี)",
                         "department", "ข้อ ๑๒.๑"))

    if isinstance(doc, StampLetterDoc) and spec.stamp_notify_only:
        text = " ".join(doc.paragraphs)
        hits = [w for w in _STAMP_ACTION if w in text]
        if hits:
            f.append(Finding("warning", "STAMP_NOTIFY_ONLY",
                             f"หนังสือประทับตราตั้งไว้ว่าใช้แจ้งเพื่อทราบเท่านั้น (ค่า stamp_notify_only) แต่พบ '{hits[0]}' "
                             "— เรื่องที่ผู้รับต้องดำเนินการควรใช้บันทึกข้อความหรือหนังสือภายนอก" + _CUSTOM, "paragraphs"))
        if doc.paragraphs and not doc.paragraphs[-1].strip().startswith("จึงเรียนมาเพื่อทราบ"):
            f.append(Finding("info", "STAMP_CONCLUSION", "หนังสือประทับตราที่แจ้งเพื่อทราบปิดด้วย 'จึงเรียนมาเพื่อทราบ'"
                             + _CUSTOM, "paragraphs"))

    if isinstance(doc, StampLetterDoc) and doc.forwarded and doc.paragraphs \
            and doc.paragraphs != forward_paragraphs(doc.issuer, doc.forwarded):
        f.append(Finding("warning", "STAMP_FORWARDED_IGNORED",
                         "ร่างมีทั้ง forwarded และ paragraphs ระบบใช้ paragraphs และไม่พิมพ์รายการหนังสือใน forwarded "
                         "— ลบ paragraphs เพื่อให้ระบบเขียนข้อความส่งสำเนาเอง หรือเขียนชื่อหนังสือต้นเรื่องใน paragraphs "
                         "แล้วลบ forwarded", "forwarded"))

    if isinstance(doc, StampLetterDoc) and doc.speed is None and len(" ".join(doc.paragraphs)) > 1200:
        f.append(Finding("info", "STAMP_SCOPE", "หนังสือประทับตราใช้เฉพาะเรื่องที่ไม่ใช่เรื่องสำคัญ — ตรวจว่าเรื่องนี้เข้าข้อ ๑๓",
                         "paragraphs", "ข้อ ๑๓"))

    # ---- หนังสือสั่งการ
    if isinstance(doc, OrderDoc):
        if not YEAR_NO_RE.match(doc.doc_no) and not find_placeholders(doc.doc_no):
            f.append(Finding("error", "ORDER_NO", "ที่ ของคำสั่งต้องเป็น 'เลขที่/ปีพุทธศักราช' (เริ่ม ๑ ทุกปีปฏิทิน)",
                             "doc_no", "ข้อ ๑๖.๒"))
        if not doc.effective:
            f.append(Finding("warning", "ORDER_EFFECTIVE", "ระบุวันใช้บังคับ", "effective", "ข้อ ๑๖.๔"))

    if isinstance(doc, RegulationDoc):
        is_reg = doc.doc_type == DocType.REGULATION
        word = "ระเบียบ" if is_reg else "ข้อบังคับ"
        ref = "ข้อ ๑๗.๖" if is_reg else "ข้อ ๑๘"
        arts = doc.articles
        if len(arts) < 3:
            f.append(Finding("error", "REG_ARTICLES", "ต้องมีอย่างน้อย ข้อ ๑ ชื่อ, ข้อ ๒ วันใช้บังคับ, ข้อสุดท้าย ผู้รักษาการ",
                             "articles", ref))
        else:
            if "เรียกว่า" not in arts[0].text:
                f.append(Finding("error", "REG_NAME", f"ข้อ ๑ ต้องเป็นชื่อ{word} ('{word}นี้เรียกว่า ...')",
                                 "articles[0]", ref))
            elif word not in arts[0].text:
                f.append(Finding("warning", "REG_WORD", f"ชื่อในข้อ ๑ ไม่ตรงชนิดเอกสาร ({word})", "articles[0]", ref))
            if "ใช้บังคับ" not in arts[1].text:
                f.append(Finding("error", "REG_EFFECTIVE", "ข้อ ๒ ต้องเป็นวันใช้บังคับ", "articles[1]", ref))
            if "รักษาการ" not in arts[-1].text:
                f.append(Finding("error", "REG_KEEPER", "ข้อสุดท้ายต้องเป็นผู้รักษาการ", f"articles[{len(arts) - 1}]", ref))
        if doc.edition and to_arabic_digits(doc.edition).strip() == "1":
            f.append(Finding("info", "REG_EDITION", "ฉบับแรกไม่ต้องระบุฉบับที่", "edition", "ข้อ ๑๗.๓"))
        if doc.doc_type == DocType.BYLAW and "อาศัยอำนาจ" not in doc.preamble:
            f.append(Finding("error", "BYLAW_AUTHORITY", "ข้อบังคับต้องอ้างกฎหมายที่ให้อำนาจ", "preamble", "ข้อ ๑๘"))

    if isinstance(doc, AnnouncementDoc) and not doc.paragraphs:
        f.append(Finding("error", "ANN_BODY", "ต้องมีข้อความ (เหตุผล + สิ่งที่ประกาศ)", "paragraphs", "ข้อ ๒๐.๓"))

    # ---- หนังสือที่เจ้าหน้าที่ทำขึ้น
    if isinstance(doc, CertificateDoc):
        if not doc.body or not doc.body[0].startswith("หนังสือฉบับนี้ให้ไว้เพื่อรับรองว่า"):
            f.append(Finding("error", "CERT_OPENING", "ข้อความต้องขึ้นต้นว่า 'หนังสือฉบับนี้ให้ไว้เพื่อรับรองว่า'",
                             "body[0]", "ข้อ ๒๔.๓"))

    if isinstance(doc, MinutesDoc):
        if not doc.agenda:
            f.append(Finding("error", "MIN_AGENDA", "ต้องมีข้อความการประชุมตามระเบียบวาระ", "agenda", "ข้อ ๒๕.๙"))
        for i, a in enumerate(doc.agenda):
            if not a.resolution or a.resolution.startswith("[ระบุ"):
                f.append(Finding("warning", "MIN_RESOLUTION", f"วาระที่ {i + 1} ยังไม่มีมติ/ข้อสรุป",
                                 f"agenda[{i}].resolution", "ข้อ ๒๕.๙"))
    return f


# ---------------------------------------------------------------------------- docx
def _rpr_child(rpr, tag: str):
    return rpr.find(tag) if rpr is not None else None


def _style_rprs(paragraph):
    """w:rPr ของสไตล์ย่อหน้าและสไตล์แม่ ไล่จากใกล้ไปไกล"""
    style = paragraph.style
    seen = 0
    while style is not None and seen < 10:
        yield style.element.rPr
        style, seen = style.base_style, seen + 1


def check_docx(path: str | Path, spec: LayoutSpec = DEFAULT_SPEC) -> list[Finding]:
    from docx import Document
    from docx.oxml.ns import qn

    d = Document(str(path))
    f: list[Finding] = []
    tol = 0.15
    for i, sec in enumerate(d.sections):
        for name, val, exp in (("ซ้าย", sec.left_margin, spec.margin_left_cm),
                               ("ขวา", sec.right_margin, spec.margin_right_cm)):
            cm = val.cm if val is not None else 0
            if abs(cm - exp) > tol:
                f.append(Finding("error", "MARGIN", f"ขอบ{name} {cm:.2f} ซม. (ควร {exp:g} ซม.)", f"section {i + 1}", _GUIDE))
        if sec.page_width and abs(sec.page_width.cm - spec.page_w_cm) > tol:
            f.append(Finding("warning", "PAPER", f"ความกว้างกระดาษ {sec.page_width.cm:.2f} ซม. ไม่ใช่ A4 "
                             f"({spec.page_w_cm:g} ซม.)", f"section {i + 1}"))

    def run_fonts(r, p) -> set[str]:
        """ชื่อฟอนต์ที่ run ใช้จริง (ascii/hAnsi/cs) — ถ้า run ไม่ระบุ ใช้ของสไตล์ย่อหน้า"""
        for rpr in (r._element.rPr, *_style_rprs(p)):
            rf = _rpr_child(rpr, qn("w:rFonts"))
            if rf is not None:
                names = {rf.get(qn(f"w:{a}")) for a in ("ascii", "hAnsi", "cs")}
                names.discard(None)
                if names:
                    return names
        return set()

    def run_sizes(r, p) -> list[float]:
        """ขนาด (พอยท์) ของ run: ทั้ง sz และ szCs (อักษรไทยใช้ szCs) — ถ้า run ไม่ระบุ ใช้ของสไตล์"""
        out: list[float] = []
        for tag in ("w:sz", "w:szCs"):
            for rpr in (r._element.rPr, *_style_rprs(p)):
                el = _rpr_child(rpr, qn(tag))
                if el is not None:
                    out.append(int(el.get(qn("w:val"))) / 2)
                    break
        return out

    paras = d.paragraphs
    cell_paras = [cp for t in d.tables for row in t.rows for cell in row.cells for cp in cell.paragraphs]
    all_paras = [*paras, *cell_paras]
    seen_ph: set[str] = set()
    for i, p in enumerate(all_paras, 1):
        where = f"ย่อหน้า {i}: {p.text.strip()[:30]}"
        f += text_quality(p.text, where)
        for ph in find_placeholders(p.text.replace("\xa0", " ")):       # ไฟล์ที่แก้ใน Word แล้วยังเหลือช่องว่าง
            if ph not in seen_ph:
                seen_ph.add(ph)
                f.append(Finding("warning", "PLACEHOLDER", f"ยังไม่ได้กรอก {ph}", where))

    bad_fonts: set[str] = set()
    arabic = small = 0
    for p in all_paras:
        if has_arabic_digits(p.text):
            arabic += 1
        for r in p.runs:
            if not r.text.strip():
                continue
            bad_fonts |= {n for n in run_fonts(r, p) if n not in NATIONAL_FONTS and n != spec.font}
            if any(sz < spec.body_pt - 0.5 for sz in run_sizes(r, p)):
                small += 1
    if bad_fonts:
        f.append(Finding("error", "FONT", f"พบฟอนต์ที่ไม่ใช่ฟอนต์แห่งชาติ: {', '.join(sorted(bad_fonts))}", "", _GUIDE))
    if small:
        f.append(Finding("warning", "FONT_SIZE", f"พบข้อความเล็กกว่า {spec.body_pt:g} พอยท์ {small} ช่วง", ""))
    if arabic:
        f.append(Finding("info", "ARABIC_DIGITS", f"พบเลขอารบิก {arabic} ย่อหน้า (หากหน่วยงานใช้เลขไทย)"))

    parts = [d.element.body, *(s.header._element for s in d.sections)]
    has_pic = any(next(part.iter(qn("pic:pic")), None) is not None for part in parts)
    first = next((p.text.strip() for p in paras if p.text.strip()), "")
    no_emblem_ok = first.startswith("รายงานการประชุม") or (not spec.news_emblem and first.startswith("ข่าว"))
    if not has_pic and not no_emblem_ok:
        f.append(Finding("warning", "GARUDA", f"ไม่พบภาพ{spec.emblem_name}", "", _GUIDE))

    # บันทึกข้อความ: หัว ๒๙ พอยท์ ตัวหนา / ป้าย ๒๐ พอยท์ ตัวหนา (ตรวจเมื่อพบหัว "บันทึกข้อความ" ในย่อหน้าต้นๆ)
    title = "บันทึกข้อความ"
    head = paras[:6]
    is_memo = False
    for p in head:
        for r in p.runs:
            if title in r.text and len(r.text.strip()) <= len(title) + 2:
                is_memo = True
                sizes = run_sizes(r, p)
                if any(sz != spec.memo_title_pt for sz in sizes) or not r.bold:
                    f.append(Finding("error", "MEMO_TITLE",
                                     f"'{title}' ต้องตัวหนา {spec.memo_title_pt:g} พอยท์", "", _GUIDE))
    if is_memo:
        for p in head:
            for label in (spec.unit_word, "วันที่", "เรื่อง"):
                for r in p.runs:
                    if r.text.strip() == label and (any(sz != spec.memo_label_pt for sz in run_sizes(r, p)) or not r.bold):
                        f.append(Finding("error", "MEMO_LABEL",
                                         f"'{label}' ต้องตัวหนา {spec.memo_label_pt:g} พอยท์", "", _GUIDE))
    return f


def summarize(findings: list[Finding]) -> dict:
    return {
        "errors": sum(x.level == "error" for x in findings),
        "warnings": sum(x.level == "warning" for x in findings),
        "infos": sum(x.level == "info" for x in findings),
        "findings": [x.dict() for x in findings],
    }
