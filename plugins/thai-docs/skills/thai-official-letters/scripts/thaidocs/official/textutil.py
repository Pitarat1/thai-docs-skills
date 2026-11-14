"""ตัวช่วยข้อความสำหรับหนังสือราชการ: placeholder, วันที่ไทย, เลขไทย/อารบิก, จุดห้ามตัดบรรทัด

ไม่พึ่ง pydantic
"""
from __future__ import annotations

import re
from datetime import date

from ..digits import to_thai_digits

ARABIC = "0123456789"
THAI = "๐๑๒๓๔๕๖๗๘๙"
_TO_ARABIC = str.maketrans(THAI, ARABIC)

THAI_MONTHS = [
    "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน",
    "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม",
]

# [ระบุ...] คือช่องที่ห้ามแต่งขึ้นเอง ปล่อยให้ผู้ใช้เติม
PLACEHOLDER_RE = re.compile(r"\[ระบุ[^\]]*\]")
ARABIC_DIGIT_RE = re.compile(r"[0-9]")
# ข้าม URL / อีเมลเมื่อตรวจเลขอารบิก
_KEEP_RE = re.compile(r"(https?://\S+|[\w.+-]+@[\w-]+\.[\w.]+)")

_CONTROL_RE = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]")


def find_placeholders(text: str) -> list[str]:
    return PLACEHOLDER_RE.findall(text or "")


def is_placeholder(text: str | None) -> bool:
    """ว่างหรือขึ้นต้นด้วย [ระบุ"""
    t = (text or "").strip()
    return not t or t.startswith("[ระบุ")


def to_arabic_digits(text: str) -> str:
    return text.translate(_TO_ARABIC)


def has_arabic_digits(text: str) -> bool:
    stripped = _KEEP_RE.sub("", text or "")
    return bool(ARABIC_DIGIT_RE.search(stripped))


def thai_date(d: date | None = None, *, with_day_prefix: bool = False) -> str:
    """date(2026, 11, 14) -> '๑๔ พฤศจิกายน ๒๕๖๙'"""
    d = d or date.today()
    s = f"{d.day} {THAI_MONTHS[d.month - 1]} {d.year + 543}"
    if with_day_prefix:
        s = "วันที่ " + s
    return to_thai_digits(s)


def clean_text(text: str) -> str:
    """ตัดอักขระควบคุม ยกเว้น \\n และ \\t"""
    return _CONTROL_RE.sub("", text or "")


def split_paragraphs(items: list[str]) -> list[str]:
    """แยกที่ \\n ตัดรายการว่าง"""
    out: list[str] = []
    for item in items:
        for line in (item or "").split("\n"):
            if line.strip():
                out.append(line)
    return out


# ---- จุดที่ห้ามตัดบรรทัด: แทนช่องว่างด้วย no-break space (U+00A0) ความกว้างเท่าเดิม
NBSP = " "
_D = "[0-9๐-๙]"
_MONTHS = "|".join(THAI_MONTHS) + r"|ม\.ค\.|ก\.พ\.|มี\.ค\.|เม\.ย\.|พ\.ค\.|มิ\.ย\.|ก\.ค\.|ส\.ค\.|ก\.ย\.|ต\.ค\.|พ\.ย\.|ธ\.ค\."
_UNITS = "ชุด|ฉบับ|คน|ราย|แห่ง|บาท|น\\.|นาฬิกา|วัน|ครั้ง|เล่ม|หน้า|อำเภอ|ปี|เดือน|ชั่วโมง|นาที"
_NO_BREAK = [
    re.compile(rf"(?<=พ\.ศ\.)[ \t]+(?={_D})"),             # พ.ศ. ๒๕๖๙
    re.compile(rf"(?<=ค\.ศ\.)[ \t]+(?={_D})"),
    re.compile(rf"(?<={_D})[ ]+(?=(?:{_MONTHS}))"),         # ๑๕ พฤศจิกายน
    re.compile(rf"(?<=[ก-๙.])[ ]+(?={_D}{{4}}(?!{_D}))"),    # พฤศจิกายน ๒๕๖๙ / ปี ๒๕๖๙
    re.compile(rf"(?<={_D})[ ]+(?=(?:{_UNITS}))"),          # ๑ ชุด / ๐๙.๓๐ น.
    re.compile(rf"(?<=จำนวน)[ ]+(?={_D})"),
    re.compile(rf"(?<=ที่)[ ]+(?={_D})"),
    re.compile(rf"(?<=เวลา)[ ]+(?={_D})"),
    re.compile(rf"(?<=โทร\.)[ ]*(?={_D})"),
    re.compile(rf"(?<={_D})[ ](?={_D})"),                   # เลขโทรศัพท์ ๐ ๒๐๐๐ ๐๐๐๐
    re.compile(r"(?:(?<=^ณ)|(?<=[^ก-๙]ณ))[ ]+"),            # ณ ห้องประชุม
    re.compile(rf"(?<=ต่อ)[ ]+(?={_D})"),                   # โทร ... ต่อ ๑๐๙
    re.compile(r"[ ]+(?=ๆ)"),                               # อื่น ๆ
]


def protect_breaks(text: str) -> str:
    """แทนช่องว่างในจุดที่ไม่ควรขึ้นบรรทัดใหม่ด้วย no-break space"""
    if not text:
        return text
    for rx in _NO_BREAK:
        text = rx.sub(NBSP, text)
    return text


# ---- รวมบรรทัดแบบ "วันที่: … / เวลา: … / สถานที่: …" เข้าเป็นประโยค
_TD = "0-9๐-๙"
_LABELS = [
    (re.compile(r"^(?:วันที่|วัน|ในวันที่)$"), "ในวันที่"),
    (re.compile(r"^(?:เวลา|ช่วงเวลา)$"), "เวลา"),
    (re.compile(r"^(?:สถานที่|ณ|สถานที่จัด|สถานที่ประชุม)$"), "ณ"),
    (re.compile(r"^(?:ผู้เข้าร่วม(?:ประชุม)?|กลุ่มเป้าหมาย)$"), "โดยมีผู้เข้าร่วม"),
]
LIST_LINE_RE = re.compile(
    r"^\s*(?:[-•*]\s*)?(วันที่|วัน|ในวันที่|เวลา|ช่วงเวลา|สถานที่(?:จัด|ประชุม)?|ณ|ผู้เข้าร่วม(?:ประชุม)?|กลุ่มเป้าหมาย)"
    r"\s*[:：]\s*(.+?)\s*$")
_LEAD_IN_RE = re.compile(r"\s*(?:โดยมี)?(?:รายละเอียด|กำหนดการ)?\s*ดังนี้\s*[:：]?\s*$")
_TIME_RANGE_RE = re.compile(
    rf"([{_TD}]{{1,2}}[.:][{_TD}]{{2}})\s*(?:น\.)?\s*(?:ถึง|-|–|—)\s*([{_TD}]{{1,2}}[.:][{_TD}]{{2}})\s*น\.?")


def _phrase(label: str, value: str) -> str:
    for rx, word in _LABELS:
        if rx.match(label):
            value = value.strip()
            if word != "โดยมีผู้เข้าร่วม":
                value = re.sub(rf"^(?:{re.escape(word)}|วันที่|เวลา|ณ)\s*", "", value)
            return f"{word} {value}".strip()
    return value


def time_range(text: str) -> str:
    """'๑๓.๓๐ น. ถึง ๑๖.๓๐ น.' → '๑๓.๓๐ – ๑๖.๓๐ น.'"""
    return _TIME_RANGE_RE.sub(lambda m: f"{m.group(1)} – {m.group(2)} น.", text)


def merge_list_lines(paragraphs: list[str]) -> list[str]:
    """รวมบรรทัด 'วันที่: … / เวลา: … / สถานที่: …' เข้ากับย่อหน้าก่อนหน้าเป็นประโยคเดียว"""
    out: list[str] = []
    for p in paragraphs:
        parts = [x for x in re.split(r"\n+", p) if x.strip()] or [p]
        for part in parts:
            m = LIST_LINE_RE.match(part)
            if m and out:
                prev = _LEAD_IN_RE.sub("", out[-1]).rstrip()
                out[-1] = f"{prev} {_phrase(m.group(1), m.group(2))}"
            else:
                out.append(part.strip())
    return [time_range(x) for x in out]
