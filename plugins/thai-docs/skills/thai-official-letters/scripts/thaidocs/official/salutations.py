"""คำขึ้นต้น / คำลงท้าย ตามฐานะผู้รับ (ภาคผนวก ๒ ท้ายระเบียบสารบรรณ) ชุดที่ใช้บ่อย

ดัดแปลงจากโครงการต้นทางภายใต้ Apache License 2.0 ดู NOTICE ของ repo
"""
from __future__ import annotations

# (คำขึ้นต้น, คำลงท้าย)
HIGH_RANK = ("กราบเรียน", "ขอแสดงความนับถืออย่างยิ่ง")
GENERAL = ("เรียน", "ขอแสดงความนับถือ")

HIGH_RANK_TITLES = [
    "ประธานองคมนตรี", "นายกรัฐมนตรี", "ประธานรัฐสภา", "ประธานสภาผู้แทนราษฎร",
    "ประธานวุฒิสภา", "ประธานศาลฎีกา", "ประธานศาลรัฐธรรมนูญ", "ประธานศาลปกครองสูงสุด",
    "รัฐบุรุษ",
]

VALID_PAIRS = {
    "กราบเรียน": {"ขอแสดงความนับถืออย่างยิ่ง"},
    "เรียน": {"ขอแสดงความนับถือ", "ขอแสดงความนับถืออย่างสูง"},
}


def suggest(recipient: str) -> tuple[str, str]:
    r = recipient.strip()
    if any(r.startswith(t) for t in HIGH_RANK_TITLES):
        return HIGH_RANK
    return GENERAL


def pair_ok(salutation: str, closing: str) -> bool:
    return closing in VALID_PAIRS.get(salutation, set())
