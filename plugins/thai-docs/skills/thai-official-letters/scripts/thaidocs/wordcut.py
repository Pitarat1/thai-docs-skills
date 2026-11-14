"""ตัวตัดคำไทยสำหรับหาจุดที่ขึ้นบรรทัดใหม่ได้โดยไม่ขาดกลางคำ

จุดตัด (boundary) คือ offset ตัวอักษรใน ``text`` ที่ขึ้นบรรทัดใหม่ "ก่อน" ตัวอักษรตำแหน่งนั้นได้
มีค่า ``0 < b < len(text)`` เสมอ (ไม่รวม 0 และ len(text))

    boundaries = (จุดตัด newmm ∩ จุดตัด khammoon threshold 0.9)   เฉพาะภายในช่วงอักษรไทย
                 ∪ ขอบช่องว่าง (ก่อนและหลังอักขระช่องว่างทุกตัว)
                 ∪ จุดเปลี่ยนไทย/ไม่ใช่ไทย

- ช่วงอักษรไทย = อักขระ U+0E00–U+0E7F ต่อกัน รวม "." ที่ตามหลังอักษรไทย (คำย่อ เช่น พ.ศ. รพ.สต.)
  จุดตัดภายในช่วงนี้ต้องให้ทั้ง newmm (พจนานุกรมรวม) และ khammoon เห็นตรงกัน
  newmm ไม่เคยรายงานคำที่ไม่รู้จัก มันซอยชื่อเฉพาะเป็นคำที่รู้จัก (เมท|ฟอร์|มิน) จึงใช้ intersection
- ช่วงที่ไม่ใช่ไทย (ละติน ตัวเลข เครื่องหมาย) ไม่มีจุดตัดภายใน เช่น ``HIS``, ``2.0``
- จุดเปลี่ยนไทย/ไม่ใช่ไทย = ระหว่างอักษรไทยกับอักขระที่ไม่ใช่ไทยและไม่ใช่ช่องว่าง
- กรองจุดตัดทุกจุด (ทั้งจาก newmm∩khammoon และจุดเปลี่ยน) ด้วย ``_break_ok``: ไม่ตัดหลังวงเล็บ/
  อัญประกาศเปิด, ไม่ตัดก่อนวงเล็บปิด/อัญประกาศปิด/เครื่องหมายวรรคตอนท้าย (. … , ฯลฯ),
  ไม่ตัดระหว่าง "." กับอักษรไทยที่ตามมา (คำย่อ รพ.สต. อยู่ด้วยกัน)
  อัญประกาศ ASCII (" ') ตัดสินจากบริบท: ต้นข้อความหรือหลังช่องว่าง/วงเล็บเปิด = เปิด, นอกนั้น = ปิด
- ช่องว่างไม่ตัดบรรทัด (U+00A0, U+202F, U+2007, U+2060, U+FEFF) ไม่นับเป็นช่องว่าง และไม่มีจุดตัด
  ติดกับอักขระเหล่านี้เลย
- ไม่มี khammoon (ไม่มี numpy หรือถูกถอดออก) → ใช้ newmm อย่างเดียว
"""
from __future__ import annotations

import unicodedata
import warnings

from .dictionary import load_words, newmm_words

KHAMMOON_THRESHOLD = 0.9
WARN_NO_KHAMMOON = "ไม่พบ khammoon/numpy: ตัดคำด้วยพจนานุกรมอย่างเดียว คำทับศัพท์อาจขาดกลางคำ"

_warned = False
_km = None  # khammoon Segmenter ตัวเดียวทั้งโปรเซส (โหลดน้ำหนัก ~4 MB ครั้งเดียว)

_NO_BREAK_BEFORE = set(".,:;!?%)]}»”’…")
_QUOTES = "\"'"
NOBREAK_SPACES = frozenset("\u00a0\u202f\u2007\u2060\ufeff")


def khammoon_available() -> bool:
    """True ถ้า import numpy และ thaidocs._khammoon ได้ และมีไฟล์โมเดลครบ"""
    try:
        import numpy  # noqa: F401

        from ._khammoon import MODEL_DIR
    except Exception:
        return False
    return all((MODEL_DIR / n).is_file() for n in ("weights.npz", "vocab.json", "meta.json"))


def _khammoon():
    global _km
    if _km is None:
        from ._khammoon import Segmenter as _KM

        _km = _KM(threshold=KHAMMOON_THRESHOLD)
    return _km


def _is_thai(ch: str) -> bool:
    return "฀" <= ch <= "๿"


def _thai_mask(text: str) -> list[bool]:
    """True ที่อักขระอยู่ในช่วงอักษรไทย (รวม "." ที่ตามหลังอักษรไทย เช่น พ.ศ.)"""
    mask: list[bool] = []
    for i, ch in enumerate(text):
        mask.append(_is_thai(ch) or (ch == "." and i > 0 and mask[i - 1]))
    return mask


def _cuts(tokens: list[str]) -> set[int]:
    out, pos = set(), 0
    for t in tokens[:-1]:
        pos += len(t)
        out.add(pos)
    return out


def _is_space(ch: str) -> bool:
    """ช่องว่างที่ตัดบรรทัดได้ (ไม่รวมช่องว่างไม่ตัดบรรทัด)"""
    return ch.isspace() and ch not in NOBREAK_SPACES


def _opens(text: str, i: int) -> bool:
    """text[i] เป็นวงเล็บ/อัญประกาศเปิด"""
    ch = text[i]
    if unicodedata.category(ch) in ("Ps", "Pi"):
        return True
    if ch in _QUOTES:
        return i == 0 or text[i - 1].isspace() or unicodedata.category(text[i - 1]) in ("Ps", "Pi")
    return False


def _closes(text: str, i: int) -> bool:
    """text[i] เป็นวงเล็บ/อัญประกาศปิด หรือเครื่องหมายที่ห้ามขึ้นต้นบรรทัด"""
    ch = text[i]
    if ch in _NO_BREAK_BEFORE or unicodedata.category(ch) in ("Pe", "Pf"):
        return True
    return ch in _QUOTES and not _opens(text, i)


def _break_ok(text: str, b: int) -> bool:
    """ตัดบรรทัดก่อน text[b] ได้หรือไม่ (ใช้กรองจุดตัดทุกจุด)"""
    a, c = text[b - 1], text[b]
    if a in NOBREAK_SPACES or c in NOBREAK_SPACES:
        return False
    if a.isspace() or c.isspace():
        return True
    if _opens(text, b - 1) or _closes(text, b):
        return False
    if a == "." and _is_thai(c):
        return False
    return True


class Segmenter:
    """ตัดคำไทย: พจนานุกรมรวม (newmm) ∩ khammoon

    words: ชุดคำ (None = ``load_words()`` ครั้งเดียวต่อ Segmenter)
    use_khammoon: None = ใช้ถ้ามี (ไม่มีจะเตือนครั้งเดียว), True = บังคับ (ไม่มี → RuntimeError),
                  False = พจนานุกรมอย่างเดียว
    """

    def __init__(self, words: frozenset[str] | None = None, use_khammoon: bool | None = None):
        global _warned
        self.words_set = load_words() if words is None else frozenset(words)
        available = khammoon_available() if use_khammoon is not False else False
        if use_khammoon is True and not available:
            raise RuntimeError(
                "ไม่พบ khammoon หรือ numpy: ติดตั้ง numpy (pip install numpy) "
                "หรือใช้ use_khammoon=False เพื่อตัดคำด้วยพจนานุกรมอย่างเดียว"
            )
        if use_khammoon is None and not available and not _warned:
            _warned = True
            warnings.warn(WARN_NO_KHAMMOON, stacklevel=2)
        self.use_khammoon = available
        self._km = _khammoon() if available else None

    def boundaries(self, text: str) -> set[int]:
        """จุดตัดเป็น offset ตัวอักษร ``0 < b < len(text)``"""
        n = len(text)
        if n < 2:
            return set()
        thai = _thai_mask(text)
        out: set[int] = set()
        if any(thai):
            cand = _cuts(newmm_words(text, self.words_set))
            if self._km is not None and cand:
                cand &= _cuts(self._km.segment(text))
            out |= {b for b in cand if thai[b - 1] and thai[b]}
        for i in range(1, n):
            a, b = text[i - 1], text[i]
            if _is_space(a) or _is_space(b):
                if not (_is_space(a) and _is_space(b)):
                    out.add(i)
            elif thai[i - 1] != thai[i] and not (a.isspace() or b.isspace()):
                out.add(i)
        return {b for b in out if _break_ok(text, b)}

    def words(self, text: str) -> list[str]:
        """แบ่ง text ตาม boundaries(); ``"".join(words(text)) == text``"""
        if not text:
            return []
        idx = [0, *sorted(self.boundaries(text)), len(text)]
        return [text[a:b] for a, b in zip(idx, idx[1:])]
