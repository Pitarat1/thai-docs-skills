"""เทสตัวตัดคำ (newmm ∩ khammoon) กับชุด gold ที่กำหนดจุดตัดด้วยมือ

รูปแบบ tests/data/gold_th.txt: บรรทัดละย่อหน้า, ``|`` = จุดตัดที่อนุญาต (ไม่ใช่ตัวอักษร),
ช่องว่างที่ตำแหน่ง p = ตัดได้ทั้งที่ p และ p+1 (ก่อนและหลังช่องว่าง)
"""
from __future__ import annotations

import warnings
from pathlib import Path

import pytest

from thaidocs import wordcut
from thaidocs.wordcut import Segmenter, khammoon_available

GOLD = Path(__file__).parent / "data" / "gold_th.txt"


def read_gold(path: Path = GOLD) -> list[tuple[str, set[int]]]:
    """คืน [(ข้อความ, ชุดจุดตัดที่อนุญาต)]"""
    paras = []
    for line in path.read_text(encoding="utf-8").splitlines():
        if not line.strip():
            continue
        allowed: set[int] = set()
        pos = 0
        for ch in line:
            if ch == "|":
                allowed.add(pos)
                continue
            if ch == " ":
                allowed.update((pos, pos + 1))
            pos += 1
        paras.append((line.replace("|", ""), allowed))
    return paras


def bad_rate(seg: Segmenter) -> tuple[float, list[str]]:
    """สัดส่วนจุดตัดที่ไม่อยู่ใน gold + ตัวอย่าง"""
    bad = tot = 0
    examples = []
    for text, allowed in read_gold():
        bs = seg.boundaries(text)
        tot += len(bs)
        for b in sorted(bs - allowed):
            bad += 1
            examples.append(text[max(0, b - 8):b] + "¦" + text[b:b + 8])
    return bad / tot, examples


@pytest.fixture(scope="module")
def seg_km():
    if not khammoon_available():
        pytest.skip("ไม่มี numpy/khammoon")
    return Segmenter(use_khammoon=True)


@pytest.fixture(scope="module")
def seg_dict():
    return Segmenter(use_khammoon=False)


def test_gold_file_shape():
    paras = read_gold()
    assert len(paras) == 10
    assert all(len(t) > 200 for t, _ in paras)


def test_text_preserved(seg_dict):
    segs = [seg_dict]
    if khammoon_available():
        segs.append(Segmenter(use_khammoon=True))
    texts = [t for t, _ in read_gold()] + ["", " ", "ระบบHIS 2.0ของรพ.", "abc", "  ไทย  "]
    for s in segs:
        for t in texts:
            ws = s.words(t)
            assert "".join(ws) == t
            assert all(ws)
            bs = s.boundaries(t)
            assert all(0 < b < len(t) for b in bs)


def test_gold_with_khammoon(seg_km):
    rate, ex = bad_rate(seg_km)
    assert rate <= 0.03, f"{rate:.1%} {ex}"


def test_gold_without_khammoon(seg_dict):
    rate, ex = bad_rate(seg_dict)
    assert rate <= 0.06, f"{rate:.1%} {ex}"


def test_without_khammoon_fallback(monkeypatch):
    monkeypatch.setattr(wordcut, "khammoon_available", lambda: False)
    monkeypatch.setattr(wordcut, "_warned", False)
    with warnings.catch_warnings(record=True) as rec:
        warnings.simplefilter("always")
        s1 = Segmenter()
        s2 = Segmenter()
    msgs = [str(w.message) for w in rec]
    assert msgs == ["ไม่พบ khammoon/numpy: ตัดคำด้วยพจนานุกรมอย่างเดียว คำทับศัพท์อาจขาดกลางคำ"]
    assert not s1.use_khammoon and not s2.use_khammoon
    t = "สำนักงานสาธารณสุขจังหวัดสระแก้ว"
    assert "".join(s1.words(t)) == t
    assert len(s1.words(t)) > 1


def test_require_khammoon_missing_raises(monkeypatch):
    monkeypatch.setattr(wordcut, "khammoon_available", lambda: False)
    with pytest.raises(RuntimeError, match="khammoon"):
        Segmenter(use_khammoon=True)


def test_explicit_false_does_not_warn(monkeypatch):
    monkeypatch.setattr(wordcut, "_warned", False)
    with warnings.catch_warnings(record=True) as rec:
        warnings.simplefilter("always")
        Segmenter(use_khammoon=False)
    assert rec == []


@pytest.mark.parametrize("use_km", [False, True])
def test_mixed_script_boundaries(use_km):
    if use_km and not khammoon_available():
        pytest.skip("ไม่มี numpy/khammoon")
    t = "ระบบHIS 2.0ของรพ."
    bs = _seg(use_km).boundaries(t)
    assert {4, 7, 8, 11} <= bs
    assert not bs & {5, 6}          # ภายใน HIS
    assert not bs & {9, 10}         # ภายใน 2.0
    assert len(t) - 1 not in bs     # ไม่แยก "." ออกจาก "รพ"


def test_thai_abbreviation_kept(seg_dict):
    t = "ตามพระราชบัญญัติพ.ศ. 2562"
    bs = seg_dict.boundaries(t)
    i = t.index("พ.ศ.")
    assert not bs & {i + 1, i + 2, i + 3}


def test_custom_words_used():
    base = Segmenter(use_khammoon=False)
    assert "ซูเปอร์แอป" not in base.words("ใช้ซูเปอร์แอปทุกวัน")
    from thaidocs.dictionary import load_words
    s = Segmenter(words=load_words() | {"ซูเปอร์แอป"}, use_khammoon=False)
    assert "ซูเปอร์แอป" in s.words("ใช้ซูเปอร์แอปทุกวัน")


_SEGS: dict[bool, Segmenter] = {}


def _seg(use_km):
    """Segmenter ที่ใช้ซ้ำ (load_words + สร้าง trie ใหม่ทุกครั้งช้า)"""
    if use_km and not khammoon_available():
        pytest.skip("ไม่มี numpy/khammoon")
    if use_km not in _SEGS:
        _SEGS[use_km] = Segmenter(use_khammoon=use_km)
    return _SEGS[use_km]


@pytest.mark.parametrize("use_km", [False, True])
@pytest.mark.parametrize("text,sub", [
    ("ทำเสร็จแล้ว.ต่อไป", "แล้ว."),
    ("สบายดี...แล้ว", "ดี..."),
    ("ข้อ ๑.ให้ดำเนินการ", "๑.ให้"),
    ("รพ.สต.บ้านนา", "รพ.สต."),
    ("ตามพระราชบัญญัติพ.ศ. 2562", "พ.ศ."),
    ("รอสักครู่…แล้วค่อยไป", "ครู่…"),
])
def test_punctuation_not_split(use_km, text, sub):
    bs = _seg(use_km).boundaries(text)
    i = text.index(sub)
    inner = set(range(i + 1, i + len(sub)))
    assert not bs & inner, (text, sorted(bs & inner))


@pytest.mark.parametrize("use_km", [False, True])
def test_comma_and_paren_not_line_start(use_km):
    s = _seg(use_km)
    t = "ระบบ(HIS)ของโรงพยาบาล,ทุกแห่ง"
    bs = s.boundaries(t)
    assert t.index("(") + 1 not in bs          # ไม่ตัดหลังวงเล็บเปิด
    assert t.index(")") not in bs              # ไม่ตัดก่อนวงเล็บปิด
    assert t.index(",") not in bs              # ไม่ตัดก่อนจุลภาค


@pytest.mark.parametrize("use_km", [False, True])
@pytest.mark.parametrize("text", ['"ประกาศ"ของ', "“ประกาศ”ของ", "เรื่อง \"ประกาศ\"ของ"])
def test_quotes_by_context(use_km, text):
    ws = _seg(use_km).words(text)
    assert "".join(ws) == text
    q = text[-4]  # อัญประกาศปิด
    assert any(w.endswith("ประกาศ" + q) for w in ws), ws   # ปิดติดคำหน้า
    assert ws[-1] == "ของ", ws                              # ตัดหลังอัญประกาศปิดได้
    assert not any(w in ('"', "“", "”") for w in ws), ws


@pytest.mark.parametrize("use_km", [False, True])
@pytest.mark.parametrize("nb", [" ", " ", " ", "⁠", "﻿"])
def test_no_break_spaces(use_km, nb):
    s = _seg(use_km)
    for text in (f"นาย{nb}สมชาย ใจดี", f"พ.ศ.{nb}2569", f"ราคา 100{nb}บาท"):
        bs = s.boundaries(text)
        i = text.index(nb)
        assert not bs & {i, i + 1}, (repr(text), sorted(bs))
        assert "".join(s.words(text)) == text


@pytest.mark.parametrize("km", [True, False])
def test_surname_ending_in_function_word_does_not_hijack(km):
    # ชื่อ/นามสกุลที่ลงท้ายด้วยคำหน้าที่ (เช่น แก้วมี) ต้องไม่ดึงคำ "มี" ออกจากประโยคทั่วไป
    if km and not khammoon_available():
        pytest.skip("ไม่มี numpy/khammoon")
    words = Segmenter(use_khammoon=km).words("จังหวัดสระแก้วมีความประสงค์")
    assert words == ["จังหวัด", "สระแก้ว", "มี", "ความประสงค์"]
