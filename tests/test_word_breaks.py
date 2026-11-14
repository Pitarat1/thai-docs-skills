"""ตรรกะนับการขึ้นบรรทัดกลางคำของ tools/word_breaks.py (ไม่ต้องใช้ Word)"""
from __future__ import annotations

import importlib.util
from pathlib import Path

_spec = importlib.util.spec_from_file_location(
    "word_breaks", Path(__file__).resolve().parent.parent / "tools" / "word_breaks.py")
wb = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(wb)

LINE = "ผล|การ|ศึกษา พบ|ว่า"


def _para(line=LINE):
    text = line.replace("|", "")
    allowed, pos = set(), 0
    for ch in line:
        if ch == "|":
            allowed.add(pos)
            continue
        if ch == " ":
            allowed.update((pos, pos + 1))
        pos += 1
    return text, allowed


def test_gold_parse_matches_markers():
    text, allowed = _para()
    assert text == "ผลการศึกษา พบว่า"
    assert {2, 5, 10, 11, 13} <= allowed


def test_break_at_word_edge_is_ok():
    text, allowed = _para()
    total, bad, mism = wb.count_breaks(["ผลการ", "ศึกษา พบว่า"], text, allowed)
    assert (total, bad, mism) == (1, [], 0)


def test_break_after_space_is_ok():
    text, allowed = _para()
    total, bad, _ = wb.count_breaks(["ผลการศึกษา", "พบว่า"], text, allowed)
    assert (total, bad) == (1, [])


def test_mid_word_break_counted():
    text, allowed = _para()
    total, bad, mism = wb.count_breaks(["ผลการศึ", "กษา พบว่า"], text, allowed)
    assert total == 1 and bad == [0] and mism == 0


def test_only_last_line_has_no_break_and_text_mismatch_flagged():
    text, allowed = _para()
    assert wb.count_breaks([text], text, allowed) == (0, [], 0)
    assert wb.count_breaks(["ผลการ", "อื่นๆ"], text, allowed)[2] == 1


def test_real_gold_parses():
    gold = wb.read_gold()
    assert len(gold) == 10 and all(t and a for t, a in gold)


def test_decomposed_sara_am_from_pdf_is_matched():
    text, allowed = "น้ำ|มัน", {2}
    text = "น้ำมัน"
    assert wb.count_breaks(["น้ํา", "มัน"], text, {3}) == (1, [], 0)
