"""ตรวจเอกสารว่าจัดหน้าไทยถูกต้อง: คืนรายการ Issue (ว่าง = ผ่าน; WRAPPED เป็นข้อมูล ไม่นับเป็นปัญหา)"""
from __future__ import annotations

from dataclasses import dataclass

from docx.oxml.ns import qn

from .digits import KEEP, _ARABIC
from .layout import _KEEP_BREAK, _PUA_RE, FONT_ATTRS, THAI, THEME_ATTRS, _is_symbol
from .ooxml import WRAPPED_PROP, is_wrapped, iter_content_parts, iter_parts  # noqa: F401


@dataclass
class Issue:
    code: str
    message: str
    count: int


def _text(el) -> str:
    return "".join(t.text or "" for t in el.iter(qn("w:t")))


def check(doc, *, thai_digits: bool = False) -> list[Issue]:
    wrapped = is_wrapped(doc)
    no_cs = no_lang = sz_bad = theme = manual = bad = arabic = 0
    fonts: set[str] = set()

    for root in iter_content_parts(doc):
        for r in root.iter(qn("w:r")):
            if not wrapped and THAI.search(_text(r)):
                rpr = r.find(qn("w:rPr"))
                if rpr is None or rpr.find(qn("w:cs")) is None:
                    no_cs += 1
                else:
                    lang = rpr.find(qn("w:lang"))
                    if lang is None or lang.get(qn("w:bidi")) != "th-TH":
                        no_lang += 1
        for p in root.iter(qn("w:p")):
            seq = [e for e in p.iter(qn("w:t"), qn("w:br"))
                   if e.tag == qn("w:t") or e.get(qn("w:type")) in (None, "textWrapping")]
            for i, e in enumerate(seq):
                if e.tag == qn("w:br"):
                    after = "".join(x.text or "" for x in seq[i + 1:] if x.tag == qn("w:t"))
                    if not after.lstrip().startswith(_KEEP_BREAK):
                        manual += 1
            full = "".join(t.text or "" for t in p.iter(qn("w:t")))
            bad += len(_PUA_RE.findall(full)) + full.count("ํา")
            if thai_digits:
                arabic += len(_ARABIC.findall(KEEP.sub("", full)))

    for root in iter_parts(doc):
        for rpr in root.iter(qn("w:rPr")):
            sz, cs = rpr.find(qn("w:sz")), rpr.find(qn("w:szCs"))
            szv = sz.get(qn("w:val")) if sz is not None else None
            csv = cs.get(qn("w:val")) if cs is not None else None
            if (szv or csv) and szv != csv:
                sz_bad += 1
        for rf in root.iter(qn("w:rFonts")):
            if _is_symbol(rf):
                continue
            if any(rf.get(qn(a)) is not None for a in THEME_ATTRS):
                theme += 1
            fonts.update(v for a in FONT_ATTRS if (v := rf.get(qn(a))))

    issues: list[Issue] = []
    if no_cs:
        issues.append(Issue("NO_CS", f"พบ run ภาษาไทย {no_cs} ตำแหน่งที่ไม่มี w:cs (Word จะขีดเส้นหยักและตัดคำผิด)", no_cs))
    if no_lang:
        issues.append(Issue("NO_LANG", f"พบ run ภาษาไทย {no_lang} ตำแหน่งที่ไม่มี w:lang bidi=th-TH (Word ตรวจคำสะกดเป็นภาษาอื่น)", no_lang))
    if sz_bad:
        issues.append(Issue("SZ_MISMATCH", f"พบ {sz_bad} ตำแหน่งที่ขนาดอักษร sz ไม่เท่ากับ szCs", sz_bad))
    if theme:
        issues.append(Issue("THEME_FONT", f"พบ {theme} ตำแหน่งที่ฟอนต์อ้างอิง theme (theme ชนะฟอนต์ที่ตั้งตรง)", theme))
    if len(fonts) > 1:
        issues.append(Issue("MULTI_FONT", f"พบฟอนต์มากกว่าหนึ่งชนิด ({', '.join(sorted(fonts))})", len(fonts)))
    if manual:
        issues.append(Issue("MANUAL_BREAK", f"พบการขึ้นบรรทัดด้วยมือ {manual} ตำแหน่ง", manual))
    if bad:
        issues.append(Issue("BAD_CHARS", f"พบอักษรไทยเพี้ยน (Private Use / นิคหิต+า) {bad} ตัว", bad))
    if thai_digits and arabic:
        issues.append(Issue("ARABIC_DIGITS", f"พบเลขอารบิก {arabic} ตัวที่ควรเป็นเลขไทย", arabic))
    if wrapped:
        issues.append(Issue("WRAPPED", "เอกสารถูกตัดคำล่วงหน้า (wrap) แล้ว", 1))
    return issues
