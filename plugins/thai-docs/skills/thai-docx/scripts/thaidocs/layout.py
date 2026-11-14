"""จัดหน้า Word ไทยขั้นสุดท้าย — ทำให้ Word ตัดคำ/กระจายบรรทัดภาษาไทยถูกต้อง

ที่มา: SARABUN.AI (การประปาส่วนภูมิภาค / Provincial Waterworks Authority, Apache-2.0)
engine/sarabun/render/thai_layout.py และ skill "thai-docx-layout-th" หมวด 9
Copyright 2026 การประปาส่วนภูมิภาค — ใช้ภายใต้ Apache License 2.0
ดัดแปลง: เปลี่ยนจากแก้ XML ดิบเป็นทำงานบน docx.Document และทุก part ที่เกี่ยวข้อง
(เนื้อหา หัว/ท้ายกระดาษ เชิงอรรถ styles numbering) ให้ทุกฟังก์ชัน idempotent ดู NOTICE ของ repo

หลักการ:
  - ทุก run ที่มีอักษรไทยต้องมี <w:cs/> + <w:lang w:val="en-US" w:bidi="th-TH"/>
    ไม่งั้น Word ถือเป็นอังกฤษ: ตัดคำผิด ช่องไฟยืด ขึ้นเส้นหยักแดง
  - ฟอนต์เดียวทั้งเอกสาร ลบ attribute theme (theme ชนะค่าที่ตั้งตรง)
  - sz = szCs (ภาษาไทยใช้ szCs), ตัวหนาต้องมีทั้ง b และ bCs
  - เรียง element ใน rPr / pPr ตาม schema
  - ซ่อนเส้นตรวจคำ (hideSpellingErrors / hideGrammaticalErrors)

ฟังก์ชันรับ docx.Document แก้ในที่ คืนจำนวนที่แก้จริง (0 เมื่อทำซ้ำ)
"""
from __future__ import annotations

import re

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from .ooxml import PPR_ORDER, RPR_ORDER, child, iter_content_parts, iter_parts, local, reorder

THAI = re.compile("[฀-๿]")
SYMBOL_FONTS = {"Symbol", "Wingdings", "Wingdings 2", "Wingdings 3", "Courier New"}
THEME_ATTRS = ("w:asciiTheme", "w:hAnsiTheme", "w:cstheme", "w:eastAsiaTheme")
FONT_ATTRS = ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia")

# อักษร Private Use (U+F700–F71A, ผัง Microsoft/WTT) จากฟอนต์ไทยรุ่นเก่า -> อักษรไทยมาตรฐาน
PUA = {
    "\uf700": "\u0e10",  # ฐ ไม่มีเชิง
    **{chr(0xF701 + i): c for i, c in enumerate("\u0e34\u0e35\u0e36\u0e37")},  # สระบน (ชิด)
    **{chr(0xF705 + i): c for i, c in enumerate("\u0e48\u0e49\u0e4a\u0e4b\u0e4c")},  # วรรณยุกต์/การันต์ (บนสระ)
    **{chr(0xF70A + i): c for i, c in enumerate("่้๊๋์")},  # วรรณยุกต์/การันต์ (ต่ำ)
    "\uf710": "ั",  # ไม้หันอากาศ
    "\uf711": "ํ",  # นิคหิต
    "\uf712": "็",  # ไม้ไต่คู้
    **{chr(0xF713 + i): c for i, c in enumerate("่้๊๋์")},  # วรรณยุกต์/การันต์ (ซ้าย)
    "\uf718": "ุ",  # สระอุ
    "\uf719": "ู",  # สระอู
    "\uf71a": "ฺ",  # พินทุ
    "\uf70f": "\u0e0d",  # ญ ไม่มีเชิง
}
_PUA_RE = re.compile("[" + "".join(PUA) + "]")

# CT_Settings: element ที่อยู่ก่อน hideSpellingErrors / hideGrammaticalErrors
SETTINGS_EARLY = {
    "writeProtection", "view", "zoom", "removePersonalInformation", "removeDateAndTime",
    "doNotDisplayPageBoundaries", "displayBackgroundShape", "printPostScriptOverText",
    "printFractionalCharacterWidth", "printFormsData", "embedTrueTypeFonts", "embedSystemFonts",
    "saveSubsetFonts", "saveFormsData", "mirrorMargins", "alignBordersAndEdges",
    "bordersDoNotSurroundHeader", "bordersDoNotSurroundFooter", "gutterAtTop",
}

_XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


def _text(el) -> str:
    return "".join(t.text or "" for t in el.iter(qn("w:t")))


# ---------- 1) อักษรเพี้ยน ----------
def normalize_chars(doc) -> int:
    """U+F70A–F71A -> อักษรไทยมาตรฐาน และ นิคหิต+า (ํา) -> สระอำ คืนจำนวน w:t ที่เปลี่ยน"""
    n = 0
    for root in iter_content_parts(doc):
        for t in root.iter(qn("w:t")):
            s = t.text or ""
            new = _PUA_RE.sub(lambda m: PUA[m.group()], s).replace("ํา", "ำ")
            if new != s:
                t.text = new
                n += 1
    return n


# ---------- 2) ฟอนต์ / ขนาด ----------
def _set_font(rf, font: str) -> bool:
    changed = False
    for a in THEME_ATTRS:
        if rf.get(qn(a)) is not None:
            del rf.attrib[qn(a)]
            changed = True
    for a in FONT_ATTRS:
        if rf.get(qn(a)) != font:
            rf.set(qn(a), font)
            changed = True
    return changed


def _is_symbol(rf) -> bool:
    return any(rf.get(qn(a)) in SYMBOL_FONTS for a in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"))


def set_font_all(doc, font: str) -> None:
    """rFonts ascii/hAnsi/cs/eastAsia ทุก run + docDefaults + styles + numbering ยกเว้นฟอนต์สัญลักษณ์
    ลบ attribute theme; run ที่มีข้อความแต่ไม่มี rFonts จะได้ rFonts ตรงๆ"""
    for root in iter_parts(doc):
        for rf in root.iter(qn("w:rFonts")):
            if not _is_symbol(rf):
                _set_font(rf, font)
    for root in iter_content_parts(doc):
        for r in root.iter(qn("w:r")):
            if r.find(qn("w:t")) is None:
                continue
            rpr = r.find(qn("w:rPr"))
            if rpr is None:
                rpr = OxmlElement("w:rPr")
                r.insert(0, rpr)
            if rpr.find(qn("w:rFonts")) is None:
                _set_font(child(rpr, "w:rFonts"), font)
                reorder(rpr, RPR_ORDER)


def align_sz(doc) -> int:
    """sz = szCs (szCs เป็นตัวจริงของภาษาไทย; มีแค่ sz ก็สร้าง szCs ให้) และ b <-> bCs
    คืนจำนวน rPr ที่แก้"""
    n = 0
    for root in iter_parts(doc):
        for rpr in root.iter(qn("w:rPr")):
            changed = False
            sz, cs = rpr.find(qn("w:sz")), rpr.find(qn("w:szCs"))
            if cs is not None and cs.get(qn("w:val")) is not None:
                if sz is None or sz.get(qn("w:val")) != cs.get(qn("w:val")):
                    child(rpr, "w:sz").set(qn("w:val"), cs.get(qn("w:val")))
                    changed = True
            elif sz is not None and sz.get(qn("w:val")) is not None:
                child(rpr, "w:szCs").set(qn("w:val"), sz.get(qn("w:val")))
                changed = True
            for a, b in (("w:b", "w:bCs"), ("w:bCs", "w:b")):
                src = rpr.find(qn(a))
                if src is not None and rpr.find(qn(b)) is None:
                    dst = OxmlElement(b)
                    if src.get(qn("w:val")) is not None:
                        dst.set(qn("w:val"), src.get(qn("w:val")))
                    rpr.append(dst)
                    changed = True
            if changed:
                reorder(rpr, RPR_ORDER)
                n += 1
    return n


# ---------- 3) ขึ้นบรรทัดด้วยมือ ----------
_KEEP_BREAK = ("- ", "– ", "• ")


def remove_manual_breaks(doc) -> int:
    """ลบ <w:br/> ธรรมดา (ไม่มี type หรือ textWrapping) ยกเว้นตัวที่ตามด้วยข้อความขึ้นต้น "- "
    (คั่นรายการ) ถ้าสองข้างเป็นคำอังกฤษ/ตัวเลขติดกันจะแทนด้วยช่องว่างหนึ่งตัว (กันคำติดกัน)
    คืนจำนวน br ที่ลบ"""
    removed = 0
    for root in iter_content_parts(doc):
        for p in root.iter(qn("w:p")):
            seq = [e for e in p.iter(qn("w:t"), qn("w:br"))
                   if e.tag == qn("w:t") or e.get(qn("w:type")) in (None, "textWrapping")]
            for i, e in enumerate(seq):
                if e.tag != qn("w:br"):
                    continue
                before = "".join(x.text or "" for x in seq[:i] if x.tag == qn("w:t"))
                after = "".join(x.text or "" for x in seq[i + 1:] if x.tag == qn("w:t"))
                if after.lstrip().startswith(_KEEP_BREAK):
                    continue
                prev_c, next_c = before[-1:], after[:1]
                glue = (prev_c and next_c and not prev_c.isspace() and not next_c.isspace()
                        and not (THAI.match(prev_c) and THAI.match(next_c)))
                if glue:
                    t = OxmlElement("w:t")
                    t.set(_XML_SPACE, "preserve")
                    t.text = " "
                    e.getparent().replace(e, t)
                else:
                    e.getparent().remove(e)
                removed += 1
    return removed


# ---------- สไตล์ ----------
def _style_map(doc) -> dict[str, dict]:
    """id -> {name, based, jc, outline}"""
    out: dict[str, dict] = {}
    for st in doc.styles.element.findall(qn("w:style")):
        sid = st.get(qn("w:styleId"))
        name = st.find(qn("w:name"))
        based = st.find(qn("w:basedOn"))
        jc = st.find(f"{qn('w:pPr')}/{qn('w:jc')}")
        ol = st.find(f"{qn('w:pPr')}/{qn('w:outlineLvl')}")
        out[sid] = {
            "name": (name.get(qn("w:val")) if name is not None else "") or "",
            "based": based.get(qn("w:val")) if based is not None else None,
            "jc": jc.get(qn("w:val")) if jc is not None else None,
            "outline": ol is not None,
        }
    return out


def _chain(styles: dict[str, dict], sid: str | None):
    seen = set()
    while sid and sid in styles and sid not in seen:
        seen.add(sid)
        yield sid, styles[sid]
        sid = styles[sid]["based"]


_HEADING = re.compile(r"^(heading|หัวเรื่อง)", re.I)
_TOC = re.compile(r"^(toc\s*\d*|table of contents|สารบัญ)", re.I)
_CAPTION = re.compile(r"^(caption|คำอธิบายภาพ)", re.I)


def _p_style_id(p) -> str | None:
    st = p.find(f"{qn('w:pPr')}/{qn('w:pStyle')}")
    return st.get(qn("w:val")) if st is not None else None


def _is_heading(p, styles) -> bool:
    if p.find(f"{qn('w:pPr')}/{qn('w:outlineLvl')}") is not None:
        return True
    for sid, info in _chain(styles, _p_style_id(p)):
        if _HEADING.match(sid) or _HEADING.match(info["name"]) or info["outline"]:
            return True
    return False


def _has_style(p, styles, rx) -> bool:
    sid = _p_style_id(p)
    if sid and rx.match(sid):
        return True
    return any(rx.match(i["name"]) for _, i in _chain(styles, sid))


def _effective_jc(p, styles) -> str | None:
    jc = p.find(f"{qn('w:pPr')}/{qn('w:jc')}")
    if jc is not None:
        return jc.get(qn("w:val"))
    for _, info in _chain(styles, _p_style_id(p)):
        if info["jc"]:
            return info["jc"]
    return None


def _set_ppr(p, tag: str, val: str | None = None):
    ppr = p.find(qn("w:pPr"))
    if ppr is None:
        ppr = OxmlElement("w:pPr")
        p.insert(0, ppr)
    el = child(ppr, tag)
    if val is not None:
        el.set(qn("w:val"), val)
    reorder(ppr, PPR_ORDER)
    return el


# ---------- 4) จัดบรรทัด / หัวข้อ ----------
_URL = re.compile(r"https?://|www\.", re.I)


def distribute(doc) -> int:
    """jc=thaiDistribute ให้ย่อหน้าเนื้อความ (เนื้อหาหลักรวมตาราง) ที่ jc ที่ใช้จริงเป็น both/left/ไม่มี
    ข้าม heading (สไตล์ Heading*/หัวเรื่อง* หรือ outlineLvl), TOC, Caption, ย่อหน้าที่มี URL, ย่อหน้าว่าง"""
    styles = _style_map(doc)
    n = 0
    for p in doc.element.body.iter(qn("w:p")):
        text = _text(p).strip()
        if not text or _URL.search(text):
            continue
        if _is_heading(p, styles) or _has_style(p, styles, _TOC) or _has_style(p, styles, _CAPTION):
            continue
        if _effective_jc(p, styles) not in (None, "both", "left"):
            continue
        _set_ppr(p, "w:jc", "thaiDistribute")
        n += 1
    return n


_HEAD_TEXT = re.compile(r"^(ส่วนที่|บทที่|ตารางที่|[0-9๐-๙]+(\.[0-9๐-๙]+)+\.?\s)")


def _is_heading_para(p, styles, text: str, maxlen: int = 90) -> bool:
    """ตัวตรวจหัวข้อกลาง (ใช้ร่วมกันของ keep_next_headings และ paragraph_spacing):
    สไตล์ heading/outlineLvl หรือข้อความสั้นขึ้นต้น ส่วนที่/บทที่/ตารางที่/เลขหลายระดับ"""
    return _is_heading(p, styles) or (len(text) <= maxlen and bool(_HEAD_TEXT.match(text)))


def keep_next_headings(doc, maxlen: int = 90) -> int:
    """keepNext ให้หัวข้อ (สไตล์ heading/outlineLvl หรือข้อความขึ้นต้น ส่วนที่/บทที่/ตารางที่/เลขหลายระดับ
    เช่น 1.1) คืนจำนวนย่อหน้าที่เพิ่ม keepNext (เลขชั้นเดียว "1. " ไม่นับ เพราะเป็นรายการ)"""
    styles = _style_map(doc)
    n = 0
    for p in doc.element.body.iter(qn("w:p")):
        text = _text(p).strip()
        if not text:
            continue
        is_head = _is_heading_para(p, styles, text, maxlen)
        if not is_head or p.find(f"{qn('w:pPr')}/{qn('w:keepNext')}") is not None:
            continue
        _set_ppr(p, "w:keepNext")
        n += 1
    return n


# ---------- 4b) ระยะบรรทัดเดี่ยว / ระยะก่อนย่อหน้าตามชนิดเอกสาร ----------
def _ppr_spacings(doc):
    """w:spacing ที่อยู่ใน pPr ของทุก part (ย่อหน้า สไตล์ docDefaults numbering) — ไม่รวม spacing ใน rPr (ระยะอักษร)"""
    for root in iter_parts(doc):
        for sp in root.iter(qn("w:spacing")):
            par = sp.getparent()
            if par is not None and local(par) == "pPr":
                yield sp


def _is_exact(sp) -> bool:
    return sp.get(qn("w:lineRule")) == "exact"


def single_spacing(doc) -> int:
    """บังคับ line=240 lineRule=auto after=0 ที่ทุก w:spacing ใน pPr ของทุก part ที่ตั้ง line หรือ after ไว้
    (ย่อหน้า สไตล์ docDefaults) ข้าม lineRule=exact (ล็อกความสูงโดยตั้งใจ) ไม่ลบ before ฯลฯ
    ลบ afterAutospacing/afterLines ที่ขัดกับ after=0 คืนจำนวน element ที่เปลี่ยน"""
    n = 0
    for sp in _ppr_spacings(doc):
        if _is_exact(sp):
            continue
        if sp.get(qn("w:line")) is None and sp.get(qn("w:after")) is None:
            continue
        want = {"w:line": "240", "w:lineRule": "auto", "w:after": "0"}
        changed = False
        for k, v in want.items():
            if sp.get(qn(k)) != v:
                sp.set(qn(k), v)
                changed = True
        for k in ("w:afterAutospacing", "w:afterLines"):
            if sp.get(qn(k)) is not None:
                del sp.attrib[qn(k)]
                changed = True
        n += changed
    return n


def count_exact_spacing(doc) -> int:
    """จำนวน w:spacing ที่ lineRule=exact (ที่ single_spacing คงไว้)"""
    return sum(1 for sp in _ppr_spacings(doc) if _is_exact(sp))


_DOC_TYPES = {"ราชการ": "official", "official": "official", "รายงาน": "report", "report": "report"}


_NUM_HEAD = re.compile(r"^[0-9๐-๙]+\.\s")


def _bold_styles(doc) -> dict[str, tuple]:
    """styleId -> (basedOn, bold|None) จาก rPr ของสไตล์"""
    out = {}
    for st in doc.styles.element.findall(qn("w:style")):
        based = st.find(qn("w:basedOn"))
        out[st.get(qn("w:styleId"))] = (based.get(qn("w:val")) if based is not None else None,
                                         _bold_of(st.find(qn("w:rPr"))))
    return out


def _bold_of(rpr):
    """True/False/None(ไม่ได้ระบุ) จาก w:b (ถ้าไม่มีใช้ w:bCs)"""
    if rpr is None:
        return None
    for tag in ("w:b", "w:bCs"):
        el = rpr.find(qn(tag))
        if el is not None:
            return el.get(qn("w:val"), "1").lower() in ("1", "true", "on")
    return None


def _style_bold(bstyles, sid):
    seen = set()
    while sid and sid in bstyles and sid not in seen:
        seen.add(sid)
        based, b = bstyles[sid]
        if b is not None:
            return b
        sid = based
    return None


def _is_bold_para(p, bstyles) -> bool:
    """ทุก run ที่มีข้อความไม่ว่างเป็นตัวหนา: rPr -> สไตล์อักขระ -> สไตล์ย่อหน้า -> ไม่หนา"""
    psid = _p_style_id(p)
    found = False
    for r in p.iter(qn("w:r")):
        if not "".join(t.text or "" for t in r.iter(qn("w:t"))).strip():
            continue
        found = True
        rpr = r.find(qn("w:rPr"))
        b = _bold_of(rpr)
        if b is None and rpr is not None:
            rs = rpr.find(qn("w:rStyle"))
            b = _style_bold(bstyles, rs.get(qn("w:val"))) if rs is not None else None
        if b is None:
            b = _style_bold(bstyles, psid)
        if not b:
            return False
    return found


def _report_head_text(p, text, bstyles) -> bool:
    """หัวข้อจากข้อความในโหมดรายงาน: ส่วนที่/บทที่/ตารางที่ เป็นหัวข้อเสมอ; ขึ้นต้นเลข (1. / 1.1) ต้องเป็นตัวหนา"""
    if _NUM_HEAD.match(text) or (_HEAD_TEXT.match(text) and text[0] in "0123456789๐๑๒๓๔๕๖๗๘๙"):
        return _is_bold_para(p, bstyles)
    return bool(_HEAD_TEXT.match(text))
_REF = re.compile(r"^อ้างถึง(\s|$)")
_ENCL = re.compile(r"^สิ่งที่ส่งมาด้วย(\s|$)")
_ITEM = re.compile(r"^([0-9๐-๙]+\.\s|\(?[0-9๐-๙]+\)\s)")
_DEAR = re.compile(r"^เรียน(\s|$)")


def _set_before(p, val: str) -> bool:
    sp = p.find(f"{qn('w:pPr')}/{qn('w:spacing')}")
    if sp is not None and sp.get(qn("w:before")) == val and sp.get(qn("w:beforeAutospacing")) is None \
            and sp.get(qn("w:beforeLines")) is None:
        return False
    sp = _set_ppr(p, "w:spacing")
    sp.set(qn("w:before"), val)
    for k in ("w:beforeAutospacing", "w:beforeLines"):
        if sp.get(qn(k)) is not None:
            del sp.attrib[qn(k)]
    return True


def paragraph_spacing(doc, doc_type: str | None) -> int:
    """ระยะก่อนย่อหน้า 6 pt (before=120) ตามชนิดเอกสาร ทำเฉพาะย่อหน้าใน body ที่ไม่อยู่ในตาราง
      ราชการ/official: ย่อหน้าเนื้อความ (ไม่ว่าง ไม่ใช่หัวข้อ jc ไม่ใช่ center/right) ระหว่างย่อหน้า "เรียน"
        (ถ้าไม่มีถือว่าไม่มีส่วนหัว) กับส่วนลงชื่อ (เริ่มที่ "ขอแสดงความนับถือ" หรือย่อหน้า center/right
        แรกหลังเนื้อความ)
      รายงาน/report: หัวข้อ before=120 เนื้อความ before=0
    คืนจำนวนย่อหน้าที่เปลี่ยน; doc_type=None คืน 0; ค่าอื่น ValueError"""
    if doc_type is None:
        return 0
    kind = _DOC_TYPES.get(doc_type.strip().lower() if isinstance(doc_type, str) else doc_type)
    if kind is None:
        raise ValueError(f"ไม่รู้จักชนิดเอกสาร {doc_type!r} ใช้ได้เฉพาะ ราชการ (official) หรือ รายงาน (report)")
    styles = _style_map(doc)
    bstyles = _bold_styles(doc) if kind == "report" else {}
    tbl = qn("w:tbl")
    paras = []
    for p in doc.element.body.iter(qn("w:p")):
        if any(a.tag == tbl for a in p.iterancestors()):
            continue
        text = _text(p).strip()
        if kind == "report":  # รายงาน: นับเลขชั้นเดียว "1. " เป็นหัวข้อด้วย (เฉพาะโหมดนี้)
            head = bool(text) and (_is_heading(p, styles) or (len(text) <= 90 and _report_head_text(p, text, bstyles)))
        else:  # ราชการ: ข้ามเฉพาะหัวข้อจริง (สไตล์ Heading/outlineLvl) ไม่ใช่ข้อความขึ้นต้นเลข/ตารางที่
            head = bool(text) and _is_heading(p, styles)
        paras.append((p, text, head, _effective_jc(p, styles) in ("center", "right")))
    n = 0
    if kind == "report":
        for p, text, head, _ in paras:
            if text:
                n += _set_before(p, "120" if head else "0")
        return n
    start = next((i + 1 for i, (_, t, _, _) in enumerate(paras) if _DEAR.match(t)), 0)
    if start:  # ส่วนหัวต่อด้วย อ้างถึง / สิ่งที่ส่งมาด้วย (และรายการเลขใต้ข้อนั้น) ชิดใต้ "เรียน"
        mode = None
        while start < len(paras):
            text = paras[start][1]
            if not text:
                pass
            elif _REF.match(text) or _ENCL.match(text):
                mode = True
            elif not (mode and _ITEM.match(text)):
                break
            start += 1
    sig = None
    for i in range(start, len(paras)):
        if paras[i][1].startswith("ขอแสดงความนับถือ"):
            sig = i
            break
    if sig is None:
        seen_body = False
        for i in range(start, len(paras)):
            _, text, head, aligned = paras[i]
            if aligned and seen_body:
                sig = i
                break
            if text and not head and not aligned:
                seen_body = True
    end = len(paras) if sig is None else sig
    for p, text, head, aligned in paras[start:end]:
        if text and not head and not aligned:
            n += _set_before(p, "120")
    return n


# ---------- 5) ภาษาไทยของ run ----------
def _set_lang(rpr) -> bool:
    lang = rpr.find(qn("w:lang"))
    if lang is None:
        lang = OxmlElement("w:lang")
        rpr.append(lang)
    changed = lang.get(qn("w:val")) != "en-US" or lang.get(qn("w:bidi")) != "th-TH"
    lang.set(qn("w:val"), "en-US")
    lang.set(qn("w:bidi"), "th-TH")
    return changed


def thai_runs(doc) -> int:
    """run ที่มี w:t ซึ่งมีอักษร U+0E00–U+0E7F: เพิ่ม w:cs และ w:lang val=en-US bidi=th-TH
    คืนจำนวน run ที่แก้ (ไฟล์ที่ wrap แล้วต้อง unwrap ก่อน — ทำที่ CLI ไม่ใช่ที่นี่)"""
    n = 0
    for root in iter_content_parts(doc):
        for r in root.iter(qn("w:r")):
            if not THAI.search(_text(r)):
                continue
            rpr = r.find(qn("w:rPr"))
            changed = False
            if rpr is None:
                rpr = OxmlElement("w:rPr")
                r.insert(0, rpr)
                changed = True
            changed |= _set_lang(rpr)
            if rpr.find(qn("w:cs")) is None:
                rpr.append(OxmlElement("w:cs"))
                changed = True
            reorder(rpr, RPR_ORDER)
            n += changed
    return n


def doc_defaults(doc, font: str, body_pt: float) -> None:
    """docDefaults: ฟอนต์/ขนาด/ภาษาไทย เสมอ; ระยะย่อหน้า single (line 240 auto, before/after 0) เสมอ; สไตล์ที่มี lang ได้ bidi=th-TH ด้วย"""
    styles = doc.styles.element
    defaults = styles.find(qn("w:docDefaults"))
    if defaults is None:
        defaults = OxmlElement("w:docDefaults")
        styles.insert(0, defaults)
    rpr_def = child(defaults, "w:rPrDefault")
    rpr = child(rpr_def, "w:rPr")
    _set_font(child(rpr, "w:rFonts"), font)
    size = str(int(round(body_pt * 2)))
    for tag in ("w:sz", "w:szCs"):
        child(rpr, tag).set(qn("w:val"), size)
    _set_lang(rpr)
    reorder(rpr, RPR_ORDER)
    ppr_def = child(defaults, "w:pPrDefault")
    ppr = child(ppr_def, "w:pPr")
    sp = child(ppr, "w:spacing")  # single เสมอ (ผู้ใช้สั่ง): line 240 auto, after 0
    for k, v in (("w:before", "0"), ("w:after", "0"), ("w:line", "240"), ("w:lineRule", "auto")):
        sp.set(qn(k), v)
    for k in ("w:beforeAutospacing", "w:afterAutospacing"):
        if sp.get(qn(k)) is not None:
            del sp.attrib[qn(k)]
    reorder(ppr, PPR_ORDER)
    reorder(defaults, ["rPrDefault", "pPrDefault"])
    for st_rpr in styles.iter(qn("w:rPr")):
        if st_rpr.find(qn("w:lang")) is not None:
            _set_lang(st_rpr)
        reorder(st_rpr, RPR_ORDER)


def hide_spelling(doc) -> None:
    """ซ่อนเส้นตรวจคำ: แทรก hideSpellingErrors, hideGrammaticalErrors ใน settings.xml ตามลำดับ schema
    (ต่อจาก element กลุ่มต้นไฟล์ที่มี) ทำซ้ำได้"""
    st = doc.settings.element
    zoom = st.find(qn("w:zoom"))  # แม่แบบ python-docx ไม่มี percent (schema บังคับ)
    if zoom is not None and zoom.get(qn("w:percent")) is None:
        zoom.set(qn("w:percent"), "100")
    for tag in ("w:hideSpellingErrors", "w:hideGrammaticalErrors"):
        if st.find(qn(tag)) is not None:
            continue
        keep = SETTINGS_EARLY | {"hideSpellingErrors"}
        pos = max([i + 1 for i, k in enumerate(st) if isinstance(k.tag, str) and local(k) in keep] or [0])
        st.insert(pos, OxmlElement(tag))


def finalize(doc, font: str = "TH Sarabun New", body_pt: float = 16.0, doc_type: str | None = None) -> None:
    """เรียกครั้งเดียวหลังเนื้อหาครบ: set_font_all, align_sz, keep_next_headings, single_spacing,
    paragraph_spacing(doc_type), doc_defaults, thai_runs, hide_spelling ตามลำดับ แล้วเรียง rPr/pPr ของทุก part ตาม schema"""
    set_font_all(doc, font)
    align_sz(doc)
    keep_next_headings(doc)
    paragraph_spacing(doc, doc_type)
    single_spacing(doc)
    doc_defaults(doc, font, body_pt)
    thai_runs(doc)
    hide_spelling(doc)
    for root in iter_parts(doc):
        for rpr in root.iter(qn("w:rPr")):
            reorder(rpr, RPR_ORDER)
        for ppr in root.iter(qn("w:pPr")):
            reorder(ppr, PPR_ORDER)
