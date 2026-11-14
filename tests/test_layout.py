from __future__ import annotations

import re

import pytest
from docx import Document
from docx.oxml.ns import qn
from lxml import etree

from thaidocs import layout, ooxml

THAI = re.compile("[฀-๿]")
LONG = "ข้อความภาษาไทยสำหรับทดสอบการจัดหน้าแบบกระจายบรรทัด"


def _runs_with_thai(doc):
    for root in ooxml.iter_content_parts(doc):
        for r in root.iter(qn("w:r")):
            if THAI.search("".join(t.text or "" for t in r.iter(qn("w:t")))):
                yield r


def _snapshot(doc):
    parts = [etree.tostring(r) for r in ooxml.iter_parts(doc)]
    parts.append(etree.tostring(doc.settings.element))
    return parts


def test_thai_runs_marked(make_doc):
    doc = make_doc([LONG, "english only"], table=True, header="หัวกระดาษ")
    layout.finalize(doc)
    runs = list(_runs_with_thai(doc))
    assert runs
    for r in runs:
        rpr = r.find(qn("w:rPr"))
        assert rpr.find(qn("w:cs")) is not None
        assert rpr.find(qn("w:lang")).get(qn("w:bidi")) == "th-TH"
    # run อังกฤษล้วนไม่ถูกติดป้าย
    eng = [r for r in doc.element.body.iter(qn("w:r"))
           if "".join(t.text or "" for t in r.iter(qn("w:t"))) == "english only"][0]
    rpr = eng.find(qn("w:rPr"))
    assert rpr is None or rpr.find(qn("w:cs")) is None


def test_sz_equals_szcs(make_doc):
    doc = make_doc(["ก", "ข"])
    p0, p1 = doc.paragraphs
    r0 = p0.runs[0]
    r0.font.size = 14 * 12700  # sz=28 เท่านั้น
    r1 = p1.runs[0]
    r1.font.bold = True
    assert layout.align_sz(doc) >= 1
    rpr0 = r0._r.rPr
    assert rpr0.find(qn("w:szCs")).get(qn("w:val")) == "28"
    assert rpr0.find(qn("w:sz")).get(qn("w:val")) == "28"
    rpr1 = r1._r.rPr
    assert rpr1.find(qn("w:b")) is not None and rpr1.find(qn("w:bCs")) is not None
    assert layout.align_sz(doc) == 0  # idempotent


def test_font_all_no_theme(make_doc):
    doc = make_doc([LONG], table=True, header="หัว", footnote=True)
    doc.paragraphs[0].runs[0].font.name = "Calibri"
    layout.set_font_all(doc, "TH Sarabun New")
    styles = etree.tostring(doc.styles.element).decode()
    for a in ("asciiTheme", "hAnsiTheme", "cstheme", "eastAsiaTheme"):
        assert a not in styles
    seen = 0
    for root in ooxml.iter_parts(doc):
        for rf in root.iter(qn("w:rFonts")):
            if rf.get(qn("w:ascii")) in layout.SYMBOL_FONTS:  # template มี Symbol ในสไตล์ bullet
                continue
            seen += 1
            assert rf.get(qn("w:ascii")) == "TH Sarabun New"
            for a in ("hAnsi", "cs", "eastAsia"):
                assert rf.get(qn("w:" + a)) == "TH Sarabun New"
    assert seen > 3


def test_font_all_keeps_symbol_fonts(make_doc):
    doc = make_doc([LONG])
    doc.paragraphs[0].runs[0].font.name = "Symbol"
    layout.set_font_all(doc, "TH Sarabun New")
    rf = doc.paragraphs[0].runs[0]._r.rPr.find(qn("w:rFonts"))
    assert rf.get(qn("w:ascii")) == "Symbol"


def test_rpr_schema_order(make_doc):
    doc = make_doc([LONG], table=True, header="หัว", footnote=True)
    r = doc.paragraphs[0].runs[0]
    r.font.bold = True
    r.font.size = 14 * 12700
    r.font.italic = True
    layout.finalize(doc)
    n = 0
    for root in ooxml.iter_parts(doc):
        for rpr in root.iter(qn("w:rPr")):
            names = [etree.QName(k).localname for k in rpr if etree.QName(k).localname in ooxml.RPR_ORDER]
            idx = [ooxml.RPR_ORDER.index(x) for x in names]
            assert idx == sorted(idx), names
            n += 1
    assert n > 5


def test_finalize_all_parts(make_doc):
    doc = make_doc(["ข้อความ"], table=True, header="หัวกระดาษไทย", footnote=True)
    layout.finalize(doc)
    texts = []
    for root in ooxml.iter_content_parts(doc):
        for r in root.iter(qn("w:r")):
            t = "".join(x.text or "" for x in r.iter(qn("w:t")))
            if THAI.search(t):
                texts.append(t)
                assert r.find(qn("w:rPr")).find(qn("w:cs")) is not None, t
    joined = "|".join(texts)
    assert "หัวกระดาษไทย" in joined and "เชิงอรรถ" in joined and "เซลล์ตาราง" in joined


def test_iter_parts_covers_footnotes_and_survives_save(make_doc, tmp_path):
    doc = make_doc(["ก"], header="ข", footnote=True)
    layout.finalize(doc)
    out = tmp_path / "out.docx"
    doc.save(out)
    re_doc = Document(str(out))
    found = [e for e in ooxml.iter_content_parts(re_doc)
             if e.tag in (qn("w:footnotes"), qn("w:hdr"), qn("w:ftr"))]
    assert {etree.QName(e).localname for e in found} >= {"footnotes", "hdr"}
    fn = [e for e in found if e.tag == qn("w:footnotes")][0]
    assert fn.find(".//" + qn("w:cs")) is not None


def test_normalize_chars(make_doc):
    doc = make_doc(["นํา", "กขคง"])
    n = layout.normalize_chars(doc)
    assert n == 2
    assert doc.paragraphs[0].text == "นำ"
    assert doc.paragraphs[1].text == "ก่ข่คังุ"
    assert layout.normalize_chars(doc) == 0


def test_distribute_skips_headings_and_urls(make_doc):
    doc = make_doc([
        ("บทนำของรายงาน", "Heading 1"),
        "ดูที่ https://x.org",
        "ดูที่ www.example.org เพิ่มเติม",
        LONG,
    ])
    n = layout.distribute(doc)
    jcs = [p._p.pPr.find(qn("w:jc")) if p._p.pPr is not None else None for p in doc.paragraphs]
    assert jcs[0] is None and jcs[1] is None and jcs[2] is None
    assert jcs[3].get(qn("w:val")) == "thaiDistribute"
    assert n == 1
    assert layout.distribute(doc) == 0


def test_distribute_skips_aligned_and_toc_caption(make_doc):
    doc = make_doc([("ศูนย์กลาง", None), ("สารบัญ", "TOC 1"), ("รูปที่ ๑", "Caption"), LONG])
    doc.paragraphs[0].alignment = 1  # center
    layout.distribute(doc)
    jcs = [p._p.pPr.find(qn("w:jc")) if p._p.pPr is not None else None for p in doc.paragraphs]
    assert jcs[0].get(qn("w:val")) == "center"
    assert jcs[1] is None or jcs[1].get(qn("w:val")) != "thaiDistribute"
    assert jcs[2] is None or jcs[2].get(qn("w:val")) != "thaiDistribute"
    assert jcs[3].get(qn("w:val")) == "thaiDistribute"


def test_remove_manual_breaks(make_doc):
    doc = make_doc(["x"])
    p = doc.paragraphs[0]
    p.clear()
    p.add_run("ทดสอบ").add_break()
    p.add_run("ข้อความ")
    p.add_run("hello").add_break()
    p.add_run("world")
    p.add_run("รายการ").add_break()
    p.add_run("- ข้อย่อย")
    from docx.enum.text import WD_BREAK
    p.add_run("หน้า").add_break(WD_BREAK.PAGE)
    assert layout.remove_manual_breaks(doc) == 2
    text = p.text
    assert "ทดสอบข้อความ" in text and "hello world" in text
    assert text.count("\n") == 1  # br ก่อน "- " ยังอยู่
    assert len(p._p.findall(".//" + qn("w:br"))) == 2  # "- " + page
    assert layout.remove_manual_breaks(doc) == 0


def test_keep_next_headings(make_doc):
    doc = make_doc([("บทนำ", "Heading 1"), "บทที่ ๑ ความนำ", "1.1 ที่มา", LONG])
    assert layout.keep_next_headings(doc) == 3
    flags = [p._p.pPr is not None and p._p.pPr.find(qn("w:keepNext")) is not None for p in doc.paragraphs]
    assert flags == [True, True, True, False]
    assert layout.keep_next_headings(doc) == 0


def test_hide_spelling(make_doc):
    doc = make_doc([LONG])
    layout.hide_spelling(doc)
    layout.hide_spelling(doc)
    st = doc.settings.element
    names = [etree.QName(k).localname for k in st]
    assert names.count("hideSpellingErrors") == 1 and names.count("hideGrammaticalErrors") == 1
    i = names.index("hideSpellingErrors")
    assert names[i + 1] == "hideGrammaticalErrors"
    early = {"writeProtection", "view", "zoom", "removePersonalInformation", "removeDateAndTime",
             "doNotDisplayPageBoundaries", "displayBackgroundShape", "printPostScriptOverText",
             "printFractionalCharacterWidth", "printFormsData", "embedTrueTypeFonts", "embedSystemFonts",
             "saveSubsetFonts", "saveFormsData", "mirrorMargins", "alignBordersAndEdges",
             "bordersDoNotSurroundHeader", "bordersDoNotSurroundFooter", "gutterAtTop"}
    assert all(n in early for n in names[:i])
    assert all(n not in early for n in names[i + 2:])


def test_hide_spelling_after_existing_early_elements(make_doc):
    doc = make_doc([LONG])
    st = doc.settings.element
    # ใส่ element ต้นไฟล์ที่ไม่มีใน template: ต้องแทรกต่อท้ายกลุ่มนี้ ไม่ใช่ที่ต้นไฟล์
    from docx.oxml import OxmlElement
    st.insert(0, OxmlElement("w:mirrorMargins"))
    st.insert(0, OxmlElement("w:view"))
    layout.hide_spelling(doc)
    names = [etree.QName(k).localname for k in st]
    i = names.index("hideSpellingErrors")
    assert names.index("view") < names.index("mirrorMargins") < i
    assert names[i + 1] == "hideGrammaticalErrors"
    # ไม่มี element กลุ่มต้นไฟล์ตามหลังสองตัวนี้
    assert all(n not in layout.SETTINGS_EARLY for n in names[i:])


def test_doc_defaults(make_doc):
    doc = make_doc([LONG])
    layout.doc_defaults(doc, "TH Sarabun New", 16)
    d = doc.styles.element.find(qn("w:docDefaults"))
    rpr = d.find(f"{qn('w:rPrDefault')}/{qn('w:rPr')}")
    assert rpr.find(qn("w:sz")).get(qn("w:val")) == "32"
    assert rpr.find(qn("w:szCs")).get(qn("w:val")) == "32"
    assert rpr.find(qn("w:rFonts")).get(qn("w:cs")) == "TH Sarabun New"
    assert rpr.find(qn("w:lang")).get(qn("w:bidi")) == "th-TH"


def test_finalize_idempotent(make_doc):
    doc = make_doc([LONG, ("บทนำ", "Heading 1")], table=True, header="หัว", footnote=True)
    layout.finalize(doc)
    a = _snapshot(doc)
    layout.finalize(doc)
    assert _snapshot(doc) == a


def test_pua_table_complete():
    for cp in range(0xF700, 0xF71B):
        c = chr(cp)
        assert c in layout.PUA, hex(cp)
        assert "\u0e00" <= layout.PUA[c] <= "\u0e7f", hex(cp)
    assert layout.PUA[chr(0xF70F)] == "\u0e0d"
    assert layout.PUA[chr(0xF700)] == "\u0e10"


def _sp(p):
    return p._p.find(f"{qn('w:pPr')}/{qn('w:spacing')}")


def _before(p):
    sp = _sp(p)
    return None if sp is None else sp.get(qn("w:before"))


def _set_spacing(el_ppr, **kw):
    from docx.oxml import OxmlElement
    sp = el_ppr.find(qn("w:spacing"))
    if sp is None:
        sp = OxmlElement("w:spacing")
        el_ppr.append(sp)
    for k, v in kw.items():
        sp.set(qn("w:" + k), v)
    ooxml.reorder(el_ppr, ooxml.PPR_ORDER)
    return sp


def _dd_spacing(doc):
    return doc.styles.element.find(f"{qn('w:docDefaults')}/{qn('w:pPrDefault')}/{qn('w:pPr')}/{qn('w:spacing')}")


def test_doc_defaults_always_single(make_doc):
    """เปลี่ยนตามผู้ใช้สั่ง (แทน Ruling I เดิมที่ไม่ทับ spacing ที่มีอยู่)"""
    doc = make_doc([LONG])
    layout.doc_defaults(doc, "TH Sarabun New", 16)
    sp = _dd_spacing(doc)
    sp.set(qn("w:after"), "200")
    sp.set(qn("w:line"), "276")
    layout.finalize(doc)
    sp = _dd_spacing(doc)
    assert (sp.get(qn("w:after")), sp.get(qn("w:line")), sp.get(qn("w:lineRule"))) == ("0", "240", "auto")


def _messy(make_doc):
    doc = make_doc([LONG, LONG, LONG], table=True, header="หัว", footnote=True)
    dd = doc.styles.element.find(f"{qn('w:docDefaults')}/{qn('w:pPrDefault')}")
    ppr = dd.find(qn("w:pPr")) if dd is not None else None
    if ppr is None:
        from docx.oxml import OxmlElement
        if dd is None:
            dd = OxmlElement("w:pPrDefault")
            doc.styles.element.find(qn("w:docDefaults")).append(dd)
        ppr = OxmlElement("w:pPr")
        dd.append(ppr)
    _set_spacing(ppr, after="200", line="276", lineRule="auto")
    normal = doc.styles["Normal"].element
    npr = normal.find(qn("w:pPr"))
    if npr is None:
        from docx.oxml import OxmlElement
        npr = OxmlElement("w:pPr")
        normal.append(npr)
    _set_spacing(npr, line="360", lineRule="auto")
    ps = doc.paragraphs
    _set_spacing(ps[0]._p.get_or_add_pPr(), line="480", lineRule="auto", after="240", afterAutospacing="1", before="60")
    _set_spacing(ps[1]._p.get_or_add_pPr(), line="300", lineRule="exact", after="100")
    cell_p = doc.tables[0].rows[0].cells[0].paragraphs[0]
    _set_spacing(cell_p._p.get_or_add_pPr(), line="400", after="120")
    hp = doc.sections[0].header.paragraphs[0]
    _set_spacing(hp._p.get_or_add_pPr(), line="500", after="80")
    return doc


def test_single_spacing_everywhere(make_doc):
    doc = _messy(make_doc)
    n = layout.single_spacing(doc)
    assert n >= 5
    sp = _dd_spacing(doc)
    assert (sp.get(qn("w:line")), sp.get(qn("w:lineRule")), sp.get(qn("w:after"))) == ("240", "auto", "0")
    nsp = doc.styles["Normal"].element.find(f"{qn('w:pPr')}/{qn('w:spacing')}")
    assert (nsp.get(qn("w:line")), nsp.get(qn("w:lineRule"))) == ("240", "auto")
    s0 = _sp(doc.paragraphs[0])
    assert (s0.get(qn("w:line")), s0.get(qn("w:lineRule")), s0.get(qn("w:after"))) == ("240", "auto", "0")
    assert s0.get(qn("w:before")) == "60"  # ไม่ลบ attribute อื่น
    assert s0.get(qn("w:afterAutospacing")) is None
    s1 = _sp(doc.paragraphs[1])
    assert (s1.get(qn("w:line")), s1.get(qn("w:lineRule")), s1.get(qn("w:after"))) == ("300", "exact", "100")
    c = _sp(doc.tables[0].rows[0].cells[0].paragraphs[0])
    assert (c.get(qn("w:line")), c.get(qn("w:after"))) == ("240", "0")
    h = _sp(doc.sections[0].header.paragraphs[0])
    assert (h.get(qn("w:line")), h.get(qn("w:after"))) == ("240", "0")
    assert layout.count_exact_spacing(doc) == 1
    assert layout.single_spacing(doc) == 0


def test_single_spacing_leaves_char_spacing(make_doc):
    from docx.oxml import OxmlElement
    doc = make_doc([LONG])
    rpr = doc.paragraphs[0].runs[0]._r.get_or_add_rPr()
    sp = OxmlElement("w:spacing")
    sp.set(qn("w:val"), "20")
    rpr.append(sp)
    layout.single_spacing(doc)
    assert sp.get(qn("w:line")) is None and sp.get(qn("w:after")) is None


def _letter(make_doc, with_dear=True):
    from docx.enum.text import WD_ALIGN_PARAGRAPH as A
    texts = ["ส่วนราชการ กองทดสอบ", "ที่ ทด 0001/1 วันที่ 1 มกราคม 2569", "เรื่อง ขอทดสอบระบบ"]
    if with_dear:
        texts.append("เรียน ผู้อำนวยการ")
    texts += ["ตามที่ได้รับแจ้งเรื่องการทดสอบ" + "ก" * 5, "รายละเอียดเพิ่มเติมตามที่แนบมา", "จึงเรียนมาเพื่อโปรดทราบ",
              "ขอแสดงความนับถือ", "นายทดสอบ ใจดี", "ผู้อำนวยการกอง"]
    doc = make_doc(texts, table=True)
    ps = doc.paragraphs
    k = len(texts)
    for p in ps[k - 3:k - 1]:
        p.alignment = A.CENTER
    ps[k - 3].alignment = A.CENTER
    return doc, ps, with_dear


@pytest.mark.parametrize("name", ["ราชการ", "official"])
def test_official_spacing(make_doc, name):
    doc, ps, _ = _letter(make_doc)
    n = layout.paragraph_spacing(doc, name)
    got = [_before(p) for p in ps]
    assert got[:4] == [None] * 4
    assert got[4:7] == ["120"] * 3
    assert got[7:] == [None] * 3
    assert n == 3
    tp = doc.tables[0].rows[0].cells[0].paragraphs[0]
    assert _sp(tp) is None or _sp(tp).get(qn("w:before")) is None
    assert layout.paragraph_spacing(doc, name) == 0


def test_official_without_dear(make_doc):
    doc, ps, _ = _letter(make_doc, with_dear=False)
    layout.paragraph_spacing(doc, "ราชการ")
    got = [_before(p) for p in ps]
    assert got[:3] == ["120"] * 3  # ไม่มี "เรียน" -> ไม่มีส่วนหัว
    assert got[3:6] == ["120"] * 3
    assert got[6:] == [None] * 3


def test_official_clears_before_autospacing(make_doc):
    doc, ps, _ = _letter(make_doc)
    _set_spacing(ps[4]._p.get_or_add_pPr(), beforeAutospacing="1", before="0")
    layout.paragraph_spacing(doc, "official")
    sp = _sp(ps[4])
    assert sp.get(qn("w:before")) == "120" and sp.get(qn("w:beforeAutospacing")) is None


def test_report_spacing(make_doc):
    doc = make_doc([("บทที่ 1 บทนำ", "Heading 1"), LONG, ("ที่มา", "Heading 2"), LONG, "1.1 วัตถุประสงค์", LONG], table=True)
    doc.paragraphs[4].runs[0].bold = True  # รอบ 3: เลขนำหน้าเป็นหัวข้อเมื่อหนาเท่านั้น
    n = layout.paragraph_spacing(doc, "รายงาน")
    ps = doc.paragraphs
    assert [_before(p) for p in ps] == ["120", "0", "120", "0", "120", "0"]
    assert n == 6
    tp = doc.tables[0].rows[0].cells[0].paragraphs[0]
    assert _sp(tp) is None or _sp(tp).get(qn("w:before")) is None
    assert layout.paragraph_spacing(doc, "report") == 0


def test_paragraph_spacing_none_and_invalid(make_doc):
    doc = make_doc([LONG])
    assert layout.paragraph_spacing(doc, None) == 0
    assert _sp(doc.paragraphs[0]) is None
    with pytest.raises(ValueError) as e:
        layout.paragraph_spacing(doc, "memo")
    assert THAI.search(str(e.value))


def test_finalize_doc_type(make_doc):
    doc = make_doc([("บทนำ", "Heading 1"), LONG])
    layout.finalize(doc, doc_type="report")
    assert [_before(p) for p in doc.paragraphs] == ["120", "0"]
    a = _snapshot(doc)
    layout.finalize(doc, doc_type="report")
    assert _snapshot(doc) == a
    with pytest.raises(ValueError):
        layout.finalize(make_doc([LONG]), doc_type="xx")


def _bold_all(p, val=None):
    for r in p.runs:
        r.bold = True if val is None else val


def test_report_single_level_numbered_headings(make_doc):
    """ผู้ใช้เปลี่ยนกติกา (รอบ 3): เลขนำหน้าเป็นหัวข้อเมื่อตัวหนาเท่านั้น"""
    doc = make_doc(["1. บทนำ", LONG, "๒. วิธีการ", LONG])
    for i in (0, 2):
        _bold_all(doc.paragraphs[i])
    layout.paragraph_spacing(doc, "report")
    assert [_before(p) for p in doc.paragraphs] == ["120", "0", "120", "0"]
    doc2 = make_doc(["1. บทนำ"])
    assert layout.keep_next_headings(doc2) == 0  # keep_next ไม่เปลี่ยน


def test_report_numbered_needs_bold(make_doc):
    from docx.enum.style import WD_STYLE_TYPE
    doc = make_doc(["1. บทนำ", "1.1 ความเป็นมา", "1. ศึกษาความชุกของโรค", "2. เปรียบเทียบผล", "3. หัวข้อสไตล์หนา",
                    ("ระเบียบวิธี", "Heading 2"), "4. ครึ่งหนา"])
    ps = doc.paragraphs
    _bold_all(ps[0]); _bold_all(ps[1])
    st = doc.styles.add_style("BoldP", WD_STYLE_TYPE.PARAGRAPH)
    st.font.bold = True
    ps[4].style = st
    ps[6].runs[0].bold = False
    layout.paragraph_spacing(doc, "report")
    assert [_before(p) for p in ps] == ["120", "120", "0", "0", "120", "120", "0"]


def test_report_bcs_only_counts_bold(make_doc):
    from docx.oxml import OxmlElement
    doc = make_doc(["2. วิธีการ"])
    rpr = doc.paragraphs[0].runs[0]._r.get_or_add_rPr()
    rpr.append(OxmlElement("w:bCs"))
    layout.paragraph_spacing(doc, "report")
    assert _before(doc.paragraphs[0]) == "120"


def test_official_numbered_lines_are_body(make_doc):
    doc = make_doc(["เรียน ผู้อำนวยการ", "1.1 รายการแรก", "ตารางที่ 1 สรุป", ("หัวข้อจริง", "Heading 1"), "จึงเรียนมาเพื่อทราบ"])
    layout.paragraph_spacing(doc, "official")
    assert [_before(p) for p in doc.paragraphs] == [None, "120", "120", None, "120"]


def test_official_dear_requires_word_boundary(make_doc):
    doc = make_doc(["เรียนรู้เรื่องใหม่", "เนื้อความ"])
    layout.paragraph_spacing(doc, "official")
    assert [_before(p) for p in doc.paragraphs] == ["120", "120"]  # ไม่มี "เรียน" จริง = ไม่มีส่วนหัว
    doc = make_doc(["เรียน", "เนื้อความ"])
    layout.paragraph_spacing(doc, "official")
    assert [_before(p) for p in doc.paragraphs] == [None, "120"]


def test_official_ref_and_enclosure_stay_tight(make_doc):
    doc = make_doc(["ที่ ทด 1/2569", "ถนนทดสอบ กรุงเทพ", "1 มกราคม 2569", "เรื่อง ขอทดสอบ", "เรียน ผู้อำนวยการ",
                    "อ้างถึง หนังสือเดิม", "สิ่งที่ส่งมาด้วย", "1. สำเนาคำสั่ง", "2. แบบฟอร์ม",
                    "ตามที่ได้รับแจ้ง", "จึงเรียนมาเพื่อทราบ", "ขอแสดงความนับถือ"])
    assert layout.paragraph_spacing(doc, "ราชการ") == 2
    assert [_before(p) for p in doc.paragraphs] == [None] * 9 + ["120", "120", None]


def test_official_ref_only(make_doc):
    doc = make_doc(["เรียน ผู้อำนวยการ", "อ้างถึง 1. หนังสือ", "(1) รายการ", "เนื้อความ"])
    layout.paragraph_spacing(doc, "official")
    assert [_before(p) for p in doc.paragraphs] == [None, None, None, "120"]


def test_official_numbered_body_without_enclosure(make_doc):
    doc = make_doc(["เรียน ผู้อำนวยการ", "1. ขอเรียนว่า", "เนื้อความ"])
    layout.paragraph_spacing(doc, "official")
    assert [_before(p) for p in doc.paragraphs] == [None, "120", "120"]
