import copy
import re
from dataclasses import replace
from pathlib import Path

import pytest
from docx import Document
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

from thaidocs.official import checker
from thaidocs.official.checker import (ALL_CODES, Finding, check_docx, check_document, summarize)
from thaidocs.official.fill import fill
from thaidocs.official.schema import DocType, load_example, parse_doc
from thaidocs.official.style import DEFAULT_SPEC

ALL_TYPES = [t.value for t in DocType]


def _set(data, dotted, value):
    keys = dotted.split(".")
    cur = data
    for k in keys[:-1]:
        cur = cur[int(k)] if isinstance(cur, list) else cur.setdefault(k, {})
    last = keys[-1]
    if isinstance(cur, list):
        cur[int(last)] = value
    else:
        cur[last] = value


def mk(doc_type, **over):
    data = copy.deepcopy(load_example(doc_type))
    for k, v in over.items():
        _set(data, k, v)
    return parse_doc(data)


def codes(doc, spec=DEFAULT_SPEC):
    return {f.code: f for f in check_document(doc, spec)}


# ------------------------------------------------------------------ ตัวอย่างต้องผ่าน
@pytest.mark.parametrize("t", ALL_TYPES)
def test_examples_after_fill_have_no_errors(t):
    d = fill(parse_doc(load_example(t)), None)
    errors = [f for f in check_document(d) if f.level == "error"]
    assert errors == []


@pytest.mark.parametrize("t", ALL_TYPES)
def test_rendered_examples_have_no_docx_errors(render_example, t):
    doc = render_example(t)
    errors = [f for f in check_docx(doc.saved_path) if f.level == "error"]
    assert errors == []


# ------------------------------------------------------------------ รหัสแต่ละตัว: รายงาน / ไม่รายงาน
LONG = "ขอความอนุเคราะห์" * 20

# code -> (ชนิด, field ที่แก้ให้ผิด, (ชนิด, field ที่แก้ให้ถูก) ที่ต้องไม่รายงาน)
CASES = {
    "PLACEHOLDER": ("external_letter", {"phone": "[ระบุ]"}, ("external_letter", {})),
    "DATE_FORMAT": ("memo", {"date": "14/11/2569"}, ("memo", {})),
    "DATE_BE": ("memo", {"date": "๑๔ พฤศจิกายน ๒๐๒๖"}, ("memo", {"date": "๑๔ พฤศจิกายน ๒๕๖๙"})),
    "DOC_NO": ("external_letter", {"doc_no": "ตย ๐๐๐๑"}, ("external_letter", {})),
    "SUBJECT_LONG": ("memo", {"subject": LONG}, ("memo", {})),
    "BODY_CAUSE": ("external_letter", {"body.cause": []}, ("external_letter", {})),
    "BODY_PURPOSE": ("memo", {"body.purpose": []}, ("memo", {})),
    "SALUTATION_PAIR": ("external_letter", {"closing": "ขอแสดงความนับถืออย่างยิ่ง"}, ("external_letter", {})),
    "SALUTATION_RANK": ("external_letter", {"to": "นายกรัฐมนตรี"}, ("external_letter", {})),
    "OWNER_UNIT": ("external_letter", {"owner_unit": ""}, ("external_letter", {})),
    "ORDER_NO": ("order", {"doc_no": "๑๒๓"}, ("order", {})),
    "ORDER_EFFECTIVE": ("order", {"effective": ""}, ("order", {})),
    "REG_ARTICLES": ("regulation", {"articles": [{"text": "ก"}, {"text": "ข"}]}, ("regulation", {})),
    "REG_NAME": ("regulation", {"articles.0.text": "ระเบียบฉบับนี้"}, ("regulation", {})),
    "REG_EFFECTIVE": ("regulation", {"articles.1.text": "ข้ออื่น"}, ("regulation", {})),
    "REG_KEEPER": ("regulation", {"articles.4.text": "ข้ออื่น"}, ("regulation", {})),
    "BYLAW_AUTHORITY": ("bylaw", {"preamble": "คณะกรรมการจึงออกข้อบังคับไว้"}, ("bylaw", {})),
    "ANN_BODY": ("announcement", {"paragraphs": []}, ("announcement", {})),
    "CERT_OPENING": ("certificate", {"body": ["นายตัวอย่าง เป็นข้าราชการ"]}, ("certificate", {})),
    "MIN_AGENDA": ("minutes", {"agenda": []}, ("minutes", {})),
    "MIN_RESOLUTION": ("minutes", {"agenda.0.resolution": "[ระบุมติที่ประชุม]"}, ("minutes", {})),
    "TEXT_PUA": ("memo", {"subject": "ขออนุมัติ"}, ("memo", {})),
    "TEXT_SARA_AM": ("memo", {"subject": "ขอทํา"}, ("memo", {"subject": "ขอทำ"})),
    "TEXT_MIXED_DIGITS": ("memo", {"subject": "ปี ๒๕6๙"}, ("memo", {"subject": "ปี ๒๕๖๙"})),
    "TEXT_BRACKET": ("memo", {"subject": "ขอ (อนุมัติ"}, ("memo", {"subject": "ขอ (อนุมัติ)"})),
    "OPENING_REF_NAN": ("memo", {"body.cause": ["ตามหนังสือที่อ้างถึง กองตัวอย่างแจ้งว่า"]},
                        ("memo", {"body.cause": ["ตามหนังสือที่อ้างถึง กองตัวอย่างแจ้งว่า... นั้น"]})),
    "CONCLUSION_RANK": ("external_letter", {"salutation": "กราบเรียน", "closing": "ขอแสดงความนับถืออย่างยิ่ง",
                                            "to": "นายกรัฐมนตรี", "body.conclusion": "จึงเรียนมาเพื่อโปรดทราบ"},
                        ("external_letter", {"salutation": "กราบเรียน", "closing": "ขอแสดงความนับถืออย่างยิ่ง",
                                             "to": "นายกรัฐมนตรี",
                                             "body.conclusion": "จึงกราบเรียนมาเพื่อโปรดทราบ"})),
    "MEMO_PROPOSAL": ("memo", {"body": None, "sections": [{"heading": "เรื่องเดิม", "paragraphs": ["ก"]}]},
                      ("memo", {"body": None, "sections": [{"heading": "เรื่องเดิม", "paragraphs": ["ก"]},
                                                          {"heading": "ข้อเสนอ", "paragraphs": ["ข"]}]})),
}


@pytest.mark.parametrize("code", sorted(CASES))
def test_code_reported_and_not_reported(code):
    bad_type, bad, (ok_type, ok) = CASES[code]
    assert code in codes(mk(bad_type, **bad)), f"{code} ควรถูกรายงาน"
    good = fill(mk(ok_type, **ok))
    assert code not in codes(good), f"{code} ไม่ควรรายงานกับร่างที่ถูกต้อง"
    assert code in ALL_CODES


def test_examples_do_not_report_most_codes():
    """ตัวอย่างทุกชนิด (ผ่าน fill) ไม่มี error และไม่มี PLACEHOLDER"""
    for t in ALL_TYPES:
        got = codes(fill(parse_doc(load_example(t))))
        assert "PLACEHOLDER" not in got, t


def test_error_levels():
    assert codes(mk("external_letter", closing="ขอแสดงความนับถืออย่างยิ่ง"))["SALUTATION_PAIR"].level == "error"
    assert codes(mk("memo", date="14/11/2569"))["DATE_FORMAT"].level == "error"
    assert codes(mk("external_letter", **{"body.purpose": []}))["BODY_PURPOSE"].level == "error"
    assert codes(mk("memo", date="๑๔ พฤศจิกายน ๒๐๒๖"))["DATE_BE"].level == "error"


def test_unknown_salutation_is_warning_not_error():
    d = mk("external_letter", to="พระครูตัวอย่าง", salutation="นมัสการ", closing="ขอนมัสการด้วยความเคารพ")
    got = codes(fill(d))
    assert got["SALUTATION_PAIR"].level == "warning"
    assert "SALUTATION_RANK" not in got
    assert not [f for f in check_document(fill(d)) if f.level == "error"]


def test_info_levels_for_memo_proposal_and_stamp_conclusion():
    memo = mk("memo", body=None, sections=[{"heading": "เรื่องเดิม", "paragraphs": ["ก"]}])
    assert codes(memo)["MEMO_PROPOSAL"].level == "info"
    spec = replace(DEFAULT_SPEC, stamp_notify_only=True)
    st = mk("stamp_letter", paragraphs=["กรมตัวอย่างขอส่งสำเนา"])
    assert codes(st, spec)["STAMP_CONCLUSION"].level == "info"
    ok = mk("stamp_letter", paragraphs=["กรมตัวอย่างขอส่งสำเนา", "จึงเรียนมาเพื่อทราบ"])
    assert "STAMP_CONCLUSION" not in codes(ok, spec)


def test_stamp_notify_only_depends_on_spec():
    st = mk("stamp_letter", paragraphs=["ขอให้ส่งข้อมูลภายในวันที่ ๒๐ ธันวาคม ๒๕๖๙"])
    assert "STAMP_NOTIFY_ONLY" not in codes(st)
    spec = replace(DEFAULT_SPEC, stamp_notify_only=True)
    f = codes(st, spec)["STAMP_NOTIFY_ONLY"]
    assert f.level == "warning"
    ok = mk("stamp_letter", paragraphs=["กรมตัวอย่างขอส่งสำเนา", "จึงเรียนมาเพื่อทราบ"])
    assert "STAMP_NOTIFY_ONLY" not in codes(ok, spec)


def test_covid_code_is_not_flagged_as_mixed_digits():
    d = mk("memo", subject="มาตรการป้องกัน COVID-19 และ ISO 27001")
    assert "TEXT_MIXED_DIGITS" not in codes(d)


def test_omit_signature_skips_signer_fields():
    d = mk("memo", **{"signer.name": "[ระบุชื่อผู้ลงนาม]"})
    assert any(f.where == "signer.name" for f in check_document(d))
    spec = replace(DEFAULT_SPEC, omit_signature=("memo",))
    assert not any(f.where.startswith("signer") for f in check_document(d, spec))


# ------------------------------------------------------------------ กติกาทั่วไปของรหัส
def _all_findings(render_example):
    out = []
    for code, (bad_type, bad, _) in CASES.items():
        out += check_document(mk(bad_type, **bad))
    out += check_document(mk("stamp_letter", paragraphs=["ขอให้ส่ง"]),
                          replace(DEFAULT_SPEC, stamp_notify_only=True))
    out += check_document(mk("external_letter", references=["ก", "ข"], subject=LONG,
                             **{"signer.name": "(นายสมมติ)", "body.cause": ["ด้วย ทำการ ครับ นั้น"]}))
    out += check_document(mk("memo", date="[ระบุ]", department="กอง"))
    out += check_docx(render_example("memo").saved_path)
    return out


def test_every_error_has_ref_and_codes_are_declared(render_example):
    fs = _all_findings(render_example)
    assert fs
    for f in fs:
        assert f.level in ("error", "warning", "info")
        assert f.code in ALL_CODES, f.code
        assert not f.code.startswith("PWA_")
        if f.level == "error":
            assert f.ref, f.code


def test_no_agency_specific_codes():
    assert not any(c.startswith("PWA_") for c in ALL_CODES)


def test_all_codes_matches_source():
    src = Path(checker.__file__).read_text(encoding="utf-8")
    found = set(re.findall(r'Finding\(\s*"(?:error|warning|info)",\s*"([A-Z][A-Z0-9_]*)"', src))
    assert found == set(ALL_CODES)
    assert len(ALL_CODES) == len(set(ALL_CODES))


def test_summarize_counts():
    fs = [Finding("error", "A", "x"), Finding("warning", "B", "y"), Finding("info", "C", "z")]
    s = summarize(fs)
    assert (s["errors"], s["warnings"], s["infos"]) == (1, 1, 1)
    assert s["findings"][0] == {"level": "error", "code": "A", "message": "x", "where": "", "ref": ""}


# ------------------------------------------------------------------ check_docx
def _saved(doc, tmp_path, name="edited.docx"):
    p = tmp_path / name
    doc.save(str(p))
    return p


def _docx_codes(path, spec=DEFAULT_SPEC):
    return {f.code: f for f in check_docx(path, spec)}


def test_docx_margin(render_example, tmp_path):
    doc = render_example("external_letter")
    assert "MARGIN" not in _docx_codes(doc.saved_path)
    doc.sections[0].left_margin = Cm(2)
    f = _docx_codes(_saved(doc, tmp_path))["MARGIN"]
    assert f.level == "error" and f.ref


def test_docx_font(render_example, tmp_path):
    doc = render_example("external_letter")
    assert "FONT" not in _docx_codes(doc.saved_path)
    for p in doc.paragraphs:
        for r in p.runs:
            rf = r._element.get_or_add_rPr().get_or_add_rFonts()
            for a in ("ascii", "hAnsi", "cs"):
                rf.set(qn(f"w:{a}"), "Arial")
    f = _docx_codes(_saved(doc, tmp_path))["FONT"]
    assert f.level == "error" and "Arial" in f.message


def test_docx_font_from_style_when_run_has_none(render_example, tmp_path):
    doc = render_example("external_letter")
    for p in doc.paragraphs:
        for r in p.runs:
            rPr = r._element.rPr
            rf = rPr.find(qn("w:rFonts")) if rPr is not None else None
            if rf is not None:
                rPr.remove(rf)
    normal = doc.styles["Normal"].element.get_or_add_rPr().get_or_add_rFonts()
    for a in ("ascii", "hAnsi", "cs"):
        normal.set(qn(f"w:{a}"), "Arial")
    assert "FONT" in _docx_codes(_saved(doc, tmp_path))


def test_docx_garuda_missing(render_example, tmp_path):
    doc = render_example("external_letter")
    assert "GARUDA" not in _docx_codes(doc.saved_path)
    for d in list(doc.element.body.iter(qn("w:drawing"))):
        d.getparent().remove(d)
    f = _docx_codes(_saved(doc, tmp_path))["GARUDA"]
    assert f.level == "warning"
    assert "ตราครุฑ" in f.message


def test_docx_garuda_uses_emblem_name(render_example, tmp_path):
    spec = replace(DEFAULT_SPEC, emblem_name="ตราหน่วยงาน")
    doc = render_example("external_letter", spec=spec)
    for d in list(doc.element.body.iter(qn("w:drawing"))):
        d.getparent().remove(d)
    assert "ตราหน่วยงาน" in _docx_codes(_saved(doc, tmp_path), spec)["GARUDA"].message


def test_docx_minutes_do_not_need_emblem(render_example, tmp_path):
    doc = render_example("minutes")
    for d in list(doc.element.body.iter(qn("w:drawing"))):
        d.getparent().remove(d)
    assert "GARUDA" not in _docx_codes(_saved(doc, tmp_path))


def _title_run(doc):
    for p in doc.paragraphs[:6]:
        for r in p.runs:
            if "บันทึกข้อความ" in r.text:
                return r
    raise AssertionError


def test_docx_memo_title_size(render_example, tmp_path):
    doc = render_example("memo")
    assert "MEMO_TITLE" not in _docx_codes(doc.saved_path)
    r = _title_run(doc)
    r.font.size = Pt(16)
    r._element.rPr.find(qn("w:szCs")).set(qn("w:val"), "32")
    f = _docx_codes(_saved(doc, tmp_path))["MEMO_TITLE"]
    assert f.level == "error" and f.ref


def test_docx_memo_label_checked_only_on_memos(render_example, tmp_path):
    doc = render_example("memo")
    assert "MEMO_LABEL" not in _docx_codes(doc.saved_path)
    assert "MEMO_LABEL" not in _docx_codes(render_example("external_letter").saved_path)
    for p in doc.paragraphs[:6]:
        for r in p.runs:
            if r.text.strip() == "เรื่อง":
                r.font.size = Pt(16)
                r._element.rPr.find(qn("w:szCs")).set(qn("w:val"), "32")
    assert "MEMO_LABEL" in _docx_codes(_saved(doc, tmp_path))


def test_docx_small_text_and_paper(render_example, tmp_path):
    doc = render_example("external_letter")
    r = next(r for r in doc.paragraphs[3].runs if r.text.strip())
    r.font.size = Pt(12)
    r._element.rPr.find(qn("w:szCs")).set(qn("w:val"), "24")
    doc.sections[0].page_width = Cm(21.59)
    got = _docx_codes(_saved(doc, tmp_path))
    assert got["FONT_SIZE"].level == "warning"
    assert got["PAPER"].level == "warning"


def test_docx_arabic_digits_and_text_quality(render_example, tmp_path):
    doc = render_example("external_letter")
    doc.add_paragraph("มาตรการ COVID-19 ISO 27001")
    assert "ARABIC_DIGITS" not in _docx_codes(_saved(doc, tmp_path, "a.docx"))
    doc.add_paragraph("ปีงบประมาณ 2570 ข้อความ")
    got = _docx_codes(_saved(doc, tmp_path, "b.docx"))
    assert got["ARABIC_DIGITS"].level == "info"
    assert got["TEXT_PUA"].level == "warning" and not got["TEXT_PUA"].ref


def test_docx_accepts_str_path(render_example):
    doc = render_example("announcement")
    assert isinstance(check_docx(str(doc.saved_path)), list)


# ------------------------------------------------------------------ ช่อง [ระบุ…] / ข้อสังเกตเพิ่ม
@pytest.mark.parametrize("t", ["memo", "external_letter", "stamp_letter"])
def test_placeholder_doc_no_is_not_a_format_error(t):
    got = codes(fill(mk(t, doc_no="[ระบุเลขที่]")))
    assert "DOC_NO" not in got and "PLACEHOLDER" in got
    assert "DOC_NO" in codes(mk(t, doc_no="ตย ๐๐๐๑"))


def test_placeholder_order_no_is_not_a_format_error():
    got = codes(fill(mk("order", doc_no="[ระบุเลขที่]")))
    assert "ORDER_NO" not in got and "PLACEHOLDER" in got
    assert "ORDER_NO" in codes(mk("order", doc_no="๑๒๓"))


def test_text_pua_is_warning_without_ref():
    f = codes(mk("memo", **CASES["TEXT_PUA"][1]))["TEXT_PUA"]
    assert f.level == "warning" and not f.ref


def test_order_default_effective_is_placeholder():
    data = copy.deepcopy(load_example("order"))
    del data["effective"]
    d = fill(parse_doc(data))
    assert "[ระบุวันใช้บังคับ]" in d.effective
    assert any(f.code == "PLACEHOLDER" and f.where == "effective" for f in check_document(d))
    assert "ORDER_EFFECTIVE" not in codes(d)


def test_stamp_forwarded_ignored_when_paragraphs_given():
    both = fill(mk("stamp_letter", paragraphs=["กรมตัวอย่างขอส่งสำเนาหนังสือดังกล่าวมาเพื่อทราบ"]))
    f = codes(both)["STAMP_FORWARDED_IGNORED"]
    assert f.level == "warning" and not f.ref
    assert "STAMP_FORWARDED_IGNORED" not in codes(fill(mk("stamp_letter")))
    only_text = fill(mk("stamp_letter", forwarded=[], paragraphs=["กรมตัวอย่างขอแจ้งเพื่อทราบ"]))
    assert "STAMP_FORWARDED_IGNORED" not in codes(only_text)


def test_docx_reports_leftover_placeholders(render_example):
    doc = render_example("memo", date="[ระบุวันที่]", doc_no="[ระบุ]/[ระบุ]",
                         **{"signer.position": ["[ระบุตำแหน่ง]"]})
    ph = [f for f in check_docx(doc.saved_path) if f.code == "PLACEHOLDER"]
    assert sorted(f.message for f in ph) == sorted(
        f"ยังไม่ได้กรอก {x}" for x in ("[ระบุวันที่]", "[ระบุ]", "[ระบุตำแหน่ง]"))
    assert all(f.level == "warning" for f in ph)
    assert not [f for f in check_docx(render_example("memo").saved_path) if f.code == "PLACEHOLDER"]
