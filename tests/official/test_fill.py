import copy

import pytest

from thaidocs.official.fill import fill, forward_paragraphs
from thaidocs.official.profile import AgencyInfo
from thaidocs.official.schema import (DocType, ExternalLetterDoc, ForwardedLetter, MinutesDoc,
                                      StampLetterDoc, load_example, parse_doc)


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


AGENCY = AgencyInfo(
    org="กรมตัวอย่าง", address=["กรมตัวอย่าง", "ถนนตัวอย่าง กรุงเทพมหานคร ๑๐๐๐๐"],
    department="กองตัวอย่าง", phone="๐ ๒๐๐๐ ๐๐๐๐", fax="๐ ๒๐๐๐ ๐๐๐๑", email="saraban@example.org",
    doc_no_prefix="ตย", signer_name="นายสมมติ ตัวอย่าง", signer_position=["อธิบดีกรมตัวอย่าง"])


# ------------------------------------------------------------------ คำขึ้นต้น / คำลงท้าย
def test_salutation_defaults_follow_recipient_rank():
    d = mk("external_letter", to="นายกรัฐมนตรี", salutation="เรียน", closing="ขอแสดงความนับถือ")
    out = fill(d)
    assert (out.salutation, out.closing) == ("กราบเรียน", "ขอแสดงความนับถืออย่างยิ่ง")


def test_salutation_kept_when_draft_changed_it():
    d = mk("external_letter", to="พระครูตัวอย่าง", salutation="นมัสการ", closing="ขอนมัสการด้วยความเคารพ")
    out = fill(d)
    assert (out.salutation, out.closing) == ("นมัสการ", "ขอนมัสการด้วยความเคารพ")


def test_salutation_kept_when_only_one_changed():
    d = mk("external_letter", to="นายกรัฐมนตรี", salutation="เรียน", closing="ขอแสดงความนับถืออย่างสูง")
    out = fill(d)
    assert (out.salutation, out.closing) == ("เรียน", "ขอแสดงความนับถืออย่างสูง")


# ------------------------------------------------------------------ ทำความสะอาดคำ
def test_subject_to_signer_cleanup():
    d = mk("external_letter", subject="เรื่อง ขออนุมัติ", to="เรียน อธิบดี",
           **{"signer.name": "(นายสมมติ ตัวอย่าง)"})
    out = fill(d)
    assert out.subject == "ขออนุมัติ"
    assert out.to == "อธิบดี"
    assert out.signer.name == "นายสมมติ ตัวอย่าง"


def test_subject_word_inside_is_kept():
    out = fill(mk("memo", subject="เรื่องเดิมที่ค้างพิจารณา"))
    assert out.subject == "เรื่องเดิมที่ค้างพิจารณา"


def test_subject_colon_and_to_krab():
    out = fill(mk("memo", subject="เรื่อง: ขออนุมัติ", to="กราบเรียน ผู้อำนวยการ"))
    assert (out.subject, out.to) == ("ขออนุมัติ", "ผู้อำนวยการ")


def test_recorder_name_parens_stripped():
    out = fill(mk("minutes", **{"recorder.name": "(นางตัวอย่าง สมมติ)"}))
    assert out.recorder.name == "นางตัวอย่าง สมมติ"


# ------------------------------------------------------------------ ข้อมูลหน่วยงาน
def test_agency_fills_placeholders_and_blanks():
    d = mk("external_letter", agency=["[ระบุส่วนราชการเจ้าของหนังสือ]"], fax=None, email=None,
           **{"signer.name": "[ระบุชื่อผู้ลงนาม]", "signer.position": ["[ระบุตำแหน่ง]"]})
    d.owner_unit, d.phone = "[ระบุส่วนราชการเจ้าของเรื่อง]", "[ระบุ]"
    out = fill(d, AGENCY)
    assert out.agency == AGENCY.address
    assert out.owner_unit == "กองตัวอย่าง"
    assert out.phone == "๐ ๒๐๐๐ ๐๐๐๐"
    assert out.fax == AGENCY.fax and out.email == "saraban@example.org"
    assert out.signer.name == "นายสมมติ ตัวอย่าง"
    assert out.signer.position == ["อธิบดีกรมตัวอย่าง"]


def test_agency_does_not_overwrite_given_values():
    d = mk("external_letter", owner_unit="กองอื่น", phone="๐ ๑๑๑๑ ๑๑๑๑", email="a@example.org",
           **{"signer.name": "นางสาวผู้อื่น ตัวอย่าง"})
    out = fill(d, AGENCY)
    assert out.owner_unit == "กองอื่น" and out.phone == "๐ ๑๑๑๑ ๑๑๑๑"
    assert out.email == "a@example.org" and out.fax == "๐ ๒๐๐๐ ๐๐๐๑"
    assert out.signer.name == "นางสาวผู้อื่น ตัวอย่าง"
    assert out.agency == d.agency


def test_agency_none_changes_nothing_but_keeps_placeholders():
    d = mk("external_letter", **{"signer.name": "[ระบุชื่อผู้ลงนาม]"})
    out = fill(d, None)
    assert out.signer.name == "[ระบุชื่อผู้ลงนาม]"
    assert out is not d


def test_agency_fills_stamp_certificate_memo_and_issuer():
    s = mk("stamp_letter", issuer="[ระบุ]", paragraphs=["ข้อความ"])
    s.owner_unit, s.phone, s.fax, s.email = "[ระบุส่วนราชการเจ้าของเรื่อง]", "[ระบุ]", None, None
    out = fill(s, AGENCY)
    assert (out.issuer, out.owner_unit, out.phone) == ("กรมตัวอย่าง", "กองตัวอย่าง", "๐ ๒๐๐๐ ๐๐๐๐")
    assert (out.fax, out.email) == (AGENCY.fax, AGENCY.email)
    c = fill(mk("certificate", agency=[]), AGENCY)
    assert c.agency == AGENCY.address
    m = fill(mk("memo", department="[ระบุ]"), AGENCY)
    assert m.department == "กองตัวอย่าง โทร. ๐ ๒๐๐๐ ๐๐๐๐"
    o = fill(mk("order", issuer="[ระบุผู้ออกคำสั่ง]"), AGENCY)
    assert o.issuer == "กรมตัวอย่าง"


def test_doc_no_prefix_is_never_inserted():
    d = mk("external_letter", doc_no="[ระบุ]/[ระบุ]")
    assert fill(d, AGENCY).doc_no == "[ระบุ]/[ระบุ]"


def test_recorder_not_filled_from_agency():
    out = fill(mk("minutes", **{"recorder.name": "[ระบุผู้จดรายงานการประชุม]"}), AGENCY)
    assert out.recorder.name == "[ระบุผู้จดรายงานการประชุม]"


# ------------------------------------------------------------------ หนังสือประทับตรา
def _fw(n):
    return [ForwardedLetter(agency="กระทรวงตัวอย่าง", speed="ด่วน" if i == 0 else None,
                            doc_no=f"ตย ๐๒๐๑/ว ๑๒{'๐๑๒'[i]}", date="๑๐ พฤศจิกายน ๒๕๖๙",
                            subject=f"เรื่องที่ {'๑๒๓'[i]}")
            for i in range(n)]


def test_forward_one_letter():
    paras = forward_paragraphs("กรมตัวอย่าง", _fw(1))
    assert len(paras) == 2
    assert paras[0] == ("กรมตัวอย่าง ขอส่งสำเนาหนังสือกระทรวงตัวอย่าง ด่วน ที่ ตย ๐๒๐๑/ว ๑๒๐ "
                        "ลงวันที่ ๑๐ พฤศจิกายน ๒๕๖๙ เรื่อง เรื่องที่ ๑")
    assert paras[1] == "จึงเรียนมาเพื่อทราบ"


def test_forward_two_letters_joined_in_one_paragraph():
    paras = forward_paragraphs("กรมตัวอย่าง", _fw(2))
    assert len(paras) == 2
    assert " และสำเนาหนังสือกระทรวงตัวอย่าง ที่ ตย ๐๒๐๑/ว ๑๒๑" in paras[0]
    assert paras[1] == "จึงเรียนมาเพื่อทราบ"


def test_forward_three_letters_listed():
    paras = forward_paragraphs("กรมตัวอย่าง", _fw(3))
    assert paras[0] == "กรมตัวอย่าง ขอส่งสำเนาหนังสือ จำนวน 3 ฉบับ ดังนี้"
    assert [p[:2] for p in paras[1:4]] == ["1.", "2.", "3."]
    assert paras[-1] == "จึงเรียนมาเพื่อทราบ"
    assert len(paras) == 5


def test_fill_writes_forward_paragraphs_when_none_given():
    d = mk("stamp_letter", paragraphs=[])
    out = fill(d)
    assert isinstance(out, StampLetterDoc)
    assert out.paragraphs == forward_paragraphs(d.issuer, d.forwarded)
    assert d.paragraphs == []


def test_fill_keeps_own_paragraphs():
    d = mk("stamp_letter", paragraphs=["ข้อความของร่างเอง"])
    assert fill(d).paragraphs == ["ข้อความของร่างเอง"]


# ------------------------------------------------------------------ ร้อยแก้วและเวลา
def test_newline_paragraphs_split_and_lists_merged():
    d = mk("external_letter", **{
        "body.purpose": ["กรมตัวอย่างขอเชิญเข้าร่วมประชุม รายละเอียดดังนี้\nวันที่: ๒๐ ธันวาคม ๒๕๖๙\n"
                         "เวลา: ๑๓.๓๐ น. ถึง ๑๖.๓๐ น.\nสถานที่: ห้องประชุมตัวอย่าง", "ย่อหน้าที่สอง\nย่อหน้าที่สาม"]})
    before = d.model_dump()
    out = fill(d)
    assert out.body.purpose == [
        "กรมตัวอย่างขอเชิญเข้าร่วมประชุม ในวันที่ ๒๐ ธันวาคม ๒๕๖๙ เวลา ๑๓.๓๐ – ๑๖.๓๐ น. ณ ห้องประชุมตัวอย่าง",
        "ย่อหน้าที่สอง", "ย่อหน้าที่สาม"]
    assert d.model_dump() == before


def test_opening_word_joined_to_next_word():
    out = fill(mk("external_letter", **{"body.cause": ["ด้วย กรมตัวอย่างมีความประสงค์"]}))
    assert out.body.cause == ["ด้วยกรมตัวอย่างมีความประสงค์"]


def test_prose_applies_to_every_paragraph_list():
    two = "ข้อความ\nอีกข้อความ"
    assert fill(mk("order", paragraphs=[two])).paragraphs == ["ข้อความ", "อีกข้อความ"]
    assert fill(mk("announcement", paragraphs=[two])).paragraphs == ["ข้อความ", "อีกข้อความ"]
    assert fill(mk("certificate", body=[two])).body == ["ข้อความ", "อีกข้อความ"]
    memo = mk("memo", body=None, sections=[{"heading": "เรื่องเดิม", "paragraphs": [two]}])
    assert fill(memo).sections[0].paragraphs == ["ข้อความ", "อีกข้อความ"]
    o = mk("order", **{"committees.2.paragraphs": [two], "committees.0.duties": [two]})
    out = fill(o)
    assert out.committees[2].paragraphs == ["ข้อความ", "อีกข้อความ"]
    assert out.committees[0].duties == ["ข้อความ", "อีกข้อความ"]
    mi = fill(mk("minutes", **{"agenda.0.discussion": [two]}))
    assert mi.agenda[0].discussion == ["ข้อความ", "อีกข้อความ"]


@pytest.mark.parametrize("given,expected", [
    ("๐๙.๓๐ น.", "๐๙.๓๐"), ("๐๙.๓๐ น", "๐๙.๓๐"), ("๐๙.๓๐น.", "๐๙.๓๐"),
    ("๐๙.๓๐ นาฬิกา", "๐๙.๓๐"), ("๐๙.๓๐", "๐๙.๓๐"), ("[ระบุ]", "[ระบุ]"),
])
def test_minutes_time_has_no_trailing_unit(given, expected):
    out = fill(mk("minutes", start_time=given, end_time=given))
    assert isinstance(out, MinutesDoc)
    assert out.start_time == expected and out.end_time == expected


def test_minutes_time_has_no_double_unit_text():
    d = fill(mk("minutes", start_time="๐๙.๓๐ น.", end_time="๑๑.๐๐ น."))
    assert "น. น." not in d.start_time + d.end_time
    assert d.start_time.endswith("๐๙.๓๐")


# ------------------------------------------------------------------ ไม่แก้ของเดิม / ตัวอย่างทั้งหมด
@pytest.mark.parametrize("t", [t.value for t in DocType])
def test_fill_returns_copy_and_is_idempotent(t):
    d = parse_doc(load_example(t))
    before = d.model_dump()
    out = fill(d, AGENCY)
    assert d.model_dump() == before
    assert type(out) is type(d)
    assert fill(out, AGENCY).model_dump() == out.model_dump()


def test_default_salutation_helper_not_applied_to_memo():
    d = mk("memo", to="นายกรัฐมนตรี")
    assert fill(d).salutation == "เรียน"
