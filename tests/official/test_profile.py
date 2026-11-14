import json
from dataclasses import replace

import pytest

from thaidocs.official.profile import (
    PLACEHOLDER_EMBLEM, AgencyInfo, ProfileError, load_profile, resolve_emblem,
)
from thaidocs.official.style import DEFAULT_SPEC, DOCUMENTED_LAYOUT_KEYS, LayoutSpec

PNG = PLACEHOLDER_EMBLEM.read_bytes()
# JPEG เล็กที่สุดที่อ่านขนาดได้ (SOI + APP0/JFIF + SOF0 1x1 + EOI)
JPG = (b"\xff\xd8" + b"\xff\xe0\x00\x10JFIF\x00\x01\x01\x00\x00\x01\x00\x01\x00\x00"
       + b"\xff\xc0\x00\x0b\x08\x00\x01\x00\x01\x01\x01\x11\x00" + b"\xff\xd9")


@pytest.fixture
def home(tmp_path, monkeypatch):
    monkeypatch.setenv("HOME", str(tmp_path))
    return tmp_path


def write(path, data):
    path.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    return path


def test_no_file_gives_defaults(home):
    assert load_profile(None) == (None, DEFAULT_SPEC)


def test_default_font_and_no_profile_field():
    assert DEFAULT_SPEC.font == "TH Sarabun New"
    assert "profile" not in LayoutSpec.__dataclass_fields__
    assert "emblem_name" in LayoutSpec.__dataclass_fields__
    assert all(k in LayoutSpec.__dataclass_fields__ for k in DOCUMENTED_LAYOUT_KEYS)


def test_home_default_file(home):
    (home / ".thaidocs").mkdir()
    write(home / ".thaidocs" / "agency.json", {"agency": {"org": "กรมตัวอย่าง"}})
    agency, spec = load_profile(None)
    assert agency == AgencyInfo(org="กรมตัวอย่าง") and spec == DEFAULT_SPEC


def test_explicit_file(home):
    p = write(home / "a.json", {
        "agency": {"org": "กรมตัวอย่าง", "address": ["ก", "ข"], "signer_position": ["ผู้อำนวยการ"]},
        "layout": {"unit_word": "หน่วยงาน", "margin_left_cm": 2.5},
    })
    agency, spec = load_profile(p)
    assert agency.org == "กรมตัวอย่าง" and agency.address == ["ก", "ข"]
    assert agency.signer_position == ["ผู้อำนวยการ"]
    assert spec == replace(DEFAULT_SPEC, unit_word="หน่วยงาน", margin_left_cm=2.5)


def test_utf8_bom(home):
    p = home / "a.json"
    p.write_bytes(b"\xef\xbb\xbf" + json.dumps({"agency": {"org": "ก"}}, ensure_ascii=False).encode())
    assert load_profile(p)[0].org == "ก"


def test_unknown_top_level_key(home):
    p = write(home / "a.json", {"agency": {}, "ตรา": 1})
    with pytest.raises(ProfileError, match="ตรา"):
        load_profile(p)


def test_unknown_layout_key(home):
    p = write(home / "a.json", {"layout": {"no_such_key": 1}})
    with pytest.raises(ProfileError, match="no_such_key"):
        load_profile(p)


def test_unknown_agency_key(home):
    p = write(home / "a.json", {"agency": {"nope": "x"}})
    with pytest.raises(ProfileError, match="nope"):
        load_profile(p)


def test_wrong_type(home):
    p = write(home / "a.json", {"layout": {"margin_left_cm": "สาม"}})
    with pytest.raises(ProfileError, match="margin_left_cm"):
        load_profile(p)


def test_bool_not_numeric(home):
    p = write(home / "a.json", {"layout": {"body_pt": True}})
    with pytest.raises(ProfileError, match="body_pt"):
        load_profile(p)


def test_int_ok_for_float(home):
    p = write(home / "a.json", {"layout": {"body_pt": 14}})
    assert load_profile(p)[1].body_pt == 14


def test_optional_null(home):
    p = write(home / "a.json", {"layout": {"emblem_png": None, "signer_paren_cm": 9.5}})
    assert load_profile(p)[1].signer_paren_cm == 9.5


def test_bad_json(home):
    p = home / "a.json"
    p.write_text("{ไม่ใช่ json", encoding="utf-8")
    with pytest.raises(ProfileError):
        load_profile(p)


def test_missing_explicit_file(home):
    with pytest.raises(ProfileError):
        load_profile(home / "none.json")


def test_omit_signature_tuple(home):
    p = write(home / "a.json", {"layout": {"omit_signature": ["memo", "order"]}})
    assert load_profile(p)[1].omit_signature == ("memo", "order")


def test_omit_signature_bad_item(home):
    p = write(home / "a.json", {"layout": {"omit_signature": ["memo", 3]}})
    with pytest.raises(ProfileError, match="omit_signature"):
        load_profile(p)


def test_emblem_order(home):
    cli = home / "cli.png"
    cli.write_bytes(PNG)
    spec_png = home / "spec.jpg"
    spec_png.write_bytes(JPG)
    spec = replace(DEFAULT_SPEC, emblem_png=str(spec_png))
    assert resolve_emblem(None, DEFAULT_SPEC) is None
    (home / ".thaidocs").mkdir()
    default = home / ".thaidocs" / "emblem.png"
    default.write_bytes(PNG)
    assert resolve_emblem(None, DEFAULT_SPEC) == default
    assert resolve_emblem(None, spec) == spec_png
    assert resolve_emblem(cli, spec) == cli


def test_emblem_invalid(home):
    folder = home / "d"
    folder.mkdir()
    bad = home / "bad.png"
    bad.write_bytes(b"GIF89a....")
    truncated = home / "truncated.png"
    truncated.write_bytes(b"\x89PNG\r\n\x1a\n" + b"garbage")
    for p in (home / "none.png", folder, bad, truncated):
        with pytest.raises(ProfileError):
            resolve_emblem(p, DEFAULT_SPEC)
    with pytest.raises(ProfileError):
        resolve_emblem(None, replace(DEFAULT_SPEC, emblem_png=str(bad)))


def test_placeholder_emblem():
    assert PLACEHOLDER_EMBLEM.is_file()
    assert PLACEHOLDER_EMBLEM.read_bytes().startswith(b"\x89PNG\r\n\x1a\n")
    assert PLACEHOLDER_EMBLEM.stat().st_size < 20000


def test_truncated_emblem_message_names_file(home):
    p = home / "cut.png"
    p.write_bytes(PNG[:40])
    with pytest.raises(ProfileError, match="cut.png") as e:
        resolve_emblem(p, DEFAULT_SPEC)
    assert "ตรา" in str(e.value)


@pytest.mark.parametrize("fmt", ["หน้า", "{} / {}", "- {0} -"])
def test_page_number_fmt_needs_one_slot(home, fmt):
    p = write(home / "a.json", {"layout": {"page_number_fmt": fmt}})
    with pytest.raises(ProfileError, match="page_number_fmt"):
        load_profile(p)


def test_page_number_fmt_ok(home):
    p = write(home / "a.json", {"layout": {"page_number_fmt": "หน้า {}"}})
    assert load_profile(p)[1].page_number_fmt == "หน้า {}"


@pytest.mark.parametrize("key,value", [("margin_left_cm", 0), ("margin_top_cm", -1), ("memo_margin_top_cm", 0),
                                       ("para_indent_cm", -0.5), ("body_pt", -16), ("para_before_pt", -6),
                                       ("page_number_pt", -1)])
def test_negative_or_zero_sizes_rejected(home, key, value):
    p = write(home / "a.json", {"layout": {key: value}})
    with pytest.raises(ProfileError, match=key):
        load_profile(p)


def test_zero_spacing_allowed(home):
    p = write(home / "a.json", {"layout": {"para_before_pt": 0, "separator_cm": 0}})
    assert load_profile(p)[1].para_before_pt == 0


def test_omit_signature_unknown_type(home):
    p = write(home / "a.json", {"layout": {"omit_signature": ["memo", "บันทึก"]}})
    with pytest.raises(ProfileError, match="omit_signature") as e:
        load_profile(p)
    assert "บันทึก" in str(e.value)


def test_doc_type_codes_match_schema():
    from thaidocs.official.profile import DOC_TYPE_CODES
    from thaidocs.official.schema import DocType
    assert set(DOC_TYPE_CODES) == {t.value for t in DocType}


def test_unread_layout_fields_removed():
    assert "date_be_rules" not in LayoutSpec.__dataclass_fields__
    assert "doc_no_prefix" not in LayoutSpec.__dataclass_fields__
    assert "doc_no_prefix" in AgencyInfo.__dataclass_fields__
