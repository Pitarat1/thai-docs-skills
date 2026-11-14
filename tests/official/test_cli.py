"""เทสคำสั่ง thaigov (thaidocs.official.cli.main)"""
from __future__ import annotations

import copy
import importlib
import json
import sys

import pytest
from docx import Document

from thaidocs.official.checker import check_document
from thaidocs.official.cli import _parser, main
from thaidocs.official.fill import fill
from thaidocs.official.schema import (DOC_TYPE_REF, DOC_TYPE_TH, MODEL_BY_TYPE, DocType, load_example, parse_doc)
from thaidocs.official.textutil import to_arabic_digits
from thaidocs.ooxml import is_wrapped

TYPES = [t.value for t in DocType]
THAI_DIGITS = "๐๑๒๓๔๕๖๗๘๙"


@pytest.fixture(autouse=True)
def home(tmp_path, monkeypatch):
    h = tmp_path / "home"
    h.mkdir()
    monkeypatch.setenv("HOME", str(h))
    monkeypatch.setenv("USERPROFILE", str(h))
    return h


@pytest.fixture
def run(capsys):
    def _run(*argv):
        code = main([str(a) for a in argv])
        cap = capsys.readouterr()
        return code, cap.out, cap.err
    return _run


def write_json(path, data):
    path.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    return path


def text_of(path) -> str:
    return "\n".join(p.text for p in Document(str(path)).paragraphs).replace("\xa0", " ")


# ------------------------------------------------------------------ types / schema / example
def test_types_lists_eleven(run):
    code, out, _ = run("types")
    assert code == 0
    lines = [x for x in out.splitlines() if x.strip()]
    assert len(lines) == 11
    for t, line in zip(DocType, lines):
        assert t.value in line and DOC_TYPE_TH[t] in line and DOC_TYPE_REF[t] in line


def test_schema_memo_has_every_field_with_description(run):
    code, out, _ = run("schema", "memo")
    assert code == 0
    for name, f in MODEL_BY_TYPE[DocType.MEMO].model_fields.items():
        assert name in out
        if f.description:
            assert f.description in out
    assert "จำเป็น" in out
    # โมเดลย่อยเยื้องใต้ field แม่
    assert any(line.startswith("    ") and "name" in line for line in out.splitlines())


def test_schema_unknown_type(run):
    code, out, err = run("schema", "xyz")
    assert code == 2
    assert "xyz" in err and "memo" in err and "minutes" in err
    assert "Traceback" not in err


def test_example_to_stdout_and_file(run, tmp_path):
    code, out, _ = run("example", "order")
    assert code == 0
    parse_doc(json.loads(out))
    dest = tmp_path / "out.json"
    code, out, _ = run("example", "order", dest)
    assert code == 0 and dest.is_file()
    parse_doc(json.loads(dest.read_text(encoding="utf-8")))
    code, _, err = run("example", "order", dest)       # ไม่เขียนทับ
    assert code == 2 and "มีอยู่แล้ว" in err


# ------------------------------------------------------------------ render
@pytest.mark.parametrize("t", TYPES)
def test_render_every_example(run, tmp_path, t):
    src = tmp_path / f"{t}.json"
    assert run("example", t, src)[0] == 0
    out_file = tmp_path / f"{t}.docx"
    code, out, err = run("render", src, out_file)
    assert code == 0, err
    Document(str(out_file))
    assert "สร้างไฟล์" in out and out_file.name in out
    for word in ("ข้อผิดพลาด", "คำเตือน", "ข้อสังเกต"):
        assert word in out


def test_render_reports_placeholder_emblem(run, tmp_path):
    src = tmp_path / "m.json"
    run("example", "memo", src)
    code, out, _ = run("render", src, tmp_path / "m.docx")
    assert code == 0 and "กรอบแทนตรา" in out and "--emblem" in out


def test_render_with_real_emblem_no_placeholder_note(run, tmp_path):
    src = tmp_path / "m.json"
    run("example", "memo", src)
    from thaidocs.official.profile import PLACEHOLDER_EMBLEM
    code, out, _ = run("render", src, tmp_path / "m.docx", "--emblem", PLACEHOLDER_EMBLEM)
    assert code == 0 and "กรอบแทนตรา" not in out


def test_render_existing_output_needs_force(run, tmp_path):
    src = tmp_path / "m.json"
    run("example", "memo", src)
    dest = tmp_path / "m.docx"
    dest.write_bytes(b"old")
    code, _, err = run("render", src, dest)
    assert code == 2 and "--force" in err and dest.read_bytes() == b"old"
    assert run("render", src, dest, "--force")[0] == 0
    Document(str(dest))


def test_render_bad_output_paths(run, tmp_path):
    src = tmp_path / "m.json"
    run("example", "memo", src)
    assert run("render", src, tmp_path / "m.txt")[0] == 2
    assert run("render", src, tmp_path / "nodir" / "m.docx")[0] == 2
    d = tmp_path / "dir.docx"
    d.mkdir()
    assert run("render", src, d, "--force")[0] == 2


def test_render_arabic_digits(run, tmp_path):
    data = copy.deepcopy(load_example("order"))
    text = to_arabic_digits(json.dumps(data, ensure_ascii=False))
    src = tmp_path / "o.json"
    src.write_text(text, encoding="utf-8")
    dest = tmp_path / "o.docx"
    code, _, err = run("render", src, dest, "--arabic-digits")
    assert code == 0, err
    body = text_of(dest)
    assert not any(c in body for c in THAI_DIGITS)
    assert any(c in body for c in "0123456789")
    dest2 = tmp_path / "o2.docx"
    run("render", src, dest2)
    assert any(c in text_of(dest2) for c in THAI_DIGITS)


def test_render_font(run, tmp_path):
    src = tmp_path / "m.json"
    run("example", "memo", src)
    dest = tmp_path / "m.docx"
    assert run("render", src, dest, "--font", "TH SarabunPSK")[0] == 0
    fonts = {r.font.name for p in Document(str(dest)).paragraphs for r in p.runs if r.text.strip()}
    assert fonts == {"TH SarabunPSK"}


def test_render_wrap_words(run, tmp_path):
    src = tmp_path / "m.json"
    run("example", "memo", src)
    dest = tmp_path / "m.docx"
    code, _, err = run("render", src, dest, "--wrap-words")
    assert code == 0, err
    assert is_wrapped(Document(str(dest)))
    plain = tmp_path / "p.docx"
    run("render", src, plain)
    assert not is_wrapped(Document(str(plain)))


def test_words_requires_wrap_words(run, tmp_path):
    src = tmp_path / "m.json"
    run("example", "memo", src)
    w = tmp_path / "w.txt"
    w.write_text("กรมตัวอย่าง\n", encoding="utf-8")
    code, _, err = run("render", src, tmp_path / "m.docx", "--words", w)
    assert code == 2 and "--wrap-words" in err
    assert run("render", src, tmp_path / "m.docx", "--wrap-words", "--words", w)[0] == 0


def test_render_profile_fills_agency(run, tmp_path):
    data = copy.deepcopy(load_example("external_letter"))
    data["owner_unit"] = "[ระบุส่วนราชการเจ้าของเรื่อง]"
    data["phone"] = "[ระบุ]"
    src = write_json(tmp_path / "e.json", data)
    prof = write_json(tmp_path / "p.json", {"agency": {"department": "กองสมมติ", "phone": "๐ ๙๙๙๙ ๐๐๐๐"}})
    dest = tmp_path / "e.docx"
    code, _, err = run("render", src, dest, "--profile", prof)
    assert code == 0, err
    body = text_of(dest)
    assert "กองสมมติ" in body and "๙๙๙๙" in body


def test_render_uses_home_profile(run, tmp_path, home):
    d = home / ".thaidocs"
    d.mkdir()
    write_json(d / "agency.json", {"agency": {"department": "กองในบ้าน"}})
    data = copy.deepcopy(load_example("external_letter"))
    data["owner_unit"] = "[ระบุส่วนราชการเจ้าของเรื่อง]"
    src = write_json(tmp_path / "e.json", data)
    dest = tmp_path / "e.docx"
    assert run("render", src, dest)[0] == 0
    assert "กองในบ้าน" in text_of(dest)


def test_bad_profile(run, tmp_path):
    src = tmp_path / "m.json"
    run("example", "memo", src)
    bad = tmp_path / "p.json"
    bad.write_text("{ไม่ใช่ json", encoding="utf-8")
    code, _, err = run("render", src, tmp_path / "m.docx", "--profile", bad)
    assert code == 2 and "ข้อผิดพลาด" in err and "Traceback" not in err
    code, _, err = run("render", src, tmp_path / "m.docx", "--profile", tmp_path / "none.json")
    assert code == 2 and "Traceback" not in err


def test_stamp_forward_digits_follow_numeral_setting(run, tmp_path):
    data = copy.deepcopy(load_example("stamp_letter"))
    one = data["forwarded"][0]
    data["forwarded"] = [dict(one, subject=f"เรื่องที่ {i}") for i in range(3)]
    data["paragraphs"] = []
    src = write_json(tmp_path / "s.json", data)
    a, b = tmp_path / "a.docx", tmp_path / "b.docx"
    assert run("render", src, a)[0] == 0
    assert run("render", src, b, "--arabic-digits")[0] == 0
    assert "จำนวน ๓ ฉบับ" in text_of(a)
    assert "จำนวน 3 ฉบับ" in text_of(b)


# ------------------------------------------------------------------ ข้อมูลเข้าเสีย
def _render_fail(run, tmp_path, content: bytes | str):
    src = tmp_path / "bad.json"
    if isinstance(content, str):
        src.write_text(content, encoding="utf-8")
    else:
        src.write_bytes(content)
    return run("render", src, tmp_path / "bad.docx")


@pytest.mark.parametrize("content", [
    "นี่ไม่ใช่ JSON", "[1, 2]", "{}", '{"doc_type": "memo", "department": 5}', b"\xff\xfe\x00bad",
])
def test_bad_input_exits_2_in_thai(run, tmp_path, content):
    code, out, err = _render_fail(run, tmp_path, content)
    assert code == 2
    assert "Traceback" not in err and "Traceback" not in out
    assert "ข้อผิดพลาด" in err
    assert not (tmp_path / "bad.docx").exists()


def test_bad_field_message_names_field(run, tmp_path):
    code, _, err = _render_fail(run, tmp_path, '{"doc_type": "memo", "department": 5}')
    assert code == 2 and "department" in err


def test_json_with_bom_is_read(run, tmp_path):
    src = tmp_path / "m.json"
    src.write_bytes(b"\xef\xbb\xbf" + json.dumps(load_example("memo"), ensure_ascii=False).encode("utf-8"))
    assert run("render", src, tmp_path / "m.docx")[0] == 0


def test_missing_input_file(run, tmp_path):
    code, _, err = run("render", tmp_path / "none.json", tmp_path / "x.docx")
    assert code == 2 and "ไม่พบไฟล์" in err


def test_bad_emblem(run, tmp_path):
    src = tmp_path / "m.json"
    run("example", "memo", src)
    txt = tmp_path / "e.png"
    txt.write_text("ไม่ใช่ภาพ", encoding="utf-8")
    d = tmp_path / "folder"
    d.mkdir()
    for bad in (tmp_path / "none.png", d, txt):
        code, _, err = run("render", src, tmp_path / "m.docx", "--emblem", bad)
        assert code == 2 and "ตรา" in err and "Traceback" not in err
    assert not (tmp_path / "m.docx").exists()


# ------------------------------------------------------------------ check
def test_check_example_json_ok(run, tmp_path):
    src = tmp_path / "m.json"
    run("example", "memo", src)
    code, out, _ = run("check", src)
    assert code == 0


def test_check_bad_date_exits_1(run, tmp_path):
    data = copy.deepcopy(load_example("memo"))
    data["date"] = "14/11/2026"
    src = write_json(tmp_path / "m.json", data)
    code, out, _ = run("check", src)
    assert code == 1 and "DATE_FORMAT" in out


def test_check_docx_after_render(run, tmp_path):
    src = tmp_path / "m.json"
    run("example", "memo", src)
    dest = tmp_path / "m.docx"
    run("render", src, dest)
    code, out, err = run("check", dest)
    assert code == 0, out + err


def test_check_other_extension(run, tmp_path):
    f = tmp_path / "m.txt"
    f.write_text("x", encoding="utf-8")
    code, _, err = run("check", f)
    assert code == 2 and ".json" in err and ".docx" in err


def test_check_json_with_profile_has_no_placeholder_for_filled(run, tmp_path):
    data = copy.deepcopy(load_example("external_letter"))
    data["owner_unit"] = "[ระบุส่วนราชการเจ้าของเรื่อง]"
    data["phone"] = "[ระบุ]"
    src = write_json(tmp_path / "e.json", data)
    _, out, _ = run("check", src)
    assert "PLACEHOLDER" in out
    prof = write_json(tmp_path / "p.json", {"agency": {"department": "กองสมมติ", "phone": "๐ ๙๙๙๙ ๐๐๐๐"}})
    _, out, _ = run("check", src, "--profile", prof)
    assert "PLACEHOLDER" not in out


def test_check_bad_docx(run, tmp_path):
    f = tmp_path / "x.docx"
    f.write_bytes(b"not a zip")
    code, _, err = run("check", f)
    assert code == 2 and "Traceback" not in err


# ------------------------------------------------------------------ pydantic ไม่มี / parser
def test_missing_pydantic(run, tmp_path, monkeypatch):
    for name in [m for m in sys.modules if m == "pydantic" or m.startswith("pydantic.")
                 or m == "thaidocs.official.schema"]:
        monkeypatch.delitem(sys.modules, name)
    monkeypatch.setitem(sys.modules, "pydantic", None)
    src = tmp_path / "m.json"
    src.write_text("{}", encoding="utf-8")
    for argv in (("types",), ("schema", "memo"), ("render", src, tmp_path / "m.docx"), ("check", src)):
        code, _, err = run(*argv)
        assert code == 2 and "pip install pydantic" in err and "Traceback" not in err
    importlib.invalidate_caches()


def test_parser_function():
    p = _parser()
    a = p.parse_args(["render", "a.json", "b.docx", "--wrap-words", "--words", "x", "--words", "y"])
    assert a.words == ["x", "y"] and a.wrap_words


def test_examples_have_no_error_or_warning(tmp_path):
    for t in DocType:
        doc = fill(parse_doc(load_example(t)))
        bad = [f for f in check_document(doc) if f.level in ("error", "warning") and f.code != "PLACEHOLDER"]
        assert bad == [], (t.value, [(f.code, f.where) for f in bad])
