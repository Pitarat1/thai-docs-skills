import subprocess
import sys

import pytest
from docx import Document

from thaidocs import cli
from thaidocs.ooxml import is_wrapped

TEXT = "อำเภอวังน้ำเย็นจังหวัดสระแก้วมีประชากรจำนวนมาก"


def _src(make_doc):
    doc = make_doc([TEXT, ("บทนำ", "Heading 1")])
    return doc.saved_path


def test_fix_roundtrip(make_doc, tmp_path, capsys):
    src = _src(make_doc)
    out = tmp_path / "out.docx"
    assert cli.main(["fix", str(src), str(out)]) == 0
    assert out.exists()
    assert cli.main(["check", str(out)]) == 0


def test_fix_on_wrapped_doc_rewraps(make_doc, tmp_path):
    src = _src(make_doc)
    w = tmp_path / "w.docx"
    assert cli.main(["wrap", str(src), str(w)]) == 0
    assert is_wrapped(Document(str(w)))
    out = tmp_path / "out.docx"
    assert cli.main(["fix", str(w), str(out)]) == 0
    d = Document(str(out))
    assert is_wrapped(d)
    ns = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    for r in d.element.body.iter(ns + "r"):
        text = "".join(t.text or "" for t in r.iter(ns + "t"))
        if any("฀" <= c <= "๿" for c in text):
            assert r.find(f"{ns}rPr/{ns}cs") is None
    assert cli.main(["check", str(out)]) == 0


def test_refuse_same_path(make_doc, capsys):
    src = _src(make_doc)
    assert cli.main(["fix", str(src), str(src)]) == 2
    assert capsys.readouterr().err


def test_refuse_word_lock(make_doc, tmp_path, capsys):
    src = _src(make_doc)
    (tmp_path / ("~$" + src.name)).write_bytes(b"x")
    assert cli.main(["fix", str(src), str(tmp_path / "o.docx")]) == 2
    assert "ปิดไฟล์" in capsys.readouterr().err
    assert not (tmp_path / "o.docx").exists()


def test_segment_output():
    r = subprocess.run([sys.executable, "-m", "thaidocs", "segment", "อำเภอวังน้ำเย็น"],
                       capture_output=True, check=True, env=None)
    assert r.stdout.decode("utf-8").strip() == "อำเภอ|วังน้ำเย็น"


def test_missing_dependency_message(make_doc, tmp_path, monkeypatch, capsys):
    src = _src(make_doc)
    import builtins
    real = builtins.__import__

    def fake(name, *a, **k):
        if name == "docx" or name.startswith("docx."):
            raise ImportError("no docx")
        return real(name, *a, **k)

    monkeypatch.setattr(builtins, "__import__", fake)
    assert cli.main(["check", str(src)]) == 2
    assert "pip install python-docx lxml" in capsys.readouterr().err


def test_out_is_directory(make_doc, tmp_path, capsys):
    src = _src(make_doc)
    d = tmp_path / "โฟลเดอร์"
    d.mkdir()
    assert cli.main(["fix", str(src), str(d)]) == 2
    err = capsys.readouterr().err
    assert "ข้อผิดพลาด" in err and "โฟลเดอร์" in err and "Traceback" not in err


def test_save_oserror_is_thai_error(make_doc, tmp_path, monkeypatch, capsys):
    import os
    src = _src(make_doc)

    def boom(*a, **k):
        raise PermissionError(13, "Permission denied")

    monkeypatch.setattr(os, "replace", boom)
    out = tmp_path / "o.docx"
    assert cli.main(["wrap", str(src), str(out)]) == 2
    err = capsys.readouterr().err
    assert "บันทึกไฟล์" in err and "Traceback" not in err
    assert not out.exists() and not (tmp_path / "o.docx.tmp").exists()


def test_fix_prints_paragraph_label(make_doc, tmp_path, capsys):
    src = _src(make_doc)
    assert cli.main(["fix", str(src), str(tmp_path / "o.docx"), "--wrap-words"]) == 0
    assert "ย่อหน้า" in [l for l in capsys.readouterr().out.splitlines() if l.startswith("ตัดคำ")][0]


def test_fix_doc_option_and_spacing_report(make_doc, tmp_path, capsys):
    from docx import Document
    from docx.oxml.ns import qn
    doc = make_doc([("บทนำ", "Heading 1"), "ข้อความภาษาไทยเนื้อความ"])
    src = doc.saved_path
    out = tmp_path / "o.docx"
    assert cli.main(["fix", str(src), str(out), "--doc", "รายงาน"]) == 0
    text = capsys.readouterr().out
    assert "ระยะบรรทัด" in text and "ระยะก่อนย่อหน้า" in text
    ps = Document(str(out)).paragraphs
    assert [p._p.find(f"{qn('w:pPr')}/{qn('w:spacing')}").get(qn("w:before")) for p in ps] == ["120", "0"]
    assert cli.main(["check", str(out)]) == 0
    out2 = tmp_path / "o2.docx"
    assert cli.main(["fix", str(out), str(out2), "--doc", "report"]) == 0
    with pytest.raises(SystemExit) as e:
        cli.main(["fix", str(src), str(tmp_path / "o3.docx"), "--doc", "memo"])
    assert e.value.code == 2
