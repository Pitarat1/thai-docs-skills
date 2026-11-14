"""เทส tools/build.py, โฟลเดอร์ skill thai-docx และ manifest ของ plugin"""
from __future__ import annotations

import importlib.util
import json
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
BUILD = ROOT / "tools" / "build.py"
SKILL = ROOT / "plugins" / "thai-docs" / "skills" / "thai-docx"
SCRIPT = SKILL / "scripts" / "thaidocx.py"


def _build_module():
    spec = importlib.util.spec_from_file_location("thaidocs_build", BUILD)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def _run(*args, cwd=ROOT):
    return subprocess.run([sys.executable, "-B", *map(str, args)], cwd=cwd,
                          capture_output=True, text=True, encoding="utf-8")


def test_build_check_in_sync():
    # ตรวจอย่างเดียว (ไม่ sync ก่อน) โฟลเดอร์ skill ที่ commit ต้องตรงกับ src/
    r = _run(BUILD, "--check")
    assert r.returncode == 0, r.stdout + r.stderr
    assert not (SKILL / "scripts" / "thaidocs" / "__pycache__").exists()
    assert not (SKILL / "scripts" / "thaidocs" / "tests").exists()


def test_check_detects_drift(tmp_path):
    build = _build_module()
    src = tmp_path / "src"
    shutil.copytree(ROOT / "src" / "thaidocs", src,
                    ignore=shutil.ignore_patterns("__pycache__", "*.pyc", "_khammoon"))
    dst = tmp_path / "dst"
    build.sync(src, dst, without_khammoon=True)
    assert build.diff(src, dst, without_khammoon=True) == []
    (dst / "cli.py").write_text("# แก้\n", encoding="utf-8")
    (src / "new.py").write_text("x = 1\n", encoding="utf-8")
    (dst / "stale.py").write_text("y = 2\n", encoding="utf-8")
    d = build.diff(src, dst, without_khammoon=True)
    assert any("cli.py" in x for x in d)
    assert any("new.py" in x for x in d)
    assert any("stale.py" in x for x in d)


@pytest.mark.parametrize("name", ["cli 2.py", "data 2", "x.py.icloud", "._cli.py", ".DS_Store"])
def test_junk_files_ignored(tmp_path, name):
    build = _build_module()
    src = tmp_path / "src"
    src.mkdir()
    (src / "a.py").write_text("a = 1\n", encoding="utf-8")
    dst = tmp_path / "dst"
    build.sync(src, dst)
    (src / name).write_text("junk\n", encoding="utf-8")
    (dst / name).write_text("other junk\n", encoding="utf-8")
    assert build.diff(src, dst) == []
    build.sync(src, dst)
    assert not (dst / name).exists()


def _zips(build, skill: Path, dist: Path) -> tuple[Path, Path]:
    full, nok = dist / "thai-docx.zip", dist / "thai-docx-nokhammoon.zip"
    build.write_zip(skill, full)
    build.write_zip(skill, nok, without_khammoon=True)
    return full, nok


def test_zip_layout(tmp_path):
    # สร้าง skill ใน tmp_path (สำเนาโฟลเดอร์ที่ commit + sync src ใหม่) ไม่แตะโฟลเดอร์ skill ใน repo
    build = _build_module()
    skill = tmp_path / "thai-docx"
    shutil.copytree(SKILL, skill, ignore=shutil.ignore_patterns("__pycache__", "thaidocs"))
    build.sync(build.SRC, skill / "scripts" / "thaidocs")
    full, nok = _zips(build, skill, tmp_path / "dist")
    for z in (full, nok):
        assert z.stat().st_size < 10 * 1024 * 1024
        names = zipfile.ZipFile(z).namelist()
        assert "thai-docx/SKILL.md" in names
        assert "thai-docx/LICENSE" in names
        assert "thai-docx/NOTICE" in names
        assert "thai-docx/scripts/thaidocx.py" in names
        assert "thai-docx/scripts/thaidocs/data/words_th.txt" in names
        assert all(n.startswith("thai-docx/") for n in names)
        assert not any("__pycache__" in n or n.endswith(".pyc") for n in names)
        assert names == sorted(names)
    assert any(n.endswith("_khammoon/model/weights.npz") for n in zipfile.ZipFile(full).namelist())
    assert not any("_khammoon" in n for n in zipfile.ZipFile(nok).namelist())
    # zip ซ้ำได้ไบต์เดิม (ไม่ฝังเวลา)
    first = full.read_bytes()
    _zips(build, skill, tmp_path / "dist")
    assert full.read_bytes() == first


def _frontmatter(text: str) -> dict[str, str]:
    assert text.startswith("---\n")
    head = text.split("---\n", 2)[1]
    out = {}
    for line in head.splitlines():
        k, _, v = line.partition(":")
        out[k.strip()] = v.strip()
    return out


def test_skill_frontmatter():
    text = (SKILL / "SKILL.md").read_text(encoding="utf-8")
    fm = _frontmatter(text)
    assert set(fm) == {"name", "description"}
    assert fm["name"] == "thai-docx"
    assert 0 < len(fm["description"]) <= 1024
    assert len(fm["description"].encode("utf-8")) <= 1024
    assert len(text.splitlines()) <= 150
    for ref in ("layout-rules.md", "troubleshooting.md", "wrap-words.md", "preview.md"):
        assert (SKILL / "references" / ref).is_file()
        assert f"references/{ref}" in text


def test_manifests():
    m = json.loads((ROOT / ".claude-plugin" / "marketplace.json").read_text(encoding="utf-8"))
    assert m["name"] == "thai-docs-skills"
    assert m["owner"]["name"] == "Pitarat1"
    (plug,) = m["plugins"]
    assert plug["name"] == "thai-docs" and plug["source"] == "./plugins/thai-docs"
    p = json.loads((ROOT / "plugins" / "thai-docs" / ".claude-plugin" / "plugin.json").read_text(encoding="utf-8"))
    assert p["name"] == "thai-docs" and p["version"] == "0.1.0" and p["description"]


_PROBE = """
import runpy, sys
script = sys.argv[1]
sys.argv = [script, "segment", "ทดสอบ"]
try:
    runpy.run_path(script, run_name="__main__")
except SystemExit as e:
    code = e.code
import thaidocs
print("MODULE", thaidocs.__file__)
sys.exit(code)
"""


def test_skill_script_runs(tmp_path):
    # -I: ไม่อ่าน PYTHONPATH / user site / cwd  -S: ไม่โหลด site-packages (จึงไม่เห็น editable install ของ src/)
    # -B: ไม่สร้าง __pycache__ ในโฟลเดอร์ skill  cwd อยู่นอก repo
    r = subprocess.run([sys.executable, "-I", "-S", "-B", str(SCRIPT), "segment", "ทดสอบ"],
                       cwd=tmp_path, capture_output=True, text=True, encoding="utf-8")
    assert r.returncode == 0, r.stderr
    assert "ทดสอบ" in r.stdout
    # มี site-packages (numpy, editable install) แต่โมดูลที่โหลดต้องมาจากโฟลเดอร์ skill ไม่ใช่ src/
    r = subprocess.run([sys.executable, "-I", "-B", "-c", _PROBE, str(SCRIPT)],
                       cwd=tmp_path, capture_output=True, text=True, encoding="utf-8")
    assert r.returncode == 0, r.stderr
    mod = r.stdout.split("MODULE", 1)[1].strip()
    assert Path(mod).resolve().is_relative_to((SKILL / "scripts" / "thaidocs").resolve()), mod


def test_license_in_skill_and_checked(tmp_path):
    build = _build_module()
    assert (SKILL / "LICENSE").read_bytes() == (ROOT / "LICENSE").read_bytes()
    skill = tmp_path / "thai-docx"
    skill.mkdir()
    build.sync_notice(skill)
    assert (skill / "LICENSE").read_bytes() == (ROOT / "LICENSE").read_bytes()
    assert build.diff_notice(skill) == []
    (skill / "LICENSE").write_text("อื่น\n", encoding="utf-8")
    assert any("LICENSE" in x for x in build.diff_notice(skill))
    (skill / "LICENSE").unlink()
    assert any("LICENSE" in x for x in build.diff_notice(skill))
