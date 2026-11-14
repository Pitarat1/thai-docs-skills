"""เทส tools/build.py, โฟลเดอร์ skill thai-docx และ manifest ของ plugin"""
from __future__ import annotations

import argparse
import importlib.util
import json
import re
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


OFFICIAL = ROOT / "plugins" / "thai-docs" / "skills" / "thai-official-letters"


def _build_module():
    spec = importlib.util.spec_from_file_location("thaidocs_build", BUILD)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def _docx(build):
    return build.SKILLS["thai-docx"]


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
    build.sync(_docx(build), src, dst, without_khammoon=True)
    assert build.diff(_docx(build), src, dst, without_khammoon=True) == []
    (dst / "cli.py").write_text("# แก้\n", encoding="utf-8")
    (src / "new.py").write_text("x = 1\n", encoding="utf-8")
    (dst / "stale.py").write_text("y = 2\n", encoding="utf-8")
    d = build.diff(_docx(build), src, dst, without_khammoon=True)
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
    sk = _docx(build)
    build.sync(sk, src, dst)
    (src / name).write_text("junk\n", encoding="utf-8")
    (dst / name).write_text("other junk\n", encoding="utf-8")
    assert build.diff(sk, src, dst) == []
    build.sync(sk, src, dst)
    assert not (dst / name).exists()


def _zips(build, skill: Path, dist: Path) -> tuple[Path, Path]:
    full, nok = dist / "thai-docx.zip", dist / "thai-docx-nokhammoon.zip"
    build.write_zip(_docx(build), skill, full)
    build.write_zip(_docx(build), skill, nok, without_khammoon=True)
    return full, nok


def test_zip_layout(tmp_path):
    # สร้าง skill ใน tmp_path (สำเนาโฟลเดอร์ที่ commit + sync src ใหม่) ไม่แตะโฟลเดอร์ skill ใน repo
    build = _build_module()
    skill = tmp_path / "thai-docx"
    shutil.copytree(SKILL, skill, ignore=shutil.ignore_patterns("__pycache__", "thaidocs"))
    build.sync(_docx(build), build.SRC, skill / "scripts" / "thaidocs")
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


_VENDOR_NAMES = ("Claude", "Codex", "Antigravity", "ChatGPT", "Gemini", "OpenAI", "Anthropic", "Copilot")


def _skill_docs(skill: Path) -> list[Path]:
    return [skill / "SKILL.md", *sorted((skill / "references").glob("*.md"))]


@pytest.mark.parametrize("skill", [SKILL, OFFICIAL], ids=lambda p: p.name)
def test_skill_references_exist(skill):
    # ทุกไฟล์ references/… ที่ SKILL.md อ้าง ต้องมีอยู่จริง
    text = (skill / "SKILL.md").read_text(encoding="utf-8")
    refs = set(re.findall(r"references/([\w.-]+\.md)", text))
    assert refs
    for ref in refs:
        assert (skill / "references" / ref).is_file(), ref


@pytest.mark.parametrize("skill", [SKILL, OFFICIAL], ids=lambda p: p.name)
def test_skill_docs_vendor_neutral(skill):
    # เนื้อหา skill ใช้กับ AI ตัวใดก็ได้: ไม่อ้างชื่อผลิตภัณฑ์/ผู้ผลิต AI
    for path in _skill_docs(skill):
        text = path.read_text(encoding="utf-8").lower()
        for name in _VENDOR_NAMES:
            assert name.lower() not in text, f"{path.name}: {name}"


def test_official_skill_frontmatter():
    from thaidocs.official.schema import DOC_TYPE_TH
    text = (OFFICIAL / "SKILL.md").read_text(encoding="utf-8")
    fm = _frontmatter(text)
    assert set(fm) == {"name", "description"}
    assert fm["name"] == "thai-official-letters"
    assert 0 < len(fm["description"]) <= 1024
    assert len(fm["description"].encode("utf-8")) <= 1024
    assert sum(th in fm["description"] for th in DOC_TYPE_TH.values()) >= 5
    assert len(text.splitlines()) <= 200
    for ref in ("doc-types.md", "writing.md", "layout-w2019.md", "agency-profile.md", "preview.md"):
        assert (OFFICIAL / "references" / ref).is_file()
        assert f"references/{ref}" in text


def test_official_skill_rules_text():
    text = (OFFICIAL / "SKILL.md").read_text(encoding="utf-8")
    assert "pip install python-docx lxml pydantic" in text
    assert "ห้ามแต่งชื่อ ตำแหน่ง เลขที่หนังสือ วันที่ ที่ผู้ใช้ไม่ได้ให้ ให้ใส่ `[ระบุ…]`" in text
    assert "บอกผู้ใช้ว่าต้องใส่ตราเอง ออกเลข และลงนามเอง" in text


def _subcommands(parser: argparse.ArgumentParser) -> dict[str, argparse.ArgumentParser]:
    (sub,) = [a for a in parser._actions if isinstance(a, argparse._SubParsersAction)]
    return dict(sub.choices)


def test_official_skill_documents_cli():
    # ทุกคำสั่งและตัวเลือกของ thaigov ต้องปรากฏใน SKILL.md
    from thaidocs.official.cli import _parser
    text = (OFFICIAL / "SKILL.md").read_text(encoding="utf-8")
    cmds = _subcommands(_parser())
    assert set(cmds) == {"types", "schema", "example", "render", "check"}
    for name, sub in cmds.items():
        assert f"`{name}" in text, name
        for action in sub._actions:
            for opt in action.option_strings:
                if opt.startswith("--") and opt != "--help":
                    assert re.search(rf"`[^`]*{re.escape(opt)}(?![\w-])", text), f"{name} {opt}"


def test_official_finding_codes_documented():
    from thaidocs.official.checker import ALL_CODES
    text = (OFFICIAL / "SKILL.md").read_text(encoding="utf-8") + \
        (OFFICIAL / "references" / "doc-types.md").read_text(encoding="utf-8")
    missing = [c for c in ALL_CODES if f"`{c}`" not in text]
    assert not missing, missing


def _json_blocks(text: str) -> list[str]:
    return re.findall(r"```json\n(.*?)```", text, flags=re.S)


def test_official_skill_example_draft_is_valid():
    # ร่างตัวอย่างใน SKILL.md ต้องผ่านการตรวจโครง และใส่ [ระบุ…] แทนสิ่งที่ผู้ใช้ไม่ได้ให้
    from thaidocs.official.schema import parse_doc
    (block,) = _json_blocks((OFFICIAL / "SKILL.md").read_text(encoding="utf-8"))
    assert "[ระบุ" in block
    parse_doc(json.loads(block))


def test_agency_profile_doc(tmp_path):
    # ตัวอย่าง agency.json ในเอกสารต้องอ่านได้จริง และเอกสารอธิบายทุก key ที่รับรอง
    from dataclasses import fields

    from thaidocs.official.profile import AgencyInfo, load_profile
    from thaidocs.official.style import DOCUMENTED_LAYOUT_KEYS
    text = (OFFICIAL / "references" / "agency-profile.md").read_text(encoding="utf-8")
    for key in (*DOCUMENTED_LAYOUT_KEYS, *(f.name for f in fields(AgencyInfo))):
        assert f"`{key}`" in text, key
    (block,) = _json_blocks(text)
    data = json.loads(block)
    data.get("layout", {}).pop("emblem_png", None)      # ไม่ต้องมีไฟล์ตราจริงในเทส
    p = tmp_path / "agency.json"
    p.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    agency, _ = load_profile(p)
    assert agency is not None and agency.org


def test_layout_doc_covers_spec_defaults():
    # ตารางค่าการพิมพ์ต้องมีทุก field ของ LayoutSpec ที่มีป้ายที่มา ([ว2019] [ระเบียบ] [แผ่นแบบ] [ปฏิบัติ]) ใน style.py
    src = (ROOT / "src" / "thaidocs" / "official" / "style.py").read_text(encoding="utf-8")
    tagged = re.findall(r"^\s+(\w+):.*#\s*\[(?:ว2019|ระเบียบ|แผ่นแบบ|ปฏิบัติ)", src, flags=re.M)
    assert len(tagged) >= 15
    text = (OFFICIAL / "references" / "layout-w2019.md").read_text(encoding="utf-8")
    missing = [k for k in tagged if f"`{k}`" not in text]
    assert not missing, missing


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


def test_skills_table():
    build = _build_module()
    assert build.SKILLS["thai-docx"].exclude == ("official",)
    assert build.SKILLS["thai-official-letters"].exclude == ()


def test_thai_docx_has_no_official():
    assert not (SKILL / "scripts" / "thaidocs" / "official").exists()


def test_official_skill_has_full_package():
    pkg = OFFICIAL / "scripts" / "thaidocs"
    assert (pkg / "official" / "__init__.py").is_file()
    assert (pkg / "layout.py").is_file()
    assert (pkg / "data" / "words_th.txt").is_file()
    assert (OFFICIAL / "scripts" / "thaigov.py").is_file()
    for f in ("NOTICE", "LICENSE"):
        assert (OFFICIAL / f).read_bytes() == (ROOT / f).read_bytes()


def test_check_covers_all_skills(tmp_path):
    # --check ต้องตรวจสำเนาของ skill ใหม่ด้วย: แก้ไฟล์ในสำเนาแล้วต้องคืน 1 (คืนไฟล์เดิมในตอนท้าย)
    target = OFFICIAL / "scripts" / "thaidocs" / "layout.py"
    original = target.read_bytes()
    try:
        target.write_text("# แก้\n", encoding="utf-8")
        r = _run(BUILD, "--check")
        assert r.returncode == 1, r.stdout + r.stderr
        assert "layout.py" in r.stdout
    finally:
        target.write_bytes(original)
    assert _run(BUILD, "--check").returncode == 0


def test_zip_per_skill(tmp_path):
    r = _run(BUILD, "--zip", "--dist", tmp_path)
    assert r.returncode == 0, r.stdout + r.stderr
    names = sorted(p.name for p in tmp_path.glob("*.zip"))
    assert names == ["thai-docx-nokhammoon.zip", "thai-docx.zip",
                     "thai-official-letters-nokhammoon.zip", "thai-official-letters.zip"]
    for n, root in (("thai-docx.zip", "thai-docx/"), ("thai-official-letters.zip", "thai-official-letters/")):
        zn = zipfile.ZipFile(tmp_path / n).namelist()
        assert all(x.startswith(root) for x in zn)
        assert (root + "NOTICE") in zn and (root + "LICENSE") in zn
    zn = zipfile.ZipFile(tmp_path / "thai-official-letters.zip").namelist()
    assert "thai-official-letters/scripts/thaigov.py" in zn
    assert "thai-official-letters/scripts/thaidocs/official/__init__.py" in zn
    assert not any("/official/" in x for x in zipfile.ZipFile(tmp_path / "thai-docx.zip").namelist())


def test_official_script_runs(tmp_path):
    # แตก zip ของ skill ลง tmp_path แล้วรันสคริปต์จากที่นั่น (cwd นอก repo, ไม่มี PYTHONPATH) ว่าใช้งานได้จริง
    build = _build_module()
    skill = tmp_path / "pack" / "thai-official-letters"
    shutil.copytree(OFFICIAL, skill, ignore=shutil.ignore_patterns("__pycache__", "thaidocs"))
    build.sync(build.SKILLS["thai-official-letters"], build.SRC, skill / "scripts" / "thaidocs")
    zpath = tmp_path / "dist" / "thai-official-letters.zip"
    build.write_zip(build.SKILLS["thai-official-letters"], skill, zpath, without_khammoon=True)
    out = tmp_path / "extract"
    zipfile.ZipFile(zpath).extractall(out)
    script = out / "thai-official-letters" / "scripts" / "thaigov.py"
    env = {k: v for k, v in __import__("os").environ.items() if k != "PYTHONPATH"}
    env["HOME"] = env["USERPROFILE"] = str(tmp_path / "home")
    cmd = [sys.executable, "-B", str(script)]
    run = lambda *a: subprocess.run([*cmd, *a], cwd=tmp_path, env=env, capture_output=True, text=True, encoding="utf-8")
    ex = run("example", "memo")
    assert ex.returncode == 0, ex.stderr
    (tmp_path / "memo.json").write_text(ex.stdout, encoding="utf-8")
    r = run("render", "memo.json", "out.docx")
    assert r.returncode == 0, r.stderr
    assert "สร้างไฟล์" in r.stdout
    assert (tmp_path / "out.docx").is_file()
