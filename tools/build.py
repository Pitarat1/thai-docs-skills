"""คัดลอกแพ็กเกจ src/thaidocs เข้าโฟลเดอร์ skill และสร้างไฟล์ zip สำหรับแจก

  python tools/build.py                      คัดลอก src/thaidocs -> skills/thai-docx/scripts/thaidocs (+ NOTICE, LICENSE)
  python tools/build.py --check              ตรวจว่าโฟลเดอร์ skill ตรงกับ src/ (ไม่ตรง -> exit 1 พร้อมรายการ)
  python tools/build.py --zip [--dist DIR]   คัดลอกแล้วสร้าง thai-docx.zip และ thai-docx-nokhammoon.zip
  --without-khammoon                         ไม่คัดลอก _khammoon/ (โมเดล CC BY-NC-SA สำหรับใช้เชิงพาณิชย์)
                                             ใช้ร่วมกับ --check ได้ (ตรวจแบบไม่มี _khammoon)
"""
from __future__ import annotations

import argparse
import fnmatch
import shutil
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src" / "thaidocs"
SKILL = ROOT / "plugins" / "thai-docs" / "skills" / "thai-docx"
DEST = SKILL / "scripts" / "thaidocs"
NOTICE = ROOT / "NOTICE"
LICENSE = ROOT / "LICENSE"
SHIPPED = (NOTICE, LICENSE)  # ไฟล์ที่คัดลอกไปไว้ราก skill
DIST = ROOT / "dist"
KHAMMOON = "_khammoon"
ZIP_ROOT = "thai-docx"
ZIP_TIME = (1980, 1, 1, 0, 0, 0)

# ไม่คัดลอก/ไม่ใส่ zip/ไม่นับตอน --check: แคช Python, metadata ของ setuptools, ไฟล์ระบบของ macOS,
# ไฟล์ที่ iCloud ยังไม่ดาวน์โหลด (*.icloud) และสำเนาขัดแย้งของ iCloud ("x 2.py", "x 2")
EXCLUDE = ("__pycache__", "*.pyc", "*.egg-info", ".DS_Store", "._*", "*.icloud", "* 2", "* 2.*", "tests")


def _icloud_copy(name: str) -> bool:
    return Path(name).stem.endswith(" 2")


def _excluded(rel: Path, without_khammoon: bool) -> bool:
    parts = rel.parts
    if without_khammoon and parts and parts[0] == KHAMMOON:
        return True
    return any(_icloud_copy(p) or any(fnmatch.fnmatch(p, pat) for pat in EXCLUDE) for p in parts)


def files(root: Path, without_khammoon: bool = False) -> dict[str, Path]:
    """ไฟล์ทั้งหมดใต้ root (ตัดรายการที่ไม่คัดลอก) {path แบบ posix สัมพัทธ์: path จริง}"""
    out = {}
    if not root.is_dir():
        return out
    for p in sorted(root.rglob("*")):
        rel = p.relative_to(root)
        if p.is_file() and not _excluded(rel, without_khammoon):
            out[rel.as_posix()] = p
    return out


def sync(src: Path = SRC, dst: Path = DEST, without_khammoon: bool = False) -> int:
    """ลบ dst ทั้งโฟลเดอร์แล้วคัดลอกไฟล์จาก src คืนจำนวนไฟล์"""
    if dst.exists():
        shutil.rmtree(dst)
    todo = files(src, without_khammoon)
    for rel, p in todo.items():
        target = dst / rel
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(p, target)
    return len(todo)


def diff(src: Path = SRC, dst: Path = DEST, without_khammoon: bool = False) -> list[str]:
    """รายการความต่างระหว่าง src กับ dst (ว่าง = ตรงกัน) ไม่นับ __pycache__ ใน dst"""
    a, b = files(src, without_khammoon), files(dst)
    out = [f"ขาดใน skill: {r}" for r in sorted(a.keys() - b.keys())]
    out += [f"เกินใน skill: {r}" for r in sorted(b.keys() - a.keys())]
    out += [f"เนื้อหาไม่ตรง: {r}" for r in sorted(a.keys() & b.keys()) if a[r].read_bytes() != b[r].read_bytes()]
    return out


def sync_notice(skill: Path = SKILL) -> None:
    skill.mkdir(parents=True, exist_ok=True)
    for f in SHIPPED:
        shutil.copy2(f, skill / f.name)


def diff_notice(skill: Path = SKILL) -> list[str]:
    out = []
    for f in SHIPPED:
        target = skill / f.name
        if not target.is_file() or target.read_bytes() != f.read_bytes():
            out.append(f"เนื้อหาไม่ตรง: {f.name}")
    return out


def write_zip(skill: Path, out: Path, without_khammoon: bool = False) -> int:
    """zip โฟลเดอร์ skill โดยมี thai-docx/ เป็นโฟลเดอร์ราก เรียงชื่อ และเวลาคงที่ (ได้ไบต์เดิมทุกครั้ง)"""
    entries = {}
    for rel, p in files(skill).items():
        if without_khammoon and rel.startswith(f"scripts/thaidocs/{KHAMMOON}/"):
            continue
        entries[f"{ZIP_ROOT}/{rel}"] = p
    out.parent.mkdir(parents=True, exist_ok=True)
    tmp = out.with_name(out.name + ".tmp")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as z:
        for name in sorted(entries):
            info = zipfile.ZipInfo(name, date_time=ZIP_TIME)
            info.compress_type = zipfile.ZIP_DEFLATED
            info.create_system = 3
            info.external_attr = 0o644 << 16
            z.writestr(info, entries[name].read_bytes())
    tmp.replace(out)
    return out.stat().st_size


def main(argv: list[str] | None = None) -> int:
    for s in (sys.stdout, sys.stderr):
        try:
            s.reconfigure(encoding="utf-8")
        except Exception:
            pass
    ap = argparse.ArgumentParser(description="คัดลอก src/thaidocs เข้าโฟลเดอร์ skill และสร้าง zip")
    ap.add_argument("--check", action="store_true", help="ตรวจอย่างเดียว ไม่แก้ไฟล์")
    ap.add_argument("--zip", action="store_true", help="สร้าง zip ใน dist/")
    ap.add_argument("--dist", type=Path, default=DIST, help="โฟลเดอร์ปลายทางของ zip")
    ap.add_argument("--without-khammoon", action="store_true", help="ไม่รวม _khammoon (โมเดลห้ามใช้เชิงพาณิชย์)")
    a = ap.parse_args(argv)

    if a.check:
        d = diff(without_khammoon=a.without_khammoon) + diff_notice()
        if d:
            print("โฟลเดอร์ skill ไม่ตรงกับ src/ ให้รัน python tools/build.py แล้ว commit ใหม่:")
            print("\n".join(f"  {x}" for x in d))
            return 1
        print("โฟลเดอร์ skill ตรงกับ src/")
        return 0

    n = sync(without_khammoon=a.without_khammoon)
    sync_notice()
    note = " (ไม่รวม _khammoon)" if a.without_khammoon else ""
    print(f"คัดลอก {n} ไฟล์ไปที่ {DEST.relative_to(ROOT)}{note}")
    if a.zip:
        targets = [("thai-docx-nokhammoon.zip", True)]
        if not a.without_khammoon:
            targets.insert(0, ("thai-docx.zip", False))
        for name, nok in targets:
            size = write_zip(SKILL, a.dist / name, without_khammoon=nok)
            print(f"สร้าง {a.dist / name} ({size / 1024 / 1024:.2f} MB)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
