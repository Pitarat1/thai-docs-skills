"""คัดลอกแพ็กเกจ src/thaidocs เข้าโฟลเดอร์ skill และสร้างไฟล์ zip สำหรับแจก

  python tools/build.py                      คัดลอก src/thaidocs -> skills/<skill>/scripts/thaidocs (+ NOTICE, LICENSE) ทุก skill
                                             (thai-docx ไม่รวม official/ ส่วน thai-official-letters รวมทั้งแพ็กเกจ)
  python tools/build.py --check              ตรวจว่าโฟลเดอร์ skill ทุกตัวตรงกับ src/ (ไม่ตรง -> exit 1 พร้อมรายการ)
  python tools/build.py --zip [--dist DIR]   คัดลอกแล้วสร้าง <skill>.zip และ <skill>-nokhammoon.zip ต่อ skill
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
from typing import NamedTuple

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src" / "thaidocs"
SKILLS_DIR = ROOT / "plugins" / "thai-docs" / "skills"
NOTICE = ROOT / "NOTICE"
LICENSE = ROOT / "LICENSE"
SHIPPED = (NOTICE, LICENSE)  # ไฟล์ที่คัดลอกไปไว้ราก skill
DIST = ROOT / "dist"
KHAMMOON = "_khammoon"
ZIP_TIME = (1980, 1, 1, 0, 0, 0)

# ไม่คัดลอก/ไม่ใส่ zip/ไม่นับตอน --check: แคช Python, metadata ของ setuptools, ไฟล์ระบบของ macOS,
# ไฟล์ที่ iCloud ยังไม่ดาวน์โหลด (*.icloud) และสำเนาขัดแย้งของ iCloud ("x 2.py", "x 2")
EXCLUDE = ("__pycache__", "*.pyc", "*.egg-info", ".DS_Store", "._*", "*.icloud", "* 2", "* 2.*", "tests")


class SkillSpec(NamedTuple):
    name: str
    exclude: tuple[str, ...] = ()  # โฟลเดอร์ระดับบนสุดใน src/thaidocs ที่ไม่คัดลอกเข้า skill นี้

    @property
    def dir(self) -> Path:
        return SKILLS_DIR / self.name

    @property
    def dest(self) -> Path:
        return self.dir / "scripts" / "thaidocs"


SKILLS: dict[str, SkillSpec] = {
    "thai-docx": SkillSpec("thai-docx", exclude=("official",)),
    "thai-official-letters": SkillSpec("thai-official-letters"),
}


def _icloud_copy(name: str) -> bool:
    return Path(name).stem.endswith(" 2")


def _excluded(rel: Path, without_khammoon: bool, exclude: tuple[str, ...] = ()) -> bool:
    parts = rel.parts
    if parts and parts[0] in exclude:
        return True
    if without_khammoon and parts and parts[0] == KHAMMOON:
        return True
    return any(_icloud_copy(p) or any(fnmatch.fnmatch(p, pat) for pat in EXCLUDE) for p in parts)


def files(root: Path, without_khammoon: bool = False, exclude: tuple[str, ...] = ()) -> dict[str, Path]:
    """ไฟล์ทั้งหมดใต้ root (ตัดรายการที่ไม่คัดลอก) {path แบบ posix สัมพัทธ์: path จริง}"""
    out = {}
    if not root.is_dir():
        return out
    for p in sorted(root.rglob("*")):
        rel = p.relative_to(root)
        if p.is_file() and not _excluded(rel, without_khammoon, exclude):
            out[rel.as_posix()] = p
    return out


def sync(skill: SkillSpec, src: Path = SRC, dst: Path | None = None, without_khammoon: bool = False) -> int:
    """ลบ dst ทั้งโฟลเดอร์แล้วคัดลอกไฟล์จาก src คืนจำนวนไฟล์"""
    dst = skill.dest if dst is None else dst
    if dst.exists():
        shutil.rmtree(dst)
    todo = files(src, without_khammoon, skill.exclude)
    for rel, p in todo.items():
        target = dst / rel
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(p, target)
    return len(todo)


def diff(skill: SkillSpec, src: Path = SRC, dst: Path | None = None, without_khammoon: bool = False) -> list[str]:
    """รายการความต่างระหว่าง src กับ dst (ว่าง = ตรงกัน) ไม่นับ __pycache__ ใน dst"""
    dst = skill.dest if dst is None else dst
    a, b = files(src, without_khammoon, skill.exclude), files(dst)
    out = [f"ขาดใน skill: {r}" for r in sorted(a.keys() - b.keys())]
    out += [f"เกินใน skill: {r}" for r in sorted(b.keys() - a.keys())]
    out += [f"เนื้อหาไม่ตรง: {r}" for r in sorted(a.keys() & b.keys()) if a[r].read_bytes() != b[r].read_bytes()]
    return out


def sync_notice(skill: Path) -> None:
    skill.mkdir(parents=True, exist_ok=True)
    for f in SHIPPED:
        shutil.copy2(f, skill / f.name)


def diff_notice(skill: Path) -> list[str]:
    out = []
    for f in SHIPPED:
        target = skill / f.name
        if not target.is_file() or target.read_bytes() != f.read_bytes():
            out.append(f"เนื้อหาไม่ตรง: {f.name}")
    return out


def write_zip(skill: SkillSpec, skill_dir: Path, out: Path, without_khammoon: bool = False) -> int:
    """zip โฟลเดอร์ skill_dir โดยมีชื่อ skill เป็นโฟลเดอร์ราก เรียงชื่อ และเวลาคงที่ (ได้ไบต์เดิมทุกครั้ง)"""
    entries = {}
    for rel, p in files(skill_dir).items():
        if without_khammoon and rel.startswith(f"scripts/thaidocs/{KHAMMOON}/"):
            continue
        entries[f"{skill.name}/{rel}"] = p
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
        d = []
        for sk in SKILLS.values():
            d += [f"{sk.name}: {x}" for x in diff(sk, without_khammoon=a.without_khammoon) + diff_notice(sk.dir)]
        if d:
            print("โฟลเดอร์ skill ไม่ตรงกับ src/ ให้รัน python tools/build.py แล้ว commit ใหม่:")
            print("\n".join(f"  {x}" for x in d))
            return 1
        print("โฟลเดอร์ skill ตรงกับ src/")
        return 0

    note = " (ไม่รวม _khammoon)" if a.without_khammoon else ""
    for sk in SKILLS.values():
        n = sync(sk, without_khammoon=a.without_khammoon)
        sync_notice(sk.dir)
        print(f"คัดลอก {n} ไฟล์ไปที่ {sk.dest.relative_to(ROOT)}{note}")
        if a.zip:
            targets = [(f"{sk.name}-nokhammoon.zip", True)]
            if not a.without_khammoon:
                targets.insert(0, (f"{sk.name}.zip", False))
            for name, nok in targets:
                size = write_zip(sk, sk.dir, a.dist / name, without_khammoon=nok)
                print(f"สร้าง {a.dist / name} ({size / 1024 / 1024:.2f} MB)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
