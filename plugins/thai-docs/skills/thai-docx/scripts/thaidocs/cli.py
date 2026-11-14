"""thaidocx: คำสั่งจัดหน้าและตัดคำไทยสำหรับเอกสาร Word

  fix IN OUT [--font] [--body-pt] [--thai-digits] [--distribute]
             [--remove-manual-breaks] [--doc ราชการ|รายงาน] [--wrap-words] [--words FILE ...]
  check IN [--thai-digits]
  wrap IN OUT [--words FILE ...]      unwrap IN OUT
  segment TEXT [--words FILE ...]

ข้อผิดพลาดแสดงเป็นภาษาไทยทาง stderr และคืน exit code 2
"""
from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path


class CliError(Exception):
    pass


def _utf8() -> None:
    for s in (sys.stdout, sys.stderr):
        try:
            s.reconfigure(encoding="utf-8")
        except Exception:
            pass


def _load_docx():
    try:
        import docx  # noqa: F401
        import lxml  # noqa: F401
    except ImportError:
        raise CliError("ไม่พบไลบรารีที่จำเป็น ติดตั้งด้วยคำสั่ง: pip install python-docx lxml") from None
    return docx


def _open(path: Path):
    docx = _load_docx()
    if not path.is_file():
        raise CliError(f"ไม่พบไฟล์ {path}")
    try:
        return docx.Document(str(path))
    except Exception as e:
        raise CliError(f"เปิดไฟล์ {path} ไม่ได้ (ไม่ใช่ .docx ที่อ่านได้): {e}") from None


def _check_paths(src: Path, out: Path | None) -> None:
    if not src.is_file():
        raise CliError(f"ไม่พบไฟล์ {src}")
    name = src.name
    for lock in (src.parent / ("~$" + name), src.parent / ("~$" + name[2:])):
        if lock.exists():
            raise CliError(f"ไฟล์ {src.name} ยังเปิดอยู่ใน Word (พบไฟล์ล็อก {lock.name}) กรุณาปิดไฟล์ก่อนแล้วสั่งใหม่")
    if out is not None:
        same = src.resolve() == out.resolve()
        if not same and out.exists():
            try:
                same = os.path.samefile(src, out)
            except OSError:
                pass
        if same:
            raise CliError("ไฟล์ปลายทางต้องเป็นคนละไฟล์กับไฟล์ต้นทาง (ไม่เขียนทับไฟล์เดิม)")
        if out.is_dir():
            raise CliError(f"ไฟล์ปลายทาง {out} เป็นโฟลเดอร์ ให้ระบุชื่อไฟล์ .docx")
        if not out.parent.is_dir():
            raise CliError(f"ไม่พบโฟลเดอร์ปลายทาง {out.parent}")


def _segmenter(words):
    from .dictionary import load_words
    from .wordcut import Segmenter
    try:
        return Segmenter(words=load_words(words or ()))
    except OSError as e:
        raise CliError(f"อ่านไฟล์คำไม่ได้: {e}") from None


def _save(doc, out: Path) -> None:
    tmp = out.with_name(out.name + ".tmp")
    try:
        doc.save(str(tmp))
        os.replace(tmp, out)
    except OSError as e:
        raise CliError(f"บันทึกไฟล์ {out} ไม่ได้: {e}") from None
    finally:
        if tmp.exists():
            tmp.unlink()


def _n(r) -> int:
    return r if isinstance(r, int) else 0


def cmd_fix(a) -> int:
    src, out = Path(a.IN), Path(a.OUT)
    _check_paths(src, out)
    doc = _open(src)
    from . import layout
    from .digits import apply_thai_digits
    from .ooxml import is_wrapped
    from . import wrap as wrapmod

    was_wrapped = is_wrapped(doc)
    lines = []
    if was_wrapped:
        lines.append(f"ถอดการตัดคำเดิมก่อนจัดหน้า ({_n(wrapmod.unwrap(doc))} ย่อหน้า)")
    lines.append(f"แก้อักษรเพี้ยน {_n(layout.normalize_chars(doc))} จุด")
    layout.set_font_all(doc, a.font)
    lines.append(f"ตั้งฟอนต์ทั้งเอกสารเป็น {a.font}")
    lines.append(f"ปรับขนาดอักษรไทยให้เท่าอักษรอังกฤษ {_n(layout.align_sz(doc))} จุด")
    if a.remove_manual_breaks:
        lines.append(f"ลบการขึ้นบรรทัดด้วยมือ {_n(layout.remove_manual_breaks(doc))} จุด")
    if a.thai_digits:
        lines.append(f"แปลงเป็นเลขไทย {_n(apply_thai_digits(doc))} จุด")
    if a.distribute:
        lines.append(f"ตั้งกระจายแบบไทย {_n(layout.distribute(doc))} ย่อหน้า")
    lines.append(f"ตรึงหัวข้อไว้กับย่อหน้าถัดไป {_n(layout.keep_next_headings(doc))} หัวข้อ")
    n_single = _n(layout.single_spacing(doc))
    n_exact = layout.count_exact_spacing(doc)
    lines.append(f"ตั้งระยะบรรทัดเดี่ยว (line 240, after 0) {n_single} จุด"
                 + (f" · คงระยะบรรทัดแบบ exact ไว้ {n_exact} จุด (ล็อกความสูงโดยตั้งใจ)" if n_exact else ""))
    if a.doc:
        lines.append(f"ตั้งระยะก่อนย่อหน้า 6 pt ตามชนิดเอกสาร ({a.doc}) {_n(layout.paragraph_spacing(doc, a.doc))} ย่อหน้า")
    layout.doc_defaults(doc, a.font, a.body_pt)
    lines.append(f"ตั้งค่าเริ่มต้นเอกสาร ({a.font} {a.body_pt} pt)")
    lines.append(f"ติดป้ายภาษาไทยให้ข้อความ {_n(layout.thai_runs(doc))} จุด")
    layout.hide_spelling(doc)
    if a.wrap_words or was_wrapped:
        seg = _segmenter(a.words)
        lines.append(f"ตัดคำไม่ให้ขาดกลางคำ {_n(wrapmod.wrap(doc, seg))} ย่อหน้า")
    _save(doc, out)
    print("\n".join(lines))
    print(f"บันทึกแล้ว: {out}")
    return 0


def cmd_check(a) -> int:
    src = Path(a.IN)
    doc = _open(src)
    from .check import check
    issues = check(doc, thai_digits=a.thai_digits)
    for i in issues:
        print(f"{i.code} {i.count} {i.message}")
    if not issues:
        print("ไม่พบปัญหา")
    return 1 if any(i.code != "WRAPPED" for i in issues) else 0


def cmd_wrap(a) -> int:
    src, out = Path(a.IN), Path(a.OUT)
    _check_paths(src, out)
    doc = _open(src)
    from . import wrap as wrapmod
    n = _n(wrapmod.wrap(doc, _segmenter(a.words)))
    _save(doc, out)
    print(f"ตัดคำ {n} ย่อหน้า บันทึกแล้ว: {out}")
    return 0


def cmd_unwrap(a) -> int:
    src, out = Path(a.IN), Path(a.OUT)
    _check_paths(src, out)
    doc = _open(src)
    from . import wrap as wrapmod
    n = _n(wrapmod.unwrap(doc))
    _save(doc, out)
    print(f"ถอดการตัดคำ {n} ย่อหน้า บันทึกแล้ว: {out}")
    return 0


def cmd_segment(a) -> int:
    print("|".join(_segmenter(a.words).words(a.TEXT)))
    return 0


def _parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(prog="thaidocx", description="จัดหน้าและตัดคำไทยสำหรับเอกสาร Word")
    sub = p.add_subparsers(dest="cmd", required=True)

    def words(s):
        s.add_argument("--words", action="append", default=[], metavar="FILE", help="ไฟล์คำเพิ่มเติม (บรรทัดละคำ) ใช้ซ้ำได้")

    s = sub.add_parser("fix", help="จัดหน้าทั้งไฟล์ตามหลัก 10 ข้อ")
    s.add_argument("IN"); s.add_argument("OUT")
    s.add_argument("--font", default="TH Sarabun New")
    s.add_argument("--body-pt", type=float, default=16, dest="body_pt")
    s.add_argument("--thai-digits", action="store_true")
    s.add_argument("--distribute", action="store_true")
    s.add_argument("--remove-manual-breaks", action="store_true")
    s.add_argument("--wrap-words", action="store_true")
    s.add_argument("--doc", choices=["ราชการ", "รายงาน", "official", "report"], default=None,
                   help="ชนิดเอกสาร: ราชการ = เนื้อความหนังสือก่อนย่อหน้า 6 pt, รายงาน = หัวข้อก่อนย่อหน้า 6 pt")
    words(s)
    s.set_defaults(fn=cmd_fix)

    s = sub.add_parser("check", help="รายงานปัญหาโดยไม่แก้ไฟล์")
    s.add_argument("IN"); s.add_argument("--thai-digits", action="store_true")
    s.set_defaults(fn=cmd_check)

    s = sub.add_parser("wrap", help="ตัดคำไม่ให้ขาดกลางคำ")
    s.add_argument("IN"); s.add_argument("OUT"); words(s)
    s.set_defaults(fn=cmd_wrap)

    s = sub.add_parser("unwrap", help="ถอดการตัดคำ")
    s.add_argument("IN"); s.add_argument("OUT")
    s.set_defaults(fn=cmd_unwrap)

    s = sub.add_parser("segment", help="พิมพ์ผลตัดคำ คำ|คำ")
    s.add_argument("TEXT"); words(s)
    s.set_defaults(fn=cmd_segment)
    return p


def main(argv: list[str] | None = None) -> int:
    _utf8()
    a = _parser().parse_args(argv)
    try:
        return a.fn(a)
    except CliError as e:
        print(f"ข้อผิดพลาด: {e}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
