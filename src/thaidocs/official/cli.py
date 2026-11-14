"""thaigov: สร้างและตรวจหนังสือราชการ

  types                                  รายการชนิดหนังสือ ๑๑ ชนิด
  schema TYPE                            field ของแต่ละชนิด
  example TYPE [OUT.json]                ตัวอย่างร่าง (พิมพ์ทาง stdout เมื่อไม่ระบุ OUT)
  render IN.json OUT.docx [--emblem FILE] [--profile FILE] [--font NAME] [--arabic-digits]
                          [--wrap-words] [--words FILE ...] [--force]
  check IN.json|IN.docx [--profile FILE]

exit code: 0 สำเร็จ, 1 check พบปัญหาระดับ error, 2 ข้อผิดพลาดการใช้งาน/ข้อมูลเข้า (ข้อความไทยทาง stderr)
pydantic นำเข้าเมื่อใช้คำสั่งเท่านั้น (ไม่มี -> แจ้งให้ติดตั้ง)
"""
from __future__ import annotations

import argparse
import importlib
import json
import os
import types as _types
import typing
from dataclasses import replace
from pathlib import Path

from ..cli import CliError, _save, _segmenter, _utf8
from .profile import ProfileError, load_profile, resolve_emblem

_LEVEL_ORDER = {"error": 0, "warning": 1, "info": 2}
_LEVEL_TH = {"error": "ข้อผิดพลาด", "warning": "คำเตือน", "info": "ข้อสังเกต"}


def _schema():
    """นำเข้า schema (ต้องใช้ pydantic) แบบเลื่อนเวลา เพื่อแจ้งภาษาไทยแทน traceback"""
    try:
        return importlib.import_module(f"{__package__}.schema")
    except ImportError as e:
        if (getattr(e, "name", None) or "").split(".")[0] == "pydantic":
            raise CliError("ไม่พบไลบรารี pydantic ติดตั้งด้วยคำสั่ง: pip install pydantic") from None
        raise


def _doc_type(sc, text: str):
    for t in sc.DocType:
        if text == t.value or text == sc.DOC_TYPE_TH[t]:
            return t
    names = "\n".join(f"  {t.value}  {sc.DOC_TYPE_TH[t]}" for t in sc.DocType)
    raise CliError(f"ไม่รู้จักชนิดหนังสือ '{text}' ชนิดที่รับ:\n{names}")


# ------------------------------------------------------------------ types / schema / example
def cmd_types(a) -> int:
    sc = _schema()
    for t in sc.DocType:
        print(f"{t.value:<16} {sc.DOC_TYPE_TH[t]}  ({sc.DOC_TYPE_REF[t]})")
    return 0


def _models_in(ann) -> list:
    """โมเดล pydantic ที่อยู่ในชนิดของ field (ผ่าน Optional / list / Union)"""
    out = []
    if isinstance(ann, type) and hasattr(ann, "model_fields"):
        return [ann]
    for arg in typing.get_args(ann):
        out += _models_in(arg)
    return out


def _type_th(ann) -> str:
    origin = typing.get_origin(ann)
    args = typing.get_args(ann)
    if origin is typing.Literal:
        return "ค่าหนึ่งใน " + " / ".join(str(getattr(v, "value", v)) for v in args)
    if origin is typing.Union or origin is _types.UnionType:
        rest = [x for x in args if x is not type(None)]
        return " หรือ ".join(_type_th(x) for x in rest)
    if origin is list:
        return "รายการของ" + _type_th(args[0]) if args else "รายการ"
    if ann is str:
        return "ข้อความ"
    if ann is int:
        return "จำนวนเต็ม"
    if ann is bool:
        return "จริง/เท็จ"
    if isinstance(ann, type) and hasattr(ann, "model_fields"):
        return f"กลุ่มข้อมูล {ann.__name__}"
    return str(ann)


def _default_th(f) -> str:
    if f.is_required():
        return "จำเป็น"
    d = f.get_default(call_default_factory=True)
    if d is None or d == [] or hasattr(type(d), "model_fields"):
        return "ไม่บังคับ"
    shown = ", ".join(map(str, d)) if isinstance(d, list) else str(getattr(d, "value", d))
    return f"ค่าเริ่มต้น {shown}"


def _first_line(text: str | None) -> str:
    return (text or "").strip().splitlines()[0].strip() if text and text.strip() else ""


def _describe(model, depth: int, out: list[str]) -> None:
    pad = "    " * depth
    for name, f in model.model_fields.items():
        sub = _models_in(f.annotation)
        desc = f.description or (_first_line(sub[0].__doc__) if sub else "")
        out.append(f"{pad}{name} ({_type_th(f.annotation)}, {_default_th(f)}): {desc}".rstrip(": ").rstrip())
        for m in sub:
            _describe(m, depth + 1, out)


def cmd_schema(a) -> int:
    sc = _schema()
    t = _doc_type(sc, a.TYPE)
    model = sc.MODEL_BY_TYPE[t]
    out = [f"{t.value} {sc.DOC_TYPE_TH[t]} ({sc.DOC_TYPE_REF[t]})"]
    if _first_line(model.__doc__):
        out.append(_first_line(model.__doc__))
    out.append("")
    _describe(model, 0, out)
    print("\n".join(out))
    return 0


def cmd_example(a) -> int:
    sc = _schema()
    t = _doc_type(sc, a.TYPE)
    text = json.dumps(sc.load_example(t), ensure_ascii=False, indent=2) + "\n"
    if a.OUT is None:
        print(text, end="")
        return 0
    out = Path(a.OUT)
    if out.is_dir():
        raise CliError(f"ไฟล์ปลายทาง {out} เป็นโฟลเดอร์ ให้ระบุชื่อไฟล์ .json")
    if out.exists():
        raise CliError(f"ไฟล์ {out} มีอยู่แล้ว (ไม่เขียนทับ) ให้ใช้ชื่อใหม่หรือลบไฟล์เดิมก่อน")
    if not out.parent.is_dir():
        raise CliError(f"ไม่พบโฟลเดอร์ปลายทาง {out.parent}")
    try:
        out.write_text(text, encoding="utf-8")
    except OSError as e:
        raise CliError(f"บันทึกไฟล์ {out} ไม่ได้: {e}") from None
    print(f"สร้างตัวอย่าง {t.value} ที่ {out}")
    return 0


# ------------------------------------------------------------------ render / check
def _read_json(path: Path):
    if not path.is_file():
        raise CliError(f"ไม่พบไฟล์ {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8-sig"))
    except OSError as e:
        raise CliError(f"อ่านไฟล์ {path} ไม่ได้: {e.strerror or e}") from None
    except UnicodeDecodeError:
        raise CliError(f"ไฟล์ {path} ไม่ใช่ข้อความ UTF-8") from None
    except json.JSONDecodeError as e:
        raise CliError(f"ไฟล์ {path} ไม่ใช่ JSON ที่ถูกต้อง (บรรทัด {e.lineno} ตำแหน่ง {e.colno}): {e.msg}") from None


def _parse(sc, data, path: Path):
    try:
        return sc.parse_doc(data)
    except sc.SchemaError as e:
        lines = "\n".join(f"  {m}" for m in e.messages)
        raise CliError(f"ร่าง {path} ไม่ตรงโครงข้อมูล:\n{lines}") from None


def _profile(path: str | None):
    try:
        return load_profile(Path(path) if path else None)
    except ProfileError as e:
        raise CliError(str(e)) from None


def _print_findings(findings) -> None:
    for x in sorted(findings, key=lambda x: _LEVEL_ORDER[x.level]):
        extra = "".join([f" [{x.where}]" if x.where else "", f" ({x.ref})" if x.ref else ""])
        print(f"  {_LEVEL_TH[x.level]} {x.code}: {x.message}{extra}")


def _counts(findings) -> str:
    n = {k: sum(x.level == k for x in findings) for k in _LEVEL_TH}
    return f"ข้อผิดพลาด {n['error']} คำเตือน {n['warning']} ข้อสังเกต {n['info']}"


def _write_bytes(data: bytes, out: Path) -> None:
    tmp = out.with_name(out.name + ".tmp")
    try:
        tmp.write_bytes(data)
        os.replace(tmp, out)
    except OSError as e:
        raise CliError(f"บันทึกไฟล์ {out} ไม่ได้: {e}") from None
    finally:
        if tmp.exists():
            tmp.unlink()


def cmd_render(a) -> int:
    sc = _schema()
    src, out = Path(a.IN), Path(a.OUT)
    if a.words and not a.wrap_words:
        raise CliError("--words ใช้ร่วมกับ --wrap-words เท่านั้น")
    if out.suffix.lower() != ".docx":
        raise CliError(f"ไฟล์ปลายทางต้องลงท้ายด้วย .docx: {out}")
    if out.is_dir():
        raise CliError(f"ไฟล์ปลายทาง {out} เป็นโฟลเดอร์ ให้ระบุชื่อไฟล์ .docx")
    if not out.parent.is_dir():
        raise CliError(f"ไม่พบโฟลเดอร์ปลายทาง {out.parent}")
    if out.exists() and not a.force:
        raise CliError(f"ไฟล์ {out} มีอยู่แล้ว ใช้ --force เพื่อเขียนทับ")

    doc = _parse(sc, _read_json(src), src)
    agency, spec = _profile(a.profile)
    if a.font:
        spec = replace(spec, font=a.font)
    if a.arabic_digits:
        spec = replace(spec, thai_numerals=False)
    from .checker import check_docx, check_document
    from .fill import fill
    from .render import DocxRenderer
    doc = fill(doc, agency)
    try:
        emblem = resolve_emblem(Path(a.emblem) if a.emblem else None, spec)
    except ProfileError as e:
        raise CliError(str(e)) from None

    renderer = DocxRenderer(spec, emblem=emblem)
    data = renderer.render(doc)
    if a.wrap_words:
        import io

        import docx
        from .. import wrap as wrapmod
        d = docx.Document(io.BytesIO(data))
        wrapmod.wrap(d, _segmenter(a.words))
        _save(d, out)
    else:
        _write_bytes(data, out)

    findings, seen = [], set()
    for x in check_document(doc, spec) + check_docx(out, spec):
        key = (x.level, x.code, x.message, x.where)
        if key not in seen:
            seen.add(key)
            findings.append(x)
    print(f"สร้างไฟล์ {out} ({_counts(findings)})")
    _print_findings([x for x in findings if x.level != "info"])
    if renderer.used_placeholder_emblem:
        print("ใช้กรอบแทนตรา ต้องใส่ตราจริงก่อนใช้งาน: ระบุไฟล์ตราด้วย --emblem FILE (PNG/JPEG) "
              "หรือวางไว้ที่ ~/.thaidocs/emblem.png")
    return 0


def cmd_check(a) -> int:
    src = Path(a.IN)
    ext = src.suffix.lower()
    if ext not in (".json", ".docx"):
        raise CliError(f"ตรวจได้เฉพาะไฟล์ .json (ร่าง) หรือ .docx: {src}")
    sc = _schema()
    from .checker import check_docx, check_document
    agency, spec = _profile(a.profile)
    if ext == ".json":
        from .fill import fill
        doc = fill(_parse(sc, _read_json(src), src), agency)
        findings = check_document(doc, spec)
    else:
        if not src.is_file():
            raise CliError(f"ไม่พบไฟล์ {src}")
        try:
            findings = check_docx(src, spec)
        except CliError:
            raise
        except Exception as e:
            raise CliError(f"เปิดไฟล์ {src} ไม่ได้ (ไม่ใช่ .docx ที่อ่านได้): {e}") from None
    if not findings:
        print("ไม่พบปัญหา")
        return 0
    print(f"ผลตรวจ {src.name}: {_counts(findings)}")
    _print_findings(findings)
    return 1 if any(x.level == "error" for x in findings) else 0


# ------------------------------------------------------------------ parser / main
def _parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(prog="thaigov", description="สร้างและตรวจหนังสือราชการ ๑๑ ชนิด")
    sub = p.add_subparsers(dest="cmd", required=True)

    s = sub.add_parser("types", help="รายการชนิดหนังสือ")
    s.set_defaults(fn=cmd_types)

    s = sub.add_parser("schema", help="แสดง field ของชนิดหนังสือ")
    s.add_argument("TYPE")
    s.set_defaults(fn=cmd_schema)

    s = sub.add_parser("example", help="ตัวอย่างร่าง JSON ของชนิดหนังสือ")
    s.add_argument("TYPE"); s.add_argument("OUT", nargs="?", default=None)
    s.set_defaults(fn=cmd_example)

    s = sub.add_parser("render", help="สร้างไฟล์ .docx จากร่าง JSON")
    s.add_argument("IN"); s.add_argument("OUT")
    s.add_argument("--emblem", metavar="FILE", help="ไฟล์ตรา PNG/JPEG")
    s.add_argument("--profile", metavar="FILE", help="ไฟล์ตั้งค่าหน่วยงาน (ค่าเริ่มต้น ~/.thaidocs/agency.json)")
    s.add_argument("--font", metavar="NAME", help="ฟอนต์ที่ใช้แทนค่าเริ่มต้น")
    s.add_argument("--arabic-digits", action="store_true", dest="arabic_digits", help="ใช้เลขอารบิก (ไม่แปลงเป็นเลขไทย)")
    s.add_argument("--wrap-words", action="store_true", dest="wrap_words", help="ตัดคำไม่ให้ขาดกลางคำ")
    s.add_argument("--words", action="append", default=[], metavar="FILE",
                   help="ไฟล์คำเพิ่มเติม (บรรทัดละคำ) ใช้ซ้ำได้ ใช้คู่กับ --wrap-words")
    s.add_argument("--force", action="store_true", help="เขียนทับไฟล์ปลายทางที่มีอยู่")
    s.set_defaults(fn=cmd_render)

    s = sub.add_parser("check", help="ตรวจร่าง .json หรือไฟล์ .docx")
    s.add_argument("IN")
    s.add_argument("--profile", metavar="FILE", help="ไฟล์ตั้งค่าหน่วยงาน")
    s.set_defaults(fn=cmd_check)
    return p


def main(argv: list[str] | None = None) -> int:
    import sys
    _utf8()
    a = _parser().parse_args(argv)
    try:
        return a.fn(a)
    except CliError as e:
        print(f"ข้อผิดพลาด: {e}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
