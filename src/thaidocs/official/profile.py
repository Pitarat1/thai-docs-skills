"""ไฟล์ตั้งค่าหน่วยงาน (agency.json) และการเลือกไฟล์ตรา

รูปแบบไฟล์: {"agency": {...}, "layout": {...}} ทั้งสองส่วนไม่บังคับ
ไม่พึ่ง pydantic/Pillow/fontTools: ตรวจชนิดค่าเทียบกับ field ของ LayoutSpec เอง
"""
from __future__ import annotations

import json
import typing
from dataclasses import dataclass, field, fields, replace
from pathlib import Path

from .style import DEFAULT_SPEC, LayoutSpec

PLACEHOLDER_EMBLEM = Path(__file__).parent / "data" / "emblem_placeholder.png"

# รหัสชนิดหนังสือ (ตรงกับ schema.DocType — เทสตรวจ) ใช้ตรวจ layout.omit_signature โดยไม่ต้องนำเข้า pydantic
DOC_TYPE_CODES = ("memo", "external_letter", "stamp_letter", "order", "regulation", "bylaw", "announcement",
                  "statement", "news", "certificate", "minutes")

_PNG = b"\x89PNG\r\n\x1a\n"
_JPEG = b"\xff\xd8\xff"


class ProfileError(ValueError):
    """ไฟล์ตั้งค่าหรือไฟล์ตราไม่ถูกต้อง (ข้อความเป็นภาษาไทย)"""


@dataclass
class AgencyInfo:
    org: str = ""
    address: list[str] = field(default_factory=list)
    department: str = ""
    phone: str = ""
    fax: str = ""
    email: str = ""
    doc_no_prefix: str = ""
    signer_name: str = ""
    signer_position: list[str] = field(default_factory=list)


_AGENCY_LISTS = {"address", "signer_position"}


def _home_dir() -> Path:
    return Path.home() / ".thaidocs"


def _coerce(name: str, value, tp, where: str):
    """แปลง/ตรวจค่าจาก JSON ให้ตรงชนิดของ field; ผิดชนิด -> ProfileError"""
    origin = typing.get_origin(tp)
    if origin is typing.Union:  # Optional[X]
        args = [a for a in typing.get_args(tp) if a is not type(None)]
        if value is None:
            return None
        return _coerce(name, value, args[0], where)
    if tp is bool:
        if isinstance(value, bool):
            return value
    elif tp is int:
        if isinstance(value, int) and not isinstance(value, bool):
            return value
    elif tp is float:
        if isinstance(value, (int, float)) and not isinstance(value, bool):
            return value
    elif tp is str:
        if isinstance(value, str):
            return value
    elif tp is tuple or origin is tuple:
        if isinstance(value, (list, tuple)):
            targs = typing.get_args(tp)
            if targs and targs[0] is str:
                if all(isinstance(v, str) for v in value):
                    return tuple(value)
            elif all(isinstance(v, (int, float)) and not isinstance(v, bool) for v in value):
                return tuple(value)
    raise ProfileError(f"ค่า {where}.{name} ชนิดไม่ถูกต้อง: {value!r}")


def _check_layout_value(key: str, value) -> None:
    """ค่าที่ชนิดถูกแต่ใช้ไม่ได้: เลขหน้า ขนาด/ระยะติดลบ ขอบเป็นศูนย์ ชนิดหนังสือที่ไม่รู้จัก"""
    if value is None:
        return
    if key == "page_number_fmt" and value.count("{}") != 1:
        raise ProfileError(f"ค่า layout.page_number_fmt ต้องมี {{}} (ตำแหน่งเลขหน้า) หนึ่งที่พอดี เช่น \"- {{}} -\": "
                           f"{value!r}")
    if key.endswith(("_cm", "_pt")) and isinstance(value, (int, float)):
        if "margin" in key and value <= 0:
            raise ProfileError(f"ค่า layout.{key} (ขอบกระดาษ) ต้องมากกว่า 0: {value!r}")
        if value < 0:
            raise ProfileError(f"ค่า layout.{key} ต้องไม่ติดลบ: {value!r}")
    if key == "omit_signature":
        bad = [v for v in value if v not in DOC_TYPE_CODES]
        if bad:
            raise ProfileError(f"ค่า layout.omit_signature มีรหัสชนิดหนังสือที่ไม่รู้จัก: {', '.join(bad)} "
                               f"(รหัสที่รับ: {', '.join(DOC_TYPE_CODES)})")


def _layout_from(data: dict) -> LayoutSpec:
    hints = typing.get_type_hints(LayoutSpec)
    names = {f.name for f in fields(LayoutSpec)}
    updates = {}
    for key, value in data.items():
        if key not in names:
            raise ProfileError(f"ไม่รู้จักค่า layout.{key} ในไฟล์ตั้งค่า")
        updates[key] = _coerce(key, value, hints[key], "layout")
        _check_layout_value(key, updates[key])
    return replace(DEFAULT_SPEC, **updates)


def _agency_from(data: dict) -> AgencyInfo:
    hints = typing.get_type_hints(AgencyInfo)
    names = {f.name for f in fields(AgencyInfo)}
    kwargs = {}
    for key, value in data.items():
        if key not in names:
            raise ProfileError(f"ไม่รู้จักค่า agency.{key} ในไฟล์ตั้งค่า")
        if key in _AGENCY_LISTS:
            if isinstance(value, str):
                value = [value]
            if not (isinstance(value, list) and all(isinstance(v, str) for v in value)):
                raise ProfileError(f"ค่า agency.{key} ต้องเป็นรายการข้อความ: {value!r}")
            kwargs[key] = list(value)
        else:
            kwargs[key] = _coerce(key, value, hints[key], "agency")
    return AgencyInfo(**kwargs)


def load_profile(path: Path | None) -> tuple[AgencyInfo | None, LayoutSpec]:
    """อ่านไฟล์ตั้งค่า; path เป็น None -> ลอง ~/.thaidocs/agency.json ถ้าไม่มี -> (None, DEFAULT_SPEC)"""
    if path is None:
        path = _home_dir() / "agency.json"
        if not path.is_file():
            return None, DEFAULT_SPEC
    path = Path(path)
    try:
        data = json.loads(path.read_text(encoding="utf-8-sig"))
    except OSError as e:
        raise ProfileError(f"อ่านไฟล์ตั้งค่า {path} ไม่ได้: {e.strerror or e}") from e
    except (UnicodeDecodeError, json.JSONDecodeError) as e:
        raise ProfileError(f"ไฟล์ตั้งค่า {path} ไม่ใช่ JSON ที่ถูกต้อง: {e}") from e
    if not isinstance(data, dict):
        raise ProfileError(f"ไฟล์ตั้งค่า {path} ต้องเป็นออบเจ็กต์ JSON")
    for key in data:
        if key not in ("agency", "layout"):
            raise ProfileError(f"ไม่รู้จัก key ระดับบน '{key}' ในไฟล์ตั้งค่า (รับเฉพาะ agency, layout)")
    for key in ("agency", "layout"):
        if key in data and not isinstance(data[key], dict):
            raise ProfileError(f"'{key}' ในไฟล์ตั้งค่าต้องเป็นออบเจ็กต์")
    agency = _agency_from(data["agency"]) if "agency" in data else None
    spec = _layout_from(data["layout"]) if "layout" in data else DEFAULT_SPEC
    return agency, spec


def _check_emblem(path: Path) -> Path:
    if not path.exists():
        raise ProfileError(f"ไม่พบไฟล์ตรา: {path}")
    if not path.is_file():
        raise ProfileError(f"ตราต้องเป็นไฟล์ ไม่ใช่โฟลเดอร์: {path}")
    try:
        with path.open("rb") as f:
            head = f.read(8)
    except OSError as e:
        raise ProfileError(f"อ่านไฟล์ตราไม่ได้: {path}") from e
    if not (head.startswith(_PNG) or head.startswith(_JPEG)):
        raise ProfileError(f"ไฟล์ตราต้องเป็น PNG หรือ JPEG: {path}")
    from docx.image.image import Image          # อ่านหัวภาพแบบเดียวกับตอนใส่ลงเอกสาร (ไม่ต้องใช้ Pillow)
    try:
        Image.from_file(str(path))
    except Exception as e:                      # ไฟล์เสีย/ตัดท้าย: python-docx โยนได้หลายชนิด
        raise ProfileError(f"ไฟล์ตราอ่านเป็นรูปไม่ได้ ({path}): {type(e).__name__} {e}".rstrip()) from e
    return path


def resolve_emblem(cli_path: Path | None, spec: LayoutSpec) -> Path | None:
    """ลำดับ: ตัวเลือกบรรทัดคำสั่ง > layout.emblem_png > ~/.thaidocs/emblem.png > None"""
    if cli_path is not None:
        return _check_emblem(Path(cli_path).expanduser())
    if spec.emblem_png:
        return _check_emblem(Path(spec.emblem_png).expanduser())
    default = _home_dir() / "emblem.png"
    if default.exists():
        return _check_emblem(default)
    return None
