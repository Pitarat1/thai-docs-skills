import json
from pathlib import Path

import pytest

from thaidocs.official.schema import (
    DOC_TYPE_REF, DOC_TYPE_TH, DocType, MemoDoc, SchemaError, load_example, parse_doc,
)
from thaidocs.official.textutil import has_arabic_digits

EXAMPLES = Path(__file__).resolve().parents[2] / "src/thaidocs/official/examples"


@pytest.mark.parametrize("t", list(DocType))
def test_example_parses(t):
    doc = parse_doc(load_example(t))
    assert doc.doc_type == t


def test_tables_cover_all_types():
    assert set(DOC_TYPE_TH) == set(DocType) == set(DOC_TYPE_REF)
    assert DOC_TYPE_REF[DocType.MEMO] == "ข้อ ๑๒"


def test_missing_doc_type():
    with pytest.raises(SchemaError) as e:
        parse_doc({"subject": "x"})
    assert "doc_type" in e.value.messages[0]


def test_unknown_doc_type_lists_choices():
    with pytest.raises(SchemaError) as e:
        parse_doc({"doc_type": "xyz"})
    msg = e.value.messages[0]
    assert all(t.value in msg for t in DocType)


def test_memo_needs_body_or_sections():
    data = load_example(DocType.MEMO)
    del data["body"]
    with pytest.raises(SchemaError) as e:
        parse_doc(data)
    joined = "\n".join(e.value.messages)
    assert "body" in joined and "sections" in joined


def test_bad_speed_points_at_field():
    data = load_example(DocType.MEMO)
    data["speed"] = "ด่วนจี๋"
    with pytest.raises(SchemaError) as e:
        parse_doc(data)
    assert e.value.messages[0].startswith("speed:")
    assert "ด่วนที่สุด" in e.value.messages[0]


def test_extra_and_missing_and_type_errors():
    data = load_example(DocType.MEMO)
    data["bogus"] = 1
    del data["subject"]
    data["to"] = ["ก"]
    with pytest.raises(SchemaError) as e:
        parse_doc(data)
    msgs = e.value.messages
    assert "bogus: ไม่รู้จัก field นี้" in msgs
    assert "subject: ต้องระบุ" in msgs
    assert "to: ต้องเป็นข้อความ" in msgs


def test_nested_location_and_non_dict():
    data = load_example(DocType.ORDER)
    data["committees"][0]["members"][0].pop("role")
    with pytest.raises(SchemaError) as e:
        parse_doc(data)
    assert e.value.messages == ["committees.0.members.0.role: ต้องระบุ"]
    with pytest.raises(SchemaError):
        parse_doc([])


def test_memo_model_is_memo():
    assert isinstance(parse_doc(load_example(DocType.MEMO)), MemoDoc)


def _date_values(o):
    if isinstance(o, dict):
        for k, v in o.items():
            if k in ("date", "effective") and isinstance(v, str):
                yield v
            yield from _date_values(v)
    elif isinstance(o, list):
        for x in o:
            yield from _date_values(x)


@pytest.mark.parametrize("path", sorted(EXAMPLES.glob("*.json")), ids=lambda p: p.stem)
def test_examples_have_no_arabic_digits_in_dates(path):
    for v in _date_values(json.loads(path.read_text(encoding="utf-8"))):
        assert not has_arabic_digits(v), v


def test_example_files_match_doc_types():
    assert {p.stem for p in EXAMPLES.glob("*.json")} == {t.value for t in DocType}
