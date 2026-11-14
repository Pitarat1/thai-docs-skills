from datetime import date

from thaidocs.official import salutations
from thaidocs.official.textutil import (
    clean_text, find_placeholders, has_arabic_digits, is_placeholder, merge_list_lines,
    protect_breaks, split_paragraphs, thai_date, to_arabic_digits,
)


def test_protect_breaks_year():
    assert protect_breaks("พ.ศ. ๒๕๖๙") == "พ.ศ. ๒๕๖๙"


def test_protect_breaks_date_has_no_plain_space():
    assert " " not in protect_breaks("๑๕ พฤศจิกายน ๒๕๖๙")


def test_protect_breaks_phone_and_repeat_mark():
    out = protect_breaks("โทร. ๐ ๒๐๐๐ ๐๐๐๐ อื่น ๆ")
    assert "๐ ๒๐๐๐ ๐๐๐๐" in out and "อื่น ๆ" in out


def test_clean_text():
    assert clean_text("ก\x07ข") == "กข"
    assert clean_text("ก\nข\tค") == "ก\nข\tค"


def test_split_paragraphs():
    assert split_paragraphs(["ก\nข", ""]) == ["ก", "ข"]


def test_thai_date():
    assert thai_date(date(2026, 11, 14)) == "๑๔ พฤศจิกายน ๒๕๖๙"
    assert thai_date(date(2026, 11, 14), with_day_prefix=True).startswith("วันที่ ๑๔")


def test_placeholders():
    assert find_placeholders("ก [ระบุชื่อ] ข [ระบุ]") == ["[ระบุชื่อ]", "[ระบุ]"]
    assert is_placeholder("") and is_placeholder("[ระบุวันที่]")
    assert not is_placeholder("ตัวอย่าง")


def test_digits():
    assert has_arabic_digits("ปี 2569") and not has_arabic_digits("ปี ๒๕๖๙")
    assert not has_arabic_digits("saraban1@example.org")
    assert to_arabic_digits("๑๒") == "12"


def test_merge_list_lines():
    out = merge_list_lines(["จึงขอเชิญประชุม รายละเอียดดังนี้", "วันที่: ๑๘ ธันวาคม ๒๕๖๙", "เวลา: ๑๓.๓๐ น. ถึง ๑๖.๓๐ น."])
    assert out == ["จึงขอเชิญประชุม ในวันที่ ๑๘ ธันวาคม ๒๕๖๙ เวลา ๑๓.๓๐ – ๑๖.๓๐ น."]


def test_salutations():
    assert salutations.suggest("นายกรัฐมนตรี") == ("กราบเรียน", "ขอแสดงความนับถืออย่างยิ่ง")
    assert salutations.suggest("อธิบดี") == ("เรียน", "ขอแสดงความนับถือ")
    assert salutations.pair_ok("เรียน", "ขอแสดงความนับถือ")
    assert not salutations.pair_ok("กราบเรียน", "ขอแสดงความนับถือ")


def test_core_modules_import_without_pydantic():
    import subprocess
    import sys

    code = (
        "import sys; sys.modules['pydantic'] = None\n"
        "import thaidocs.official.textutil, thaidocs.official.salutations\n"
        "import thaidocs.official.style, thaidocs.official.profile\n"
    )
    r = subprocess.run([sys.executable, "-c", code], capture_output=True, text=True)
    assert r.returncode == 0, r.stderr


def test_has_arabic_digits_follows_core_keep_rules():
    assert not has_arabic_digits("โรค COVID-19 ระบาด")
    assert not has_arabic_digits("มาตรฐาน ISO 27001")
    assert not has_arabic_digits("เว็บ https://example.org/a1 อีเมล a1@example.org")
    assert has_arabic_digits("ปี 2569")
    assert has_arabic_digits("โรค COVID-19 ระบาดปี 2569")
