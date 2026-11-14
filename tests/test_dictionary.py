from pathlib import Path

import pytest

from thaidocs import dictionary
from thaidocs.dictionary import load_words, newmm_words

DATA = Path(dictionary.__file__).parent / "data"


@pytest.fixture(autouse=True)
def fake_home(tmp_path, monkeypatch):
    """ไม่ให้เทสพึ่งไฟล์ ~/.thaidocs/words.txt จริง"""
    home = tmp_path / "home"
    home.mkdir()
    monkeypatch.setenv("HOME", str(home))
    monkeypatch.setenv("USERPROFILE", str(home))
    monkeypatch.setattr(Path, "home", classmethod(lambda cls: home))
    return home


def test_newmm_keeps_text():
    t = "ด้วยสำนักงานสาธารณสุขจังหวัดสระแก้วมีความประสงค์"
    toks = newmm_words(t, load_words())
    assert "".join(toks) == t
    assert "สาธารณสุข" in toks and "ความประสงค์" in toks


def test_newmm_keeps_whitespace_tokens():
    t = "ด้วย  สำนักงาน\tabc 123\nสระแก้ว"
    toks = newmm_words(t, load_words())
    assert "".join(toks) == t
    assert "  " in toks and "\t" in toks and "\n" in toks


def test_newmm_empty():
    assert newmm_words("", load_words()) == []


def test_places_kept_whole():
    if not (DATA / "places_th.txt").read_text(encoding="utf-8").strip():
        pytest.skip("places_th.txt ว่าง")
    assert "วังน้ำเย็น" in newmm_words("อำเภอวังน้ำเย็น", load_words())


def test_data_files_sorted_unique():
    for name in ("words_th.txt", "names_th.txt", "places_th.txt", "domain_th.txt"):
        lines = (DATA / name).read_text(encoding="utf-8").splitlines()
        assert lines, name
        assert all(x and x == x.strip() for x in lines), name
    for name in ("names_th.txt", "places_th.txt"):
        lines = (DATA / name).read_text(encoding="utf-8").splitlines()
        assert lines == sorted(set(lines)), name


def test_domain_words_present():
    w = load_words()
    for x in ("ชำนาญการพิเศษ", "หมอพร้อม", "แรนซัมแวร์", "เมทฟอร์มิน"):
        assert x in w


def test_user_words(tmp_path):
    p = tmp_path / "w.txt"
    p.write_text("# หมายเหตุ\n\nสสจ.สระแก้ว\n", encoding="utf-8")
    w = load_words([p])
    assert "สสจ.สระแก้ว" in w
    assert "# หมายเหตุ" not in w and "" not in w
    assert isinstance(w, frozenset)


def test_home_words_file(fake_home):
    (fake_home / ".thaidocs").mkdir()
    (fake_home / ".thaidocs" / "words.txt").write_text("คำของฉันเอง\n", encoding="utf-8")
    assert "คำของฉันเอง" in load_words()


def test_user_word_changes_segmentation(tmp_path):
    t = "สสจ.สระแก้ว"
    base = load_words()
    p = tmp_path / "w.txt"
    p.write_text("สสจ.สระแก้ว\n", encoding="utf-8")
    toks = newmm_words(t, load_words([p]))
    assert toks == [t]
    assert newmm_words(t, base) != [t]


def test_trie_cache_by_id():
    w = load_words()
    newmm_words("สระแก้ว", w)
    newmm_words("สระแก้ว", w)
    assert len(dictionary._TRIE_CACHE) >= 1
