"""พจนานุกรมรวมสำหรับตัดคำ และตัวตัดคำ newmm (คัดจาก PyThaiNLP ใน _vendor)"""

from __future__ import annotations

from collections.abc import Iterable
from pathlib import Path

from ._vendor.newmm import segment as _segment
from ._vendor.trie import Trie

DATA_DIR = Path(__file__).parent / "data"
DATA_FILES = ("words_th.txt", "names_th.txt", "places_th.txt", "domain_th.txt")

# id(words) -> (words, trie) เก็บ words ไว้เพื่อไม่ให้ id ถูกนำกลับไปใช้ซ้ำ
_TRIE_CACHE: dict[int, tuple[frozenset[str], Trie]] = {}


def _read_words(path: Path) -> set[str]:
    out: set[str] = set()
    with open(path, encoding="utf-8-sig") as f:
        for line in f:
            w = line.strip()
            if w and not w.startswith("#"):
                out.add(w)
    return out


def load_words(extra: Iterable[str | Path] = ()) -> frozenset[str]:
    """รวมไฟล์ใน data/ + ไฟล์ของผู้ใช้ + ~/.thaidocs/words.txt (ถ้ามี)"""
    words: set[str] = set()
    for name in DATA_FILES:
        words |= _read_words(DATA_DIR / name)
    for p in extra:
        words |= _read_words(Path(p))
    try:
        home_file = Path.home() / ".thaidocs" / "words.txt"
        if home_file.is_file():
            words |= _read_words(home_file)
    except (OSError, RuntimeError):
        pass
    return frozenset(words)


def _trie_for(words: frozenset[str]) -> Trie:
    hit = _TRIE_CACHE.get(id(words))
    if hit is not None and hit[0] is words:
        return hit[1]
    trie = Trie(words)
    _TRIE_CACHE[id(words)] = (words, trie)
    return trie


def newmm_words(text: str, words: frozenset[str]) -> list[str]:
    """ตัดคำด้วย newmm คง whitespace เป็น token; ``"".join(result) == text``"""
    if not text:
        return []
    return _segment(text, _trie_for(words))
