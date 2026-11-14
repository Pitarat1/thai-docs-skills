"""ความสะอาดของไฟล์ที่ git ติดตาม: ไม่มี path ในเครื่อง อีเมลส่วนตัว หรือไฟล์ ` 2` ที่ iCloud สร้าง"""
from __future__ import annotations

import fnmatch
import re
import subprocess
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]

_PATH_PATTERNS = (
    re.compile(r"/Users/[A-Za-z]"),
    re.compile(r"/home/[a-z]"),
    re.compile(r"[A-Z]:\\Users\\"),
)
_EMAIL = re.compile(r"[\w.+-]+@[\w-]+(?:\.[\w-]+)*\.[A-Za-z]{2,}")
_ALLOWED_EMAIL_DOMAINS = ("example.com", "users.noreply.github.com", "anthropic.com")
_ICLOUD_DUP = re.compile(r" 2(\.[^/]*)?$")


def _tracked() -> list[str]:
    try:
        r = subprocess.run(["git", "ls-files", "-z"], cwd=ROOT, capture_output=True, check=True)
    except (OSError, subprocess.CalledProcessError):
        pytest.skip("ไม่มี git หรือไม่ใช่ repo git")
    return [n for n in r.stdout.decode("utf-8").split("\0") if n]


def _skipped(name: str) -> bool:
    return fnmatch.fnmatch(name, "*/data/*.txt") or "_khammoon/model/" in name


def _texts():
    for name in _tracked():
        if _skipped(name) or not (ROOT / name).is_file():
            continue
        try:
            yield name, (ROOT / name).read_bytes().decode("utf-8")
        except UnicodeDecodeError:
            continue


def test_no_local_paths():
    bad = [f"{n}: {m.group(0)}" for n, t in _texts() for p in _PATH_PATTERNS for m in p.finditer(t)]
    assert not bad, bad


def test_emails_are_placeholders():
    bad = []
    for n, t in _texts():
        for m in _EMAIL.finditer(t):
            if not m.group(0).lower().endswith(tuple("@" + d for d in _ALLOWED_EMAIL_DOMAINS)):
                bad.append(f"{n}: {m.group(0)}")
    assert not bad, bad


def test_no_icloud_duplicates():
    bad = [n for n in _tracked() if _ICLOUD_DUP.search(n)]
    assert not bad, bad
