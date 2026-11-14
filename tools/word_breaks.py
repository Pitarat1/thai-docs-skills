"""วัดการขึ้นบรรทัดกลางคำใน Word for Mac จริง (เครื่องมือผู้พัฒนา ไม่อยู่ใน pytest)

ต่อยอดจากการทดลองใน spec: สร้างไฟล์ docx สองแบบจากชุด gold (tests/data/gold_th.txt) × เยื้องขวา 7 ค่า
  base = finalize อย่างเดียว (ย่อหน้า thaiDistribute, TH Sarabun New 16 pt)
  wrap = base + thaidocs.wrap.wrap
ส่งให้ Word แปลงเป็น PDF (สำเนาเก็บที่ dist/word_breaks_<variant>.pdf) ทีละไฟล์ (AppleScript) อ่านบรรทัดด้วย pdfplumber แล้วนับจุดขึ้นบรรทัดที่ไม่ใช่ขอบคำ

    .venv/bin/python tools/word_breaks.py [--write-sample]

เกณฑ์ผ่าน: wrap กลางคำ <= 3% และ base > wrap (ไม่ผ่านคืน exit 1)
ไฟล์ทดสอบในโฟลเดอร์ sandbox ของ Word (ชื่อขึ้นต้น thaidocs_wb_) ถูกลบเสมอเมื่อจบ
"""
from __future__ import annotations

import argparse
import difflib
import os
import subprocess
import sys
import unicodedata
import uuid
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))

RIGHT_CM = [0, 1.3, 2.6, 3.9, 5.2, 6.5, 7.8]
SAMPLE_CM = [0, 2.6, 5.2]
GOLD = ROOT / "tests" / "data" / "gold_th.txt"
SAMPLE = ROOT / "tests" / "data" / "word_check.docx"
DIST = ROOT / "dist"
IMAGE = DIST / "word_breaks.png"
IMAGE_DOC = ROOT / "docs" / "images" / "before-after.png"
WORD_DIR = Path.home() / "Library/Containers/com.microsoft.Word/Data/Documents"
PREFIX = "thaidocs_wb_"
STRIP = {" ", "​", "⁠", "﻿", "\n", "\t", " "}


# ---------- ส่วน Python ล้วน (ทดสอบได้โดยไม่ต้องมี Word) ----------

def read_gold(path: Path = GOLD) -> list[tuple[str, set[int]]]:
    """[(ข้อความ, จุดตัดที่อนุญาต)]: ``|`` = ตัดได้, ช่องว่างที่ p = ตัดได้ที่ p และ p+1"""
    paras = []
    for line in path.read_text(encoding="utf-8").splitlines():
        if not line.strip():
            continue
        allowed: set[int] = set()
        pos = 0
        for ch in line:
            if ch == "|":
                allowed.add(pos)
                continue
            if ch == " ":
                allowed.update((pos, pos + 1))
            pos += 1
        paras.append((line.replace("|", ""), allowed))
    return paras


def norm(s: str) -> str:
    """NFC ตัดช่องว่าง/ตัวล่องหน; PDF ของ Word ถอด สระอำ เป็น นิคหิต+สระอา (ํา) จึงรวมกลับ"""
    s = unicodedata.normalize("NFC", s).replace("ํา", "ำ")
    return "".join(c for c in s if c not in STRIP)


def allowed_stripped(text: str, allowed: set[int]) -> set[int]:
    """จุดตัดที่อนุญาต เป็นตำแหน่งในข้อความที่ตัดช่องว่าง/ตัวอักขระล่องหนออกแล้ว
    (ตำแหน่งที่ข้ามช่องว่างนับว่าตัดได้)"""
    orig = [i for i, c in enumerate(text) if c not in STRIP]
    return {k for k, i in enumerate(orig) if i in allowed or (k and orig[k - 1] != i - 1)}


def count_breaks(lines: list[str], text: str, allowed: set[int]) -> tuple[int, list[int], int]:
    """จากบรรทัดที่อ่านได้ของย่อหน้าเดียว คืน (จำนวนจุดขึ้นบรรทัด, ดัชนีบรรทัดที่ลงท้ายกลางคำ, 1 ถ้าข้อความไม่ตรง)

    ดัชนีคือบรรทัดที่ถูกตัดต่อ (ไม่รวมบรรทัดสุดท้าย) ใช้ difflib จับคู่เพื่อทนตัวอักษรที่ถอดผิดเล็กน้อย
    """
    lines = [norm(l) for l in lines]
    keep = [i for i, l in enumerate(lines) if l]
    ls = [lines[i] for i in keep]
    ref = norm(text)
    got = "".join(ls)
    ok = allowed_stripped(text, allowed)
    sm = difflib.SequenceMatcher(None, got, ref, autojunk=False)
    blocks = sm.get_matching_blocks()

    def to_ref(k: int):
        for a, b, n in blocks:
            if a <= k <= a + n:
                return b + (k - a)
        return None

    total, bad, pos = 0, [], 0
    for n, l in enumerate(ls[:-1]):
        pos += len(l)
        r = to_ref(pos)
        if r is None or r in (0, len(ref)):
            continue
        total += 1
        if r not in ok:
            bad.append(keep[n])
    return total, bad, int(got != ref)


# ---------- สร้างเอกสาร ----------

def build_doc(variant: str, rights: list[float]):
    """variant: base | wrap — ย่อหน้าละหน้า (ขึ้นหน้าใหม่) ตามลำดับ gold × rights"""
    from docx import Document
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Cm, Pt

    from thaidocs.layout import finalize
    from thaidocs.wordcut import Segmenter
    from thaidocs.wrap import wrap

    d = Document()
    sec = d.sections[0]
    sec.page_width, sec.page_height = Cm(21), Cm(29.7)
    sec.left_margin, sec.right_margin, sec.top_margin, sec.bottom_margin = Cm(3), Cm(2), Cm(2), Cm(2)
    first = True
    for text, _ in read_gold():
        for r in rights:
            p = d.add_paragraph(text)
            pf = p.paragraph_format
            pf.right_indent, pf.first_line_indent = Cm(r), Cm(1.25)
            pf.space_after, pf.line_spacing = Pt(0), 1.0
            pf.page_break_before = not first
            first = False
            ppr = p._p.get_or_add_pPr()
            jc = OxmlElement("w:jc")
            jc.set(qn("w:val"), "thaiDistribute")
            ppr.append(jc)
    finalize(d)
    if variant == "wrap":
        wrap(d, Segmenter())
    return d


# ---------- Word (AppleScript) ----------

def word_running() -> bool:
    out = subprocess.run(["osascript", "-e", 'application "Microsoft Word" is running'],
                         capture_output=True, text=True)
    return out.stdout.strip() == "true"


def convert(docx: Path, pdf: Path) -> None:
    """docx -> pdf ด้วย Word ทั้งสองไฟล์ต้องอยู่ใน sandbox ของ Word; ปิดเฉพาะเอกสารที่เปิดเอง"""
    script = f'''
tell application "Microsoft Word"
    open POSIX file "{docx}"
    set d to active document
    save as d file name "{pdf}" file format format PDF
    close d saving no
end tell'''
    r = subprocess.run(["osascript", "-e", script], capture_output=True, text=True, timeout=300)
    if r.returncode != 0 or not pdf.exists():
        raise RuntimeError(f"AppleScript/Word ล้มเหลว (exit {r.returncode}): {r.stderr.strip()}")


def measure(variant: str, rights: list[float], pdf_out: Path):
    """สร้าง -> แปลง -> อ่านบรรทัดทั้งไฟล์; คืนผลรายหน้า [(text_idx, right, total, bad_idx, mism, lines)]"""
    import pdfplumber

    gold = read_gold()
    name = f"{PREFIX}{uuid.uuid4().hex[:8]}_{variant}"
    docx, pdf = WORD_DIR / f"{name}.docx", WORD_DIR / f"{name}.pdf"
    pages = [(i, r) for i in range(len(gold)) for r in rights]
    try:
        build_doc(variant, rights).save(docx)
        convert(docx, pdf)
        pdf_out.parent.mkdir(parents=True, exist_ok=True)
        pdf_out.write_bytes(pdf.read_bytes())  # อ่านจากสำเนาใน dist/ (ไฟล์ใน sandbox ถูกลบเมื่อจบ)
    finally:
        for f in (docx, pdf):
            try:
                f.unlink()
            except FileNotFoundError:
                pass
    results = []
    P = pdfplumber.open(pdf_out)  # เปิดค้างไว้ให้ render_pair ใช้ (ปิดตอนโปรเซสจบ)
    if len(P.pages) != len(pages):
        raise RuntimeError(f"{variant}: ได้ {len(P.pages)} หน้า คาด {len(pages)}")
    for pg, (gi, r) in zip(P.pages, pages):
        texts = [l["text"] for l in pg.extract_text_lines()]
        text, allowed = gold[gi]
        total, bad, mism = count_breaks(texts, text, allowed)
        results.append((gi, r, total, bad, mism, pg.extract_text_lines(), pg))
    return results


# ---------- ภาพเทียบ ----------

def render_pair(base_res, wrap_res, out: Path) -> tuple[int, float]:
    """เลือกหน้าที่ base ผิดมากกว่า wrap มากที่สุด วางเทียบซ้าย (base) ขวา (wrap) ขีดแดงที่ท้ายบรรทัดกลางคำ"""
    from PIL import Image, ImageDraw, ImageFont

    k = max(range(len(base_res)), key=lambda i: (len(base_res[i][3]) - len(wrap_res[i][3]), -i))
    imgs = []
    for res in (base_res[k], wrap_res[k]):
        gi, r, total, bad, mism, lines, pg = res
        scale = 110 / 72
        im = pg.to_image(resolution=110).original.convert("RGB")
        dr = ImageDraw.Draw(im, "RGBA")
        for b in bad:
            ch = lines[b]["chars"][-1]
            dr.rectangle([ch["x0"] * scale - 3, ch["top"] * scale - 2, ch["x1"] * scale + 3, ch["bottom"] * scale + 2],
                         fill=(255, 0, 0, 70), outline=(220, 0, 0, 255))
        ys = [l["top"] for l in lines], [l["bottom"] for l in lines]
        box = (0, max(0, int(min(ys[0]) * scale) - 24), im.width, min(im.height, int(max(ys[1]) * scale) + 24))
        imgs.append((im.crop(box), len(bad), total))
    h = max(i[0].height for i in imgs)
    try:
        font = ImageFont.truetype("/System/Library/Fonts/Supplemental/Tahoma.ttf", 22)
    except OSError:
        font = ImageFont.load_default()
    canvas = Image.new("RGB", (imgs[0][0].width * 2 + 30, h + 44), "white")
    d = ImageDraw.Draw(canvas)
    for n, (im, bad, total) in enumerate(imgs):
        x = n * (im.width + 30)
        canvas.paste(im, (x, 44))
        d.text((x + 10, 8), f"{'base' if n == 0 else 'wrap'}: {bad}/{total} mid-word", fill=(0, 0, 0), font=font)
    out.parent.mkdir(parents=True, exist_ok=True)
    canvas.save(out)
    return k, base_res[k][1]


def tally(res):
    total = sum(r[2] for r in res)
    bad = sum(len(r[3]) for r in res)
    return total, bad, sum(r[4] for r in res)


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("--write-sample", action="store_true", help="เขียน tests/data/word_check.docx (ไม่ต้องใช้ Word) แล้วจบ")
    a = ap.parse_args(argv)
    if a.write_sample:
        build_doc("wrap", SAMPLE_CM).save(SAMPLE)
        print(f"เขียน {SAMPLE}")
        return 0
    if not WORD_DIR.is_dir():
        print("BLOCKED: ไม่พบโฟลเดอร์ sandbox ของ Word (ไม่ได้ติดตั้ง Word for Mac?)", file=sys.stderr)
        return 2
    was_running = word_running()
    out = {}
    try:
        for v in ("base", "wrap"):  # ทีละไฟล์
            out[v] = measure(v, RIGHT_CM, DIST / f"word_breaks_{v}.pdf")
    finally:
        if not was_running:
            subprocess.run(["osascript", "-e", 'tell application "Microsoft Word" to quit saving no'], capture_output=True)
    print(f"{'variant':8} {'breaks':>7} {'mid-word':>9} {'%':>7}  text-mismatch-pages")
    rate = {}
    for v, res in out.items():
        t, b, m = tally(res)
        rate[v] = b / max(t, 1)
        print(f"{v:8} {t:7} {b:9} {rate[v]:7.1%}  {m}")
    k, r = render_pair(out["base"], out["wrap"], IMAGE)
    IMAGE_DOC.parent.mkdir(parents=True, exist_ok=True)
    IMAGE_DOC.write_bytes(IMAGE.read_bytes())
    print(f"ภาพเทียบ: {IMAGE} (ย่อหน้า {out['base'][k][0]} เยื้องขวา {r} ซม.) -> {IMAGE_DOC}")
    ok = rate["wrap"] <= 0.03 and rate["base"] > rate["wrap"]
    print("ผ่าน" if ok else "ไม่ผ่าน (ต้อง wrap <= 3% และ base > wrap)")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
