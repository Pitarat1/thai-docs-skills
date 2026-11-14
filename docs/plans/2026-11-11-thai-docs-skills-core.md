# ชุด skill เอกสารไทย — โครงการย่อยที่ 1 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** สร้าง repo `Pitarat1/thai-docs-skills` ที่มี skill `thai-docx` (จัดหน้า Word ไทย + ตัดคำไม่ให้ขาดกลางคำ) ใช้ได้กับ Claude Code, claude.ai, Codex, Antigravity พร้อมหน้าแนะนำ (Artifact)

**Architecture:** โค้ด Python ต้นฉบับอยู่ที่ `src/thaidocs/` (layout, digits, check, wordcut, wrap, cli) `tools/build.py` คัดลอกเข้า `plugins/thai-docs/skills/thai-docx/scripts/` เพื่อให้โฟลเดอร์ skill ใช้ได้ด้วยตัวเองเมื่อ zip ตัวตัดคำ = newmm (คัดจาก PyThaiNLP) ∩ khammoon (backend numpy, ถอดได้) บนพจนานุกรมรวม wrap แยกคำละ run ไม่ติดป้ายไทย + ZWSP

**Tech Stack:** Python ≥3.10, python-docx, lxml, numpy (เสริม), pytest, GitHub Actions

**Spec:** `docs/specs/2026-11-11-thai-docs-skills-core-design.md`

**ต้นทางที่ต้องเปิดอ่านขณะทำ:**
- โค้ด SARABUN.AI `render/thai_layout.py` (151 บรรทัด, ใช้จริง; อยู่ในโครงการภายในของผู้พัฒนา)
- skill จัดหน้าเดิม `thai-docx-layout-th` (ภายใน) `SKILL.md` (skill เดิม หลัก 10 ข้อ + โค้ด lxml)
- โค้ดทดลองตัดคำ (ภายใน) (`hyb.py`, `gen.py`, `gen2.py`, `measure.py`, `gold.txt`) — ถ้าโฟลเดอร์ชั่วคราวหายแล้ว ข้อมูลที่จำเป็นอยู่ใน spec และในแผนนี้ครบ
- PyThaiNLP 5.4 `pythainlp/tokenize/newmm.py`, `tcc_p.py`, `pythainlp/util/trie.py`; khammoon 0.1.0 `numpy_model.py`, `segmenter.py`, `tcc.py`, `chars.py`, `model/` (ติดตั้งใน venv ชั่วคราวเพื่อคัดลอก: `pip download` แล้วแตกไฟล์ในโฟลเดอร์ใหม่ที่ว่าง)

## Global Constraints

- ที่อยู่ repo: โฟลเดอร์ `thai-docs-skills` บนเครื่องผู้พัฒนา · remote: `github.com/Pitarat1/thai-docs-skills`
- **ห้าม push / สร้าง repo บน GitHub / ทำ release / เผยแพร่ artifact เป็นสาธารณะ จนกว่าผู้ใช้อนุมัติ** (Task 12)
- git identity แบบ repo-local: `user.name Pitarat1`, `user.email <id>+Pitarat1@users.noreply.github.com` (`id` จาก `gh api users/Pitarat1 --jq .id`) ห้ามใช้อีเมลส่วนตัว
- ฟอนต์เริ่มต้น `TH Sarabun New`, ขนาดเนื้อความ 16 pt, ตัวเลขอารบิกเป็นค่าเริ่มต้น
- สคริปต์ไม่เรียก AI/เครือข่ายใดๆ
- dependency บังคับ: `python-docx>=1.1`, `lxml`; เสริม: `numpy` (khammoon) ห้ามพึ่ง PyThaiNLP ตอนรัน
- เนื้อหา SKILL.md/references ไม่อ้างเครื่องมือเฉพาะของ AI ตัวใด (เขียนเป็น "รันคำสั่ง …")
- ข้อความถึงผู้ใช้ (CLI, README, SKILL.md) เป็นภาษาไทย
- OUT ต้องไม่ใช่ไฟล์เดียวกับ IN; ปฏิเสธถ้ามีไฟล์ล็อก `~$<ชื่อ>` ข้าง IN (ไฟล์เปิดใน Word)
- iCloud สร้างไฟล์ `* 2.*` ในโฟลเดอร์ Documents: ตรวจ `git status` ก่อน commit ทุกครั้ง ห้าม commit ไฟล์ที่ชื่อลงท้าย ` 2`
- ทุกไฟล์ที่คัดจากโครงการอื่นต้องมีหัวไฟล์ระบุที่มา + สัญญาอนุญาต และอยู่ใน `NOTICE`

## Review Focus

1. **`fix` กับไฟล์ที่ wrap แล้ว** — `thai_runs` จะคืน `cs` ให้ run ที่ wrap ไว้ ทำให้ wrap พังครึ่งๆ คาดว่า: `fix` ตรวจ `is_wrapped` → unwrap ก่อน แล้ว wrap คืนตอนท้าย (Task 7 `test_fix_on_wrapped_doc_rewraps`)
2. **ย่อหน้าที่มี hyperlink / field / ข้อความอังกฤษ-URL ปน** — wrap ต้องไม่แตะ `instrText`, ข้อความใน hyperlink ยังคลิกได้, ZWSP เฉพาะระหว่างไทย-ไทย (Task 6 `test_wrap_keeps_fields_and_links`, `test_wrap_mixed_script`)
3. **ข้อความในตาราง หัว/ท้ายกระดาษ เชิงอรรถ** — ทุกคำสั่งต้องประมวลผลทุก part ไม่ใช่แค่ body (Task 3 `test_finalize_all_parts`, Task 6 `test_wrap_tables_headers`)
4. **wrap ซ้ำ / unwrap แล้ว wrap** — ข้อความไม่เปลี่ยน ไม่มี ZWSP ซ้อน (Task 6 `test_wrap_idempotent`, `test_roundtrip_text_equal`)
5. **ไม่มี numpy / ถอด khammoon** — ยังทำงานด้วย newmm อย่างเดียวและเตือนครั้งเดียว (Task 5 `test_without_khammoon_fallback`)

---

### Task 1: โครง repo + แพ็กเกจเปล่า + CI ทดสอบ

**Files:**
- Create: `pyproject.toml` (ชื่อแพ็กเกจ `thaidocs`, src layout, extras `khammoon=["numpy"]`, `dev=["pytest","numpy"]`), `src/thaidocs/__init__.py` (`__version__ = "0.1.0"`), `tests/test_smoke.py`, `.gitignore` (`dist/`, `.venv/`, `__pycache__/`, `* 2.*`), `.github/workflows/test.yml` (Python 3.10 และ 3.13, `pip install -e .[dev]`, `pytest -q`)

**Interfaces:** Produces: แพ็กเกจ `thaidocs` import ได้

- [ ] **Step 1:** (ทำแล้ว 11 พ.ย. 2569: โฟลเดอร์, `git init -b main`, identity repo-local, spec และแผนใน `docs/`) ตรวจด้วย `git config user.email`
- [ ] **Step 2:** เขียน `tests/test_smoke.py::test_version` → `assert thaidocs.__version__ == "0.1.0"`
- [ ] **Step 3:** สร้างไฟล์ทั้งหมดข้างบน, `python3 -m venv .venv && .venv/bin/pip install -e .[dev]`
- [ ] **Step 4:** `.venv/bin/pytest -q` → `1 passed`
- [ ] **Step 5:** commit `chore: scaffold thai-docs-skills`

### Task 2: พจนานุกรม + newmm (ไม่มี khammoon)

**Files:**
- Create: `src/thaidocs/_vendor/__init__.py`, `src/thaidocs/_vendor/tcc.py`, `src/thaidocs/_vendor/trie.py`, `src/thaidocs/_vendor/newmm.py` (คัดจาก PyThaiNLP 5.4, ตัด import ภายนอกออก), `src/thaidocs/data/words_th.txt`, `src/thaidocs/data/names_th.txt` (ชื่อชาย+หญิง+นามสกุล+จังหวัด จาก PyThaiNLP รวมเป็นไฟล์เดียว เรียง ไม่ซ้ำ), `src/thaidocs/data/places_th.txt`, `src/thaidocs/data/domain_th.txt`, `src/thaidocs/dictionary.py`, `NOTICE`, `tests/test_dictionary.py`

**Interfaces:**
- Produces: `dictionary.load_words(extra: Iterable[str | Path] = ()) -> frozenset[str]` (รวมไฟล์ใน `data/` + ไฟล์ผู้ใช้ + `~/.thaidocs/words.txt` ถ้ามี; ข้ามบรรทัดว่าง/ขึ้นต้น `#`) · `dictionary.newmm_words(text: str, words: frozenset[str]) -> list[str]` (คง whitespace เป็น token; cache trie ตาม `id(words)`)

- [ ] **Step 1: ตรวจสัญญาอนุญาตข้อมูล** — อ่าน `pythainlp/corpus/default_db.json` / README ของ PyThaiNLP หาสัญญาอนุญาตของ `words_th.txt`, ชื่อคน, จังหวัด; หาข้อมูลชื่ออำเภอ/ตำบลจาก data.go.th (กรมการปกครอง) ที่สัญญาอนุญาตเปิด ถ้าใช้ได้ → `places_th.txt` = ชื่ออำเภอ+ตำบลไม่ซ้ำ (ตัดคำนำหน้า "อำเภอ/เขต/ตำบล/แขวง"); ถ้าหาแหล่งที่อนุญาตไม่ได้ → `places_th.txt` เหลือเฉพาะชื่ออำเภอจากแหล่งที่อนุญาต หรือไฟล์ว่างพร้อมหมายเหตุใน README บันทึกผลใน `NOTICE`
- [ ] **Step 2:** `domain_th.txt` เริ่มต้น (ขยายได้ภายหลัง): `ชำนาญการ ชำนาญการพิเศษ ชำนาญงาน เชี่ยวชาญ ปฏิบัติการ หมอพร้อม แพลตฟอร์ม แอปพลิเคชัน ไซเบอร์ แรนซัมแวร์ ฟิชชิง เมทฟอร์มิน แอมโลดิพีน` บรรทัดละคำ
- [ ] **Step 3: เขียนเทส**
  - `test_newmm_keeps_text`: `"".join(newmm_words(t, W)) == t` สำหรับข้อความ gold ทุกย่อหน้า (Task 4 สร้างไฟล์ gold; ที่นี่ใช้ประโยค `"ด้วยสำนักงานสาธารณสุขจังหวัดสระแก้วมีความประสงค์"`)
  - `test_places_kept_whole`: `"วังน้ำเย็น" in newmm_words("อำเภอวังน้ำเย็น", load_words())` (ข้ามด้วย `pytest.skip` ถ้า `places_th.txt` ว่าง)
  - `test_user_words`: ไฟล์ชั่วคราวมีคำ `สสจ.สระแก้ว` → `load_words([path])` มีคำนั้น
- [ ] **Step 4:** คัดโค้ด + เขียน `dictionary.py`, รันเทส → ผ่าน
- [ ] **Step 5:** commit `feat(wordcut): vendored newmm and dictionaries`

### Task 3: layout — จัดหน้า Word ไทย

**Files:**
- Create: `src/thaidocs/ooxml.py` (helper: `RPR_ORDER`, `PPR_ORDER`, `reorder`, `child`, `iter_parts(doc)` คืน root XML ของ document, header*, footer*, footnotes, endnotes, styles, numbering ที่มีอยู่), `src/thaidocs/layout.py`, `tests/conftest.py` (fixture `make_doc(paragraphs: list[str], *, table=False, header=None, footnote=False)` สร้างไฟล์ด้วย python-docx), `tests/test_layout.py`

**Interfaces:**
- Produces (ทุกฟังก์ชันรับ `docx.Document`, แก้ในที่, คืนจำนวนที่แก้):
  `normalize_chars(doc) -> int` (U+F70A–F71A → อักษรไทยมาตรฐานตามตารางใน skill เดิม, `ํา`→`ำ`) · `set_font_all(doc, font: str) -> None` (rFonts ascii/hAnsi/cs/eastAsia ทุก run + docDefaults + numbering ยกเว้นฟอนต์สัญลักษณ์ `Symbol Wingdings "Wingdings 2" "Wingdings 3" "Courier New"`, ลบ attribute theme) · `align_sz(doc) -> int` (sz=szCs, b↔bCs) · `remove_manual_breaks(doc) -> int` (ลบ `<w:br/>` ที่ไม่มี type ยกเว้นตัวที่ตามด้วยข้อความขึ้นต้น `- `) · `distribute(doc) -> int` (jc=thaiDistribute ให้ย่อหน้าที่ jc เป็น both/left/ไม่มี และไม่ใช่ heading/TOC/caption/ย่อหน้าที่มี URL) · `keep_next_headings(doc) -> int` · `thai_runs(doc) -> int` (run ที่มีอักษร U+0E00–U+0E7F: เพิ่ม `w:cs`, `w:lang val=en-US bidi=th-TH`; ไฟล์ที่ wrap แล้วจะถูก unwrap ก่อนโดย CLI ใน Task 7 จึงไม่ต้องรู้จัก wrap ที่นี่) · `hide_spelling(doc) -> None` · `doc_defaults(doc, font: str, body_pt: float) -> None` · `finalize(doc, font: str = "TH Sarabun New", body_pt: float = 16.0) -> None` (= set_font_all, align_sz, keep_next_headings, doc_defaults, thai_runs, hide_spelling ตามลำดับ)

- [ ] **Step 1: เขียนเทส**
  - `test_thai_runs_marked`: หลัง `finalize` ทุก `w:r` ที่มีไทยมี `w:cs` และ `w:lang/@w:bidi == "th-TH"`
  - `test_sz_equals_szcs`: run ที่ตั้ง sz=28 ไว้ → หลัง `align_sz` มี szCs=28; run ตัวหนามีทั้ง `b` และ `bCs`
  - `test_font_all_no_theme`: หลัง `set_font_all(doc,"TH Sarabun New")` ไม่มี attribute `asciiTheme/hAnsiTheme/cstheme` ใน styles.xml และทุก rFonts/@ascii == "TH Sarabun New"
  - `test_rpr_schema_order`: ลำดับลูกของทุก `rPr` เป็นลำดับย่อยของ `RPR_ORDER`
  - `test_finalize_all_parts`: เอกสารที่มีตาราง หัวกระดาษ และเชิงอรรถ → run ไทยในทุก part มี `w:cs`
  - `test_normalize_chars`: `"นํา"` → `"นำ"`
  - `test_distribute_skips_headings_and_urls`: Heading 1 และย่อหน้า `ดูที่ https://x.org` ไม่ได้ thaiDistribute; ย่อหน้าธรรมดาได้
  - `test_hide_spelling`: settings.xml มี `hideSpellingErrors` และ `hideGrammaticalErrors` และลำดับ element ถูกตาม schema (อยู่หลังกลุ่มต้นไฟล์)
- [ ] **Step 2:** รัน → FAIL
- [ ] **Step 3:** เขียน `ooxml.py`, `layout.py` โดยรวมตรรกะจาก `thai_layout.py` และ skill เดิมหมวด 9 (แปลงจาก lxml ล้วนเป็นทำงานบน `doc.part` และ part ที่เกี่ยวข้องของ python-docx) หัวไฟล์ระบุที่มา SARABUN.AI (Apache-2.0)
- [ ] **Step 4:** รัน → ผ่าน
- [ ] **Step 5:** commit `feat(layout): Thai Word layout finalize`

### Task 4: เลขไทย + check

**Files:** Create `src/thaidocs/digits.py`, `src/thaidocs/check.py`, `tests/test_digits.py`, `tests/test_check.py`

**Interfaces:**
- Produces: `digits.to_thai_digits(text: str) -> str` (คง URL/อีเมล/DOI และตัวเลขที่ติดอักษรละติน เช่น `HC3`, `ISO 27001` ไม่แปลง) · `digits.apply_thai_digits(doc) -> int` (แปลงเฉพาะ `w:t` ต่อย่อหน้า ความยาวเท่าเดิม, `numFmt decimal→thaiNumbers`, `pgNumType fmt=thaiNumbers` ทุก sectPr) · `check.Issue(code: str, message: str, count: int)` (dataclass) · `check.check(doc, *, thai_digits: bool = False) -> list[Issue]` รหัส: `NO_CS`, `SZ_MISMATCH`, `THEME_FONT`, `MULTI_FONT`, `MANUAL_BREAK`, `BAD_CHARS`, `ARABIC_DIGITS` (เฉพาะ thai_digits=True), `WRAPPED` (แจ้งข้อมูล ไม่นับเป็นปัญหา)

- [ ] **Step 1: เทส** `test_to_thai_digits`: `"ปี 2569 ดู https://a.b/12 HC3"` → `"ปี ๒๕๖๙ ดู https://a.b/12 HC3"` · `test_instrtext_untouched`: field `TOC \o "1-3"` คงเดิมหลัง `apply_thai_digits` · `test_split_runs`: `"25"` และ `"69"` อยู่คนละ run → `"๒๕"`,`"๖๙"` · `test_check_clean_after_finalize`: เอกสารหลัง `finalize` → `check` ไม่มี issue ใดนอกจาก `WRAPPED` · `test_check_finds_no_cs`: เอกสารดิบจาก python-docx ที่มีไทย → มี `NO_CS`
- [ ] **Step 2–4:** FAIL → เขียน → PASS
- [ ] **Step 5:** commit `feat: Thai digits and document check`

### Task 5: khammoon (เสริม) + Segmenter + ชุดทดสอบ gold

**Files:**
- Create: `src/thaidocs/_khammoon/` (คัด `__init__.py` แบบย่อ, `numpy_model.py`, `segmenter.py`, `tcc.py`, `chars.py`, `model/weights.npz`, `model/vocab.json`, `model/meta.json`, `model/LICENSE` — ไม่เอา `model.onnx`, `torch_infer.py`, `cli.py`), `src/thaidocs/wordcut.py`, `tests/data/gold_th.txt`, `tests/test_wordcut.py`

**Interfaces:**
- Consumes: `dictionary.load_words`, `dictionary.newmm_words`
- Produces: `wordcut.khammoon_available() -> bool` · `class wordcut.Segmenter(words: frozenset[str] | None = None, use_khammoon: bool | None = None)` (None = ใช้ถ้ามี; ถ้าขอ True แต่ไม่มี → `RuntimeError` ข้อความไทย; ถ้า None และไม่มี → `warnings.warn` ครั้งเดียว `"ไม่พบ khammoon/numpy: ตัดคำด้วยพจนานุกรมอย่างเดียว คำทับศัพท์อาจขาดกลางคำ"`) · `Segmenter.boundaries(text: str) -> set[int]` = (จุดตัด newmm ∩ จุดตัด khammoon threshold 0.9) ∪ ขอบช่องว่าง ∪ จุดเปลี่ยนไทย/ไม่ใช่ไทย · `Segmenter.words(text: str) -> list[str]`

- [ ] **Step 1: gold** — คัด 4 ย่อหน้าจาก `gold.txt` ของการทดลอง (รูปแบบ: `|` = จุดตัดที่อนุญาต, ช่องว่าง = ตัดได้) แล้วเขียนเพิ่ม 6 ย่อหน้าใหม่ (บันทึกข้อความขออนุมัติ, คำสั่งแต่งตั้ง, บทคัดย่องานวิจัยพยาบาล, รายงานผลตัวชี้วัด, ข่าวประชาสัมพันธ์, คู่มือใช้งานระบบ) ให้มีชื่อคน ชื่อยา ศัพท์ไอที อย่างน้อยย่อหน้าละ 3 คำ
- [ ] **Step 2: เทส**
  - `test_text_preserved`: `"".join(seg.words(t)) == t` ทุกย่อหน้า
  - `test_gold_with_khammoon`: (skip ถ้าไม่มี numpy) อัตราจุดตัดที่ไม่อยู่ใน gold ≤ 0.03
  - `test_gold_without_khammoon`: `Segmenter(use_khammoon=False)` อัตรา ≤ 0.06
  - `test_without_khammoon_fallback`: monkeypatch `khammoon_available` → False; `Segmenter()` ทำงาน และเตือน 1 ครั้งเมื่อสร้าง 2 ตัว
  - `test_mixed_script_boundaries`: `"ระบบHIS 2.0ของรพ."` มีจุดตัดที่ 4 และ 7 และไม่มีจุดตัดภายใน `HIS`
- [ ] **Step 3:** คัด khammoon (หัวไฟล์ Apache-2.0 / โมเดล CC BY-NC-SA 4.0), เขียน `wordcut.py`
- [ ] **Step 4:** PASS (ถ้า gold ใหม่เกินเกณฑ์: เพิ่มคำใน `domain_th.txt` เท่านั้น ห้ามแก้ gold ให้ผ่าน; ถ้ายังเกิน รายงานผู้ใช้พร้อมตัวอย่าง)
- [ ] **Step 5:** commit `feat(wordcut): dictionary ∩ khammoon segmenter`

### Task 6: wrap / unwrap

**Files:** Create `src/thaidocs/wrap.py`, `tests/test_wrap.py`

**Interfaces:**
- Consumes: `wordcut.Segmenter`, `ooxml.iter_parts`, `ooxml.reorder`
- Produces: `wrap.is_wrapped(doc) -> bool` (custom property `thaidocs.wrapped == "1"`) · `wrap.wrap(doc, seg: Segmenter) -> int` (จำนวนย่อหน้าที่แก้) · `wrap.unwrap(doc) -> int` · ค่าคงที่ `ZWSP = "​"`

ขั้นวิธี wrap ต่อย่อหน้า (ตาม spec): ลบ ZWSP เดิม → ต่อข้อความ run ปกติทั้งย่อหน้า (`w:r` ที่เป็นลูกตรงของ `w:p` และใน `w:hyperlink`; ไม่รวม run ที่มี `fldChar`/`instrText`, `w:sym`, `w:tab`, `w:br` นับเป็นตัวแบ่ง) → `seg.boundaries` → แยก run ที่จุดตัด (deepcopy `rPr`) → run ที่มีไทย: ลบ `w:cs`, `w:lang val=en-US bidi=en-US`, rFonts 4 ช่องเท่ากัน, sz=szCs → ต่อท้าย ZWSP ให้คำไทยที่คำถัดไปขึ้นต้นด้วยไทย (ในย่อหน้าเดียวกัน แม้คนละ run) → ตั้ง custom property · unwrap: ลบ ZWSP, คืน `w:cs` + bidi th-TH, รวม run ติดกันที่ `rPr` เหมือนกัน (เทียบ XML ที่ canonicalize), ลบ property

- [ ] **Step 1: เทส**
  - `test_wrap_marks_runs`: หลัง wrap ไม่มี run ไทยที่มี `w:cs`; ข้อความไม่รวม ZWSP == เดิม; มี ZWSP ระหว่าง `อำเภอ` กับ `วังน้ำเย็น`
  - `test_wrap_idempotent`: wrap 2 ครั้ง → XML ของ body เท่ากับ wrap ครั้งเดียว
  - `test_roundtrip_text_equal`: unwrap(wrap(d)) → ข้อความทุกย่อหน้าเท่าต้นฉบับ, ทุก run ไทยมี `w:cs`, `is_wrapped` False
  - `test_wrap_keeps_formatting`: ย่อหน้ามีคำตัวหนาและเลขยกกำลัง → หลัง wrap run ที่มีข้อความนั้นยังมี `w:b` / `vertAlign=superscript`
  - `test_wrap_keeps_fields_and_links`: field PAGE และ hyperlink ภาษาไทย → `instrText` ไม่เปลี่ยน, hyperlink ยังมี `r:id` เดิมและข้อความเดิม (ไม่นับ ZWSP)
  - `test_wrap_mixed_script`: `"ระบบ HIS ของโรงพยาบาล"` → ไม่มี ZWSP ติดกับ `HIS`
  - `test_wrap_tables_headers`: ข้อความไทยในตารางและหัวกระดาษถูก wrap
- [ ] **Step 2–4:** FAIL → เขียน → PASS
- [ ] **Step 5:** commit `feat(wrap): keep Thai words whole in Word line breaks`

### Task 7: CLI `thaidocx`

**Files:** Create `src/thaidocs/cli.py`, `src/thaidocs/__main__.py`, `tests/test_cli.py`

**Interfaces:**
- Consumes: ทุกโมดูลข้างบน
- Produces: `cli.main(argv: list[str] | None = None) -> int` คำสั่ง `fix IN OUT [--font] [--body-pt] [--thai-digits] [--distribute] [--remove-manual-breaks] [--wrap-words] [--words FILE ...]`, `check IN [--thai-digits]` (exit 1 ถ้ามี issue ที่ไม่ใช่ WRAPPED; พิมพ์บรรทัดละ `รหัส จำนวน ข้อความ`), `wrap IN OUT [--words]`, `unwrap IN OUT`, `segment TEXT [--words]` (พิมพ์ `คำ|คำ`); ข้อผิดพลาด → ข้อความไทยทาง stderr exit 2

ลำดับ `fix`: ถ้า `is_wrapped` → unwrap ก่อนและจำไว้ → normalize_chars → set_font_all → align_sz → (remove_manual_breaks) → (apply_thai_digits) → (distribute) → keep_next_headings → doc_defaults → thai_runs → hide_spelling → wrap ถ้า `--wrap-words` หรือเคย wrap

- [ ] **Step 1: เทส** `test_fix_roundtrip` (fix แล้ว `check` exit 0) · `test_fix_on_wrapped_doc_rewraps` (wrap → fix ไม่ใส่ `--wrap-words` → ยัง `is_wrapped` และไม่มี run ไทยที่มี cs) · `test_refuse_same_path` (exit 2) · `test_refuse_word_lock` (มี `~$x.docx` ข้าง `x.docx` → exit 2 ข้อความมีคำว่า "ปิดไฟล์") · `test_segment_output` (`segment "อำเภอวังน้ำเย็น"` พิมพ์ `อำเภอ|วังน้ำเย็น`) · `test_missing_dependency_message` (monkeypatch import docx ล้ม → ข้อความบอก `pip install python-docx lxml`)
- [ ] **Step 2–4:** FAIL → เขียน → PASS
- [ ] **Step 5:** commit `feat(cli): thaidocx commands`

### Task 8: ตรวจกับ Word for Mac (เครื่องมือผู้พัฒนา)

**Files:** Create `tools/word_breaks.py` (ย้ายตรรกะจาก `gen.py`/`gen2.py`/`measure.py` ของการทดลอง: สร้างไฟล์จาก `tests/data/gold_th.txt` × เยื้องขวา `[0,1.3,2.6,3.9,5.2,6.5,7.8]` ซม. แบบ `base` และ `wrap` (ผ่าน `thaidocs.wrap`) → copy ไป `~/Library/Containers/com.microsoft.Word/Data/Documents/` → AppleScript `save as … file format format PDF` (รันทีละไฟล์ เบื้องหลัง) → pdfplumber นับการตัดกลางคำ → ลบไฟล์ทดสอบออกจากโฟลเดอร์ Word), `tools/requirements-dev.txt` (`pdfplumber`), `tests/data/word_check.docx` (ไฟล์ wrap แล้วสำหรับส่งให้คนเปิดใน Windows/LibreOffice)

- [ ] **Step 1:** รัน `.venv/bin/python tools/word_breaks.py` → พิมพ์ตาราง base/wrap; เกณฑ์ผ่าน: wrap ≤ 3% และ base > wrap
- [ ] **Step 2:** ดูภาพเทียบที่สคริปต์บันทึก (`dist/word_breaks.png`) ว่ากระจายแบบไทยเต็มขอบสองด้าน
- [ ] **Step 3:** commit `test: Word for Mac line-break check`

### Task 9: SKILL.md + references + build + plugin manifests

**Files:**
- Create: `plugins/thai-docs/.claude-plugin/plugin.json` (`name: thai-docs`, `version: 0.1.0`, description ไทย), `.claude-plugin/marketplace.json` (`name: thai-docs-skills`, owner `Pitarat1`, plugin `thai-docs` source `./plugins/thai-docs`), `plugins/thai-docs/skills/thai-docx/SKILL.md`, `…/references/layout-rules.md`, `…/references/troubleshooting.md`, `…/references/wrap-words.md`, `…/references/preview.md`, `…/scripts/thaidocx.py` (ตัวเรียก: เพิ่มโฟลเดอร์ตัวเองใน `sys.path` แล้ว `from thaidocs.cli import main`), `tools/build.py`, `tests/test_build.py`
- `scripts/thaidocs/` ถูกสร้างโดย build เท่านั้น (ใส่ใน `.gitignore`? **ไม่** — commit ผลลัพธ์ไว้เพื่อให้ติดตั้งผ่าน plugin/คัดลอกโฟลเดอร์ได้ทันที; CI ตรวจว่าตรงกับ `src/`)

**Interfaces:**
- Produces: `tools/build.py [--without-khammoon] [--zip]` → sync `src/thaidocs` → `plugins/thai-docs/skills/thai-docx/scripts/thaidocs` (ลบ `__pycache__`); `--zip` → `dist/thai-docx.zip` (โฟลเดอร์ `thai-docx/` ที่ราก zip) และ `dist/thai-docx-nokhammoon.zip`; `--check` → exit 1 ถ้า scripts/ ไม่ตรงกับ src/

SKILL.md frontmatter: `name: thai-docx`, description (ไทย ≤ 1024 ตัวอักษร) ระบุ: สร้าง/แก้ไฟล์ Word ภาษาไทย ฟอนต์ TH Sarabun New เลขไทย กระจายแบบไทย ตัดคำไม่ให้ขาดกลางคำ ตรวจก่อนส่ง · เนื้อหา ≤ 150 บรรทัด: เมื่อไรใช้, ติดตั้ง dependency, ขั้นตอน (เอกสารใหม่: python-docx → `layout.finalize` / ไฟล์เดิม: `check` → `fix` → ดูผล → `wrap` เป็นขั้นสุดท้าย), ตารางคำสั่ง, ตรวจก่อนส่ง, ชี้ไป references · references ย้ายเนื้อหาจาก skill เดิม (หมวด 1, 3, 4, 5, 6, 7, 8) ปรับให้อ้างคำสั่ง CLI แทนโค้ดที่ต้องพิมพ์เอง และ `preview.md` ใส่ 3 ทาง: Word for Mac (AppleScript), LibreOffice `soffice --headless --convert-to pdf`, ฟอนต์แทน (สูตร Sarabun→TH SarabunPSK ของ skill เดิม ปรับชื่อเป็น TH Sarabun New)

- [ ] **Step 1: เทส** `test_build_check_in_sync` (รัน build แล้ว `--check` exit 0) · `test_zip_layout` (zip มี `thai-docx/SKILL.md`, `thai-docx/scripts/thaidocx.py`, ไม่มี `__pycache__`; ขนาด < 10 MB; ฉบับ nokhammoon ไม่มี `weights.npz`) · `test_skill_frontmatter` (มี `name`, `description` ≤ 1024 ตัวอักษร, SKILL.md ≤ 150 บรรทัด) · `test_skill_script_runs` (`python plugins/…/scripts/thaidocx.py segment "ทดสอบ"` exit 0 ใน subprocess ที่ไม่มี `src` ใน path)
- [ ] **Step 2–4:** FAIL → เขียน → PASS; เพิ่ม `python tools/build.py --check` ใน `test.yml`
- [ ] **Step 5:** commit `feat: thai-docx skill, plugin manifests, build`

### Task 10: README, LICENSE, NOTICE, release workflow, ตรวจแพลตฟอร์ม

**Files:** Create `README.md`, `LICENSE` (Apache-2.0), แก้ `NOTICE`, `.github/workflows/release.yml` (on tag `v*`: pytest → `build.py --zip` → แนบ 2 zip ใน GitHub Release)

- [ ] **Step 1: ตรวจข้อเท็จจริงจากเอกสารทางการ** (WebFetch/WebSearch) แล้วบันทึกในหัวข้อติดตั้ง: ตำแหน่งโฟลเดอร์ skill ของ Codex และ Antigravity, วิธีอัปโหลด skill ใน claude.ai และขนาด zip สูงสุด, คำสั่ง plugin marketplace ของ Claude Code ข้อใดหาไม่เจอ เขียนว่า "ยังไม่ได้ยืนยัน"
- [ ] **Step 2:** README ภาษาไทย: ทำอะไรได้ (ภาพก่อน/หลังจาก Task 8), ติดตั้ง 4 แบบ, คำสั่งที่ใช้บ่อย, ข้อแลกเปลี่ยนของ wrap (4 ข้อตาม spec), ตาราง "ทดสอบกับ" (Word for Mac ✓ พร้อมตัวเลข, Windows/LibreOffice "รอตรวจ"), สัญญาอนุญาตแยกส่วน (ตาราง spec) + วิธีใช้ฉบับไม่มี khammoon
- [ ] **Step 3: ลองใช้จริง** — ติดตั้ง plugin ใน Claude Code จากโฟลเดอร์ local (`/plugin marketplace add <path ของ repo บนเครื่อง>`) แล้วสั่งงานตัวอย่าง "จัดไฟล์ tests/data/word_check.docx ให้ตัดคำไม่ขาด" ใน session ใหม่ ดูว่า skill ถูกเรียกและผลผ่าน `check`; ทำซ้ำกับ Codex หรือ Antigravity ถ้ามีในเครื่อง บันทึกผลใน README
- [ ] **Step 4:** commit `docs: README, licenses, release workflow`

### Task 11: Artifact หน้าแนะนำ

**Files:** Create `site/index.html` (หน้าเดียว ไม่มีไฟล์ภายนอกนอกจาก Google Fonts Sarabun), `site/before-after.png` (จาก Task 8)

- [ ] **Step 1:** โหลด skill `artifact-design` (ตามกติกาของเครื่องมือ Artifact) แล้วเขียนหน้า: หัว "ชุด skill เอกสารไทย", ทำอะไรได้, ภาพก่อน/หลัง + ตัวเลข 12.8% → 2.0%, แท็บติดตั้ง 4 แบบ (เนื้อหาเดียวกับ README), คำสั่งที่ใช้บ่อย, ข้อแลกเปลี่ยน, สัญญาอนุญาต, ลิงก์ `https://github.com/Pitarat1/thai-docs-skills` (ทำงานหลัง Task 12)
- [ ] **Step 2:** ดูหน้าในเบราว์เซอร์ที่ความกว้างมือถือและเดสก์ท็อป โหมดสว่าง/มืด
- [ ] **Step 3:** เผยแพร่เป็น Artifact (ส่วนตัว) ส่งลิงก์ให้ผู้ใช้
- [ ] **Step 4:** commit `docs: landing page`

### Task 12: ขึ้น GitHub (ต้องได้รับอนุมัติ)

- [ ] **Step 1:** ขอผู้ใช้: (ก) ล็อกอิน `gh auth login` เป็น Pitarat1 เอง (ห้ามขอรหัส/โทเค็นในแชท) (ข) ยืนยันชื่อ repo, สาธารณะหรือส่วนตัว (ค) อนุมัติ push และสร้าง release `v0.1.0`
- [ ] **Step 2:** หลังอนุมัติ: `gh repo create Pitarat1/thai-docs-skills --<public|private> --source . --push` (ใช้บัญชี Pitarat1 ครั้งเดียว ไม่สลับ active account ถาวร) ตรวจว่า Actions `test` เขียว
- [ ] **Step 3:** `git tag v0.1.0 && git push origin v0.1.0` → ตรวจว่า Release มี 2 zip
- [ ] **Step 4:** ทดสอบติดตั้งจาก GitHub จริง: `/plugin marketplace add Pitarat1/thai-docs-skills` และดาวน์โหลด zip จาก Release
- [ ] **Step 5:** ส่งไฟล์ `tests/data/word_check.docx` ให้ผู้ใช้ส่งต่อคนในทีมเปิดด้วย Word for Windows / LibreOffice แล้วอัปเดตตาราง "ทดสอบกับ" ใน README
