# ชุด skill เอกสารไทย — โครงการย่อยที่ 2 หนังสือราชการ Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** เพิ่ม skill `thai-official-letters` ที่สร้างและตรวจหนังสือราชการ 11 ชนิดเป็น .docx จากร่าง JSON ตามระเบียบงานสารบรรณและ ว 2019

**Architecture:** แพ็กเกจย่อย `src/thaidocs/official/` (schema → fill → render → checker → cli) นำมาจากแกน SARABUN.AI ของระบบต้นทาง ตัดส่วน AI และของเฉพาะหน่วยงาน ขั้นสุดท้ายของตัวสร้างใช้ `thaidocs.layout.finalize` และ `thaidocs.digits.to_thai_digits` ของแกนโครงการ 1 `tools/build.py` คัดลอกแพ็กเกจเข้าแต่ละ skill ตามรายการ โดย `thai-docx` ไม่ได้ `official/`

**Tech Stack:** Python ≥3.10, python-docx, lxml, pydantic ≥2 (เฉพาะ skill ใหม่), numpy (เสริม), pytest

**Spec:** `docs/specs/2026-11-14-thai-official-letters-design.md` (ผู้ลงมือทำอ่านคู่กับแผนนี้)

**ต้นทางที่ต้องเปิดอ่านขณะทำ (อ่านอย่างเดียว ห้ามแก้):** โฟลเดอร์ `engine/sarabun` ของโครงการภายในของผู้พัฒนา (ตำแหน่งอยู่ในบันทึกในเครื่อง ไม่เขียนลง repo) ไฟล์ที่ใช้: `schema.py`, `style.py`, `render/docx_renderer.py`, `checker.py`, `salutations.py`, `thai_utils.py`, `prose.py`, ส่วน `_postprocess`/`forward_paragraphs`/`parse_doc` ใน `service.py`, และ `tests/fixtures/*.json` (ดูโครงเท่านั้น ห้ามคัดข้อมูล)

## Global Constraints

- ทำงานบน branch `feat/official-letters` (แตกจาก `public-main`) **ห้าม push / release / เผยแพร่ artifact จนกว่าผู้ใช้อนุมัติแต่ละครั้ง**
- แก้โค้ดที่ `src/thaidocs/` เท่านั้น แล้วรัน `.venv/bin/python tools/build.py` และ `--check` ก่อน commit ทุกครั้งที่แตะ `src/`
- ห้ามแก้ไฟล์ในโครงการต้นทาง
- สคริปต์ไม่เรียก AI/เครือข่าย
- ฟอนต์เริ่มต้น `TH Sarabun New` ขนาดเนื้อความ 16 pt เลขไทยเป็นค่าเริ่มต้นของหนังสือราชการ
- dependency บังคับของ skill ใหม่: `python-docx>=1.1`, `lxml`, `pydantic>=2` ห้ามบังคับ Pillow / fontTools
- skill `thai-docx` ต้องมีเนื้อหาและพฤติกรรมเท่า v0.1.0 (ไม่มี `official/` ในสำเนา ไม่ต้องใช้ pydantic)
- ห้ามมีชื่อหน่วยงานจริง ชื่อบุคคลจริง เบอร์โทรจริง ข้อมูลภายใน หรือ path ในเครื่อง ในไฟล์ที่ git ติดตาม ตัวอย่างและเทสใช้ "กรมตัวอย่าง" "นายสมมติ ตัวอย่าง" โทร "๐ ๒๐๐๐ ๐๐๐๐" ชื่อ กปภ. มีได้เฉพาะใน `NOTICE`/`README.md` ส่วนที่มา
- ข้อความถึงผู้ใช้ (CLI, SKILL.md, README) เป็นภาษาไทย ไม่อ้างเครื่องมือเฉพาะของ AI ตัวใด
- ข้อผิดพลาดของ CLI: ข้อความไทยทาง stderr exit code 2 ไม่มี traceback
- iCloud สร้างไฟล์ `* 2.*`: ตรวจ `git status` ก่อน commit ห้าม commit ไฟล์ที่ชื่อลงท้าย ` 2`
- git identity แบบ repo-local ที่ตั้งไว้แล้ว ห้ามใช้อีเมลส่วนตัว
- เครื่อง RAM 16 GB: อย่ารันงานหนักซ้อนกัน แปลง docx→PDF ด้วย Word ทีละไฟล์

## Review Focus

1. **ไฟล์ร่างเสีย:** ไม่ใช่ JSON, มี BOM, ไม่มี `doc_type`, `doc_type` ไม่รู้จัก, field ผิดชนิด → ข้อความไทยชี้ field, exit 2 (Task 7)
2. **เนื้อหายาวมาก:** คำสั่งแต่งตั้ง 60 คน เรื่องยาวเกินหนึ่งบรรทัด ข้อความ 3 หน้า → สร้างได้ มีเลขหน้าตั้งแต่หน้าที่สอง (Task 5)
3. **ส่วนที่เว้นว่าง:** `references=[]`, `attachments=[]`, `copies_to=[]`, `fax=None`, ชื่อผู้ลงนามเป็น `""` → ไม่มีป้ายลอยหรือวงเล็บว่าง (Task 4)
4. **ไฟล์ตราใช้ไม่ได้:** path ไม่มีอยู่, ไม่ใช่รูป, เป็นโฟลเดอร์ → ข้อความไทย exit 2 (Task 3, 7)
5. **ข้อความแปลก:** มี `\n` กลางย่อหน้า อักขระควบคุม เลขไทยปนอารบิก แท็บ → ไม่ล้ม `\n` แยกเป็นย่อหน้าใหม่ อักขระควบคุมถูกตัด (Task 2, 6)

---

### Task 1: build หลาย skill และโครงแพ็กเกจ `official`

**Files:**
- Modify: `tools/build.py`, `pyproject.toml`, `tests/test_build.py`
- Create: `src/thaidocs/official/__init__.py` (docstring อย่างเดียว), `plugins/thai-docs/skills/thai-official-letters/scripts/thaigov.py`

**Interfaces:**
- Produces: ใน `tools/build.py` ค่าคงที่ `SKILLS: dict[str, SkillSpec]` โดย `SkillSpec(name: str, exclude: tuple[str, ...])`; `SKILLS["thai-docx"].exclude == ("official",)`, `SKILLS["thai-official-letters"].exclude == ()`; ฟังก์ชันเดิม `sync`, `diff`, `write_zip` รับ `skill: SkillSpec`; `--zip` สร้าง `<name>.zip` และ `<name>-nokhammoon.zip` ต่อ skill โฟลเดอร์รากใน zip = ชื่อ skill

- [ ] **Step 1:** เพิ่มเทสใน `tests/test_build.py`: `test_thai_docx_has_no_official` (หลัง sync ไม่มี `scripts/thaidocs/official/` ใน thai-docx), `test_official_skill_has_full_package` (มี `official/__init__.py`, `layout.py`, `data/words_th.txt`), `test_check_covers_all_skills` (แก้ไฟล์ในสำเนา skill ใหม่แล้ว `--check` คืน 1), `test_zip_per_skill` (ได้ 4 ไฟล์ zip ชื่อตามข้างบน) ปรับเทสเดิมให้เรียกผ่าน `SKILLS["thai-docx"]`
- [ ] **Step 2:** รัน `.venv/bin/python -m pytest tests/test_build.py -q` คาดว่าเทสใหม่ FAIL
- [ ] **Step 3:** แก้ `tools/build.py` ตาม Interfaces; `thaigov.py` เป็นจุดเข้าแบบเดียวกับ `thaidocx.py` แต่ import `thaidocs.official.cli.main` (ยังไม่มีจนถึง Task 7 ให้ `official/__init__.py` มีอย่างเดียวก่อน และ `thaigov.py` เขียนไว้เลย); `pyproject.toml`: รุ่น `0.2.0`, extra `official = ["pydantic>=2"]`, `dev` เพิ่ม `pydantic>=2`, script `thaigov = "thaidocs.official.cli:main"`, package-data เพิ่ม `official/examples/*.json`, `official/data/*.png`
- [ ] **Step 4:** `.venv/bin/pip install -e ".[dev]"` แล้ว `.venv/bin/python tools/build.py && .venv/bin/python tools/build.py --check` คาดว่า "ตรงกับ src/"; `git diff --stat plugins/thai-docs/skills/thai-docx` ต้องว่าง
- [ ] **Step 5:** รันเทสทั้งชุด `.venv/bin/python -m pytest -q` ผ่านทั้งหมด แล้ว commit `build: รองรับหลาย skill และโครงแพ็กเกจ official`

### Task 2: โครงข้อมูล ตัวช่วยข้อความ คำขึ้นต้น และตัวอย่าง

**Files:**
- Create: `src/thaidocs/official/schema.py`, `textutil.py`, `salutations.py`, `examples/<doc_type>.json` × 11
- Test: `tests/official/test_schema.py`, `tests/official/test_textutil.py`, `tests/official/__init__.py`

**Interfaces:**
- Produces (`schema`): `DocType` (ค่า `memo, external_letter, stamp_letter, order, regulation, bylaw, announcement, statement, news, certificate, minutes`), `DOC_TYPE_TH`, `DOC_TYPE_REF: dict[DocType, str]` (ข้อของระเบียบ เช่น "ข้อ ๑๒"), โมเดล `MemoDoc, ExternalLetterDoc, StampLetterDoc, OrderDoc, RegulationDoc, AnnouncementDoc, StatementDoc, CertificateDoc, MinutesDoc` และโมเดลย่อยตามต้นทาง, `AnyDoc`, `MODEL_BY_TYPE`, `parse_doc(data: dict) -> AnyDoc` (โยน `SchemaError(ValueError)` ที่มี `.messages: list[str]` ภาษาไทยรูป `"<ตำแหน่ง field>: <ปัญหา>"`), `load_example(doc_type: DocType) -> dict`
- Produces (`textutil`): `PLACEHOLDER_RE`, `find_placeholders(text) -> list[str]`, `is_placeholder(text) -> bool` (ว่างหรือขึ้นต้น `[ระบุ`), `THAI_MONTHS`, `thai_date(d: date | None = None, *, with_day_prefix: bool = False) -> str`, `has_arabic_digits`, `to_arabic_digits`, `protect_breaks(text) -> str`, `clean_text(text) -> str` (ตัดอักขระควบคุมยกเว้น `\n` `\t`), `split_paragraphs(items: list[str]) -> list[str]` (แยกที่ `\n` ตัดรายการว่าง), `LIST_LINE_RE`, `merge_list_lines(paragraphs) -> list[str]`
- Produces (`salutations`): `suggest(recipient) -> tuple[str, str]`, `pair_ok(salutation, closing) -> bool`

- [ ] **Step 1:** เขียนเทส: ตัวอย่างทั้ง 11 ชนิด `parse_doc(load_example(t))` ผ่านและ `doc.doc_type == t`; ไม่มี `doc_type` → `SchemaError` ที่ `.messages[0]` มีคำว่า `doc_type`; `doc_type="xyz"` → ข้อความมีรายการชนิดที่รับ; `MemoDoc` ไม่มีทั้ง `body` และ `sections` → ข้อความมี "body" และ "sections"; `speed="ด่วนจี๋"` → ข้อความชี้ `speed`; ไฟล์ตัวอย่างทุกไฟล์ไม่มีเลขอารบิกใน field วันที่; `protect_breaks("พ.ศ. ๒๕๖๙")` มี ` ` แทนช่องว่าง; `protect_breaks("๑๕ พฤศจิกายน ๒๕๖๙")` ไม่มีช่องว่างธรรมดา; `clean_text("ก\x07ข")=="กข"`; `split_paragraphs(["ก\nข", ""])==["ก","ข"]`; `thai_date(date(2026,11,14))=="๑๔ พฤศจิกายน ๒๕๖๙"`; `suggest("นายกรัฐมนตรี")==("กราบเรียน","ขอแสดงความนับถืออย่างยิ่ง")`
- [ ] **Step 2:** รัน `.venv/bin/python -m pytest tests/official -q` คาดว่า FAIL (ไม่มีโมดูล)
- [ ] **Step 3:** นำ `schema.py` มาจากต้นทาง: คงชื่อ field และค่าเริ่มต้น ลบคำอ้างหน่วยงานในคำอธิบาย (เปลี่ยน "แบบ ๒/๒ ของ …" เป็น "บันทึกเสนอเรื่องแบบมีหัวข้อ", "แนวปฏิบัติ …" ออก) โมเดลทุกตัวตั้ง `extra="forbid"`; `parse_doc` แปลง `pydantic.ValidationError` เป็นข้อความไทย (missing → "ต้องระบุ", extra → "ไม่รู้จัก field นี้", literal/enum → "ต้องเป็นหนึ่งใน …", ชนิดผิด → "ต้องเป็นข้อความ/รายการ/…")
- [ ] **Step 4:** นำ `textutil.py` (จาก `thai_utils.py` + ส่วนของ `prose.py` ที่ระบุ) และ `salutations.py` มา ไม่นำ `to_thai_digits` ของต้นทาง (ใช้ของแกน)
- [ ] **Step 5:** เขียนตัวอย่าง 11 ไฟล์ขึ้นใหม่ด้วยหน่วยงานและบุคคลสมมติตาม Global Constraints แต่ละไฟล์ใช้ field ครบทุกส่วนที่ชนิดนั้นมี (หนังสือภายนอกมี อ้างถึง สิ่งที่ส่งมาด้วย สำเนาส่ง; คำสั่งมี `committees` พร้อม `subgroups`; ระเบียบมี `chapter`; รายงานการประชุมมี 3 วาระ)
- [ ] **Step 6:** รันเทส PASS; build + `--check`; commit `feat(official): โครงข้อมูล 11 ชนิด ตัวช่วยข้อความ และตัวอย่าง`

### Task 3: ค่าการพิมพ์ ไฟล์ตั้งค่าหน่วยงาน และตรา

**Files:**
- Create: `src/thaidocs/official/style.py`, `profile.py`, `data/emblem_placeholder.png`
- Test: `tests/official/test_profile.py`

**Interfaces:**
- Produces (`style`): `LayoutSpec` (dataclass frozen ตามต้นทาง ไม่มี field `profile`, `emblem_name` คงไว้), `DEFAULT_SPEC` (`font="TH Sarabun New"`), `NATIONAL_FONTS`, `DOCUMENTED_LAYOUT_KEYS = ("font", "body_pt", "margin_left_cm", "margin_right_cm", "margin_top_cm", "margin_bottom_cm", "thai_numerals", "unit_word", "omit_signature", "page_number_fmt", "emblem_png", "stamp_notify_only", "highlight_placeholders")`
- Produces (`profile`): `AgencyInfo` dataclass (`org: str = ""`, `address: list[str]`, `department: str = ""`, `phone: str = ""`, `fax: str = ""`, `email: str = ""`, `doc_no_prefix: str = ""`, `signer_name: str = ""`, `signer_position: list[str]`); `ProfileError(ValueError)`; `load_profile(path: Path | None) -> tuple[AgencyInfo | None, LayoutSpec]` (path เป็น None → ลอง `~/.thaidocs/agency.json` ไม่มี → `(None, DEFAULT_SPEC)`); `resolve_emblem(cli_path: Path | None, spec: LayoutSpec) -> Path | None`; `PLACEHOLDER_EMBLEM: Path`

- [ ] **Step 1:** เขียนเทส (ใช้ `monkeypatch.setenv("HOME", tmp_path)`): ไม่มีไฟล์ → `(None, DEFAULT_SPEC)`; ไฟล์มี `{"agency": {"org": "กรมตัวอย่าง"}, "layout": {"unit_word": "หน่วยงาน", "margin_left_cm": 2.5}}` → ได้ค่าตามนั้น ค่าอื่นเท่า `DEFAULT_SPEC`; key ระดับบนที่ไม่ใช่ `agency`/`layout` → `ProfileError` ข้อความมีชื่อ key; key ใน `layout` ที่ไม่มีใน `LayoutSpec` → `ProfileError`; ค่าชนิดผิด (`"margin_left_cm": "สาม"`) → `ProfileError`; JSON เสีย → `ProfileError`; `omit_signature` รับรายการแล้วเก็บเป็น tuple; ลำดับตรา: `cli_path` > `spec.emblem_png` > `~/.thaidocs/emblem.png` > `None`; `resolve_emblem` กับ path ที่ไม่มีอยู่ เป็นโฟลเดอร์ หรือไฟล์ที่ขึ้นต้นไม่ใช่ลายเซ็น PNG/JPEG → `ProfileError`; `PLACEHOLDER_EMBLEM.is_file()` และขึ้นต้นด้วยลายเซ็น PNG
- [ ] **Step 2:** รันเทส คาดว่า FAIL
- [ ] **Step 3:** นำ `style.py` มา (คงคำอธิบายที่มาของแต่ละค่า ลบคำอ้างหน่วยงาน) เขียน `profile.py`; สร้าง `emblem_placeholder.png` ครั้งเดียวด้วยสคริปต์ชั่วคราวในโฟลเดอร์ scratch (กรอบสี่เหลี่ยมสีเทา 300×330 มีเส้นทแยง พื้นโปร่งใส) แล้ว commit ไฟล์ผลลัพธ์ ไม่ commit สคริปต์
- [ ] **Step 4:** รันเทส PASS; build + `--check`; commit `feat(official): ค่าการพิมพ์ ไฟล์ตั้งค่าหน่วยงาน และตรา`

### Task 4: ตัวสร้าง — แกนและหนังสือโต้ตอบ (บันทึกข้อความ ภายนอก ประทับตรา)

**Files:**
- Create: `src/thaidocs/official/render.py`
- Test: `tests/official/test_render_letters.py`, `tests/official/conftest.py`

**Interfaces:**
- Consumes: `schema` (Task 2), `LayoutSpec`, `PLACEHOLDER_EMBLEM` (Task 3), `textutil.protect_breaks/PLACEHOLDER_RE/clean_text/split_paragraphs`, `thaidocs.layout.finalize(doc, font, body_pt)`, `thaidocs.digits.to_thai_digits`
- Produces: `DocxRenderer(spec: LayoutSpec = DEFAULT_SPEC, emblem: Path | None = None)`; `.render(doc: AnyDoc) -> bytes`; `.render_to(doc, path) -> Path`; `.used_placeholder_emblem: bool` (จริงเมื่อเอกสารที่เพิ่งสร้างใช้กรอบแทนตรา); ใน `tests/official/conftest.py` fixture `render_example(doc_type: str, **overrides) -> docx.Document` (โหลดตัวอย่าง แก้ field ตาม overrides สร้าง แล้วเปิดกลับด้วย python-docx) และตัวช่วย `para(doc, startswith: str)` คืนย่อหน้าแรกที่ข้อความขึ้นต้นตามนั้น

- [ ] **Step 1:** เขียนเทส (ค่าจาก `LayoutSpec` เทียบด้วยความคลาดเคลื่อน ≤ 0.02 ซม. / 0.1 pt):
  - ทั้ง 3 ชนิด: หน้า A4 21.0×29.7 ซม. ขอบซ้าย 3.0 ขวา 2.0 บน 1.5 ล่าง 2.0; ทุก run ฟอนต์ `TH Sarabun New`; `thaidocs.check.check(doc)` คืนรายการว่าง; หัวกระดาษมี field `PAGE` และตั้ง different-first-page
  - บันทึกข้อความ: ย่อหน้า "บันทึกข้อความ" ตัวหนา 29 pt ระยะบรรทัด exact 35 pt; ป้าย "ส่วนราชการ" "ที่" "วันที่" "เรื่อง" ตัวหนา 20 pt; มีรูป 1 รูปสูง 1.5 ซม.
  - หนังสือภายนอก: รูปสูง 3.0 ซม.; ย่อหน้าเนื้อความทุกย่อหน้า `jc=thaiDistribute` เยื้องบรรทัดแรก 2.5 ซม. ก่อนย่อหน้า 6 pt; ย่อหน้าคำลงท้ายก่อนย่อหน้า 12 pt; ระหว่างคำลงท้ายกับ "(ชื่อ)" มีย่อหน้าว่าง 3 ย่อหน้า; ชื่อผู้ลงนามอยู่ในวงเล็บ; วันที่เยื้องซ้าย 8.0 ซม.
  - หนังสือประทับตรา: มี "ถึง" ก่อนย่อหน้า 8 pt; ไม่มีคำลงท้าย
  - `[ระบุ…]`: run ที่ข้อความตรง `PLACEHOLDER_RE` มีไฮไลต์เหลือง run อื่นไม่มี
  - เลข: ค่าเริ่มต้นข้อความทั้งเอกสารไม่มี `[0-9]` ยกเว้นใน URL/อีเมล; `spec.thai_numerals=False` ไม่มี `[๐-๙]` ที่มาจากเลขอารบิกในร่าง
  - ชั้นความเร็ว `speed="ด่วนที่สุด"`: run ตัวหนา ≥ 32 pt สี `FF0000`
  - Review Focus 3: `references=[]`, `attachments=[]`, `copies_to=[]`, `fax=None` → ไม่มีย่อหน้าขึ้นต้น "อ้างถึง" "สิ่งที่ส่งมาด้วย" "สำเนาส่ง" "โทรสาร"; `signer.name=""` → ไม่มีข้อความ "()"
  - `spec.omit_signature=("memo",)` → ไม่มีย่อหน้าชื่อ/ตำแหน่งผู้ลงนาม
  - ไม่ส่ง `emblem` → `used_placeholder_emblem is True`; ส่งไฟล์ PNG → `False`
- [ ] **Step 2:** รันเทส คาดว่า FAIL
- [ ] **Step 3:** นำ `DocxRenderer` มาจากต้นทางเฉพาะ: ตั้งเอกสาร ตัวช่วย (`_run`, `_par`, `_line`, `_garuda*`, `_classification`, `_body`, `_centered*`, `_signature`, `_labelled`, `_numbered`, `_contact_block`, `_separator`, `_top_gap`, `_fit_one_line`, `text_width_cm`, `_inline_to_anchor`) และ `_external`, `_memo`, `_stamp` การเปลี่ยน: ไม่นำ `render_custom`/`appendix`/`body_writer`/`_memo_footer`; ตราใช้ `emblem` หรือ `PLACEHOLDER_EMBLEM` (ไม่มี Pillow); `_run` เรียก `clean_text` ก่อน แล้ว `to_thai_digits` ของแกน; `_body` ผ่าน `split_paragraphs`; `text_width_cm` ใช้ fontTools เมื่อ import ได้ มิฉะนั้นใช้ค่าประมาณ; ขั้นสุดท้าย `layout.finalize(self.d, spec.font, spec.body_pt)` ถ้า `finalize` ทำให้ค่าในเทสเพี้ยน (เช่น keepNext หรือระยะ) ให้แก้ที่ `render.py` (ตั้งค่าซ้ำหลัง `finalize`) ห้ามแก้ `layout.py`
- [ ] **Step 4:** รันเทส PASS; รันเทสทั้ง repo ผ่าน; build + `--check`; commit `feat(official): ตัวสร้างบันทึกข้อความ หนังสือภายนอก หนังสือประทับตรา`

### Task 5: ตัวสร้าง — ชนิดที่เหลือ 8 ชนิด

**Files:**
- Modify: `src/thaidocs/official/render.py`
- Test: `tests/official/test_render_others.py`

**Interfaces:**
- Consumes: `DocxRenderer`, `render_example`, `para` (Task 4)
- Produces: `DocxRenderer.render` รับครบ `AnyDoc` ทุกแบบ

- [ ] **Step 1:** เขียนเทส:
  - parametrize ทั้ง 11 ชนิด: สร้างได้ เปิดได้ `thaidocs.check.check(doc) == []` ขอบกระดาษตาม spec ไม่มีเลขอารบิกนอก URL/อีเมล
  - คำสั่ง: บรรทัดแรกหลังตราขึ้นต้น "คำสั่ง" ตัวหนา กึ่งกลาง; มี "ที่ " ตามด้วยเลขที่; มี "สั่ง ณ วันที่"; รายชื่อกรรมการจากตัวอย่างปรากฏครบทุกชื่อและเรียงตามร่าง; คณะย่อยมีเลข "๒.๑"
  - ระเบียบ/ข้อบังคับ: มี "ว่าด้วย"; ทุก `Article` ขึ้นต้น "ข้อ ๑" "ข้อ ๒" … ต่อเนื่อง; `chapter` ปรากฏเป็นย่อหน้ากึ่งกลางก่อนข้อนั้น; ข้อบังคับใช้คำ "ข้อบังคับ"
  - ประกาศ: `word="แจ้งความ"` → หัวเป็น "แจ้งความ"; มี "ประกาศ ณ วันที่"
  - แถลงการณ์/ข่าว: ไม่มีย่อหน้าชื่อผู้ลงนาม; ข่าวใช้คำ "ข่าว"; `number` ปรากฏเป็น "ฉบับที่"
  - หนังสือรับรอง: มี "เลขที่"; มี "ให้ไว้ ณ วันที่"; `photo_holder` → มีข้อความลงชื่อผู้ได้รับการรับรอง
  - รายงานการประชุม: มี "ผู้มาประชุม" "ผู้ไม่มาประชุม" "เริ่มประชุมเวลา" "เลิกประชุมเวลา"; ทุกวาระมีชื่อวาระและ "มติที่ประชุม"; มี "ผู้จดรายงานการประชุม"
  - Review Focus 2: คำสั่งที่ `committees[0].members` 60 คน และ `paragraphs` ยาวรวม ≥ 6,000 อักขระ สร้างได้ ชื่อคนที่ 60 อยู่ในเอกสาร; เรื่องยาว 200 อักขระในหนังสือภายนอกสร้างได้และข้อความเรื่องครบ
- [ ] **Step 2:** รันเทส คาดว่า FAIL (KeyError ชนิดที่ยังไม่รองรับ)
- [ ] **Step 3:** นำ `_title_block`, `_order_body`, `_issued`, `_issuer_signature`, `_order`, `_committee`, `_members`, `_duties`, `_regulation`, `_announcement`, `_statement`, `_certificate`, `_minutes` มาจากต้นทาง ด้วยกติกาเปลี่ยนเดียวกับ Task 4
- [ ] **Step 4:** รันเทส PASS; build + `--check`; commit `feat(official): ตัวสร้างคำสั่ง ระเบียบ ข้อบังคับ ประกาศ แถลงการณ์ ข่าว หนังสือรับรอง รายงานการประชุม`

### Task 6: เติมส่วนตายตัว และตัวตรวจ

**Files:**
- Create: `src/thaidocs/official/fill.py`, `checker.py`
- Test: `tests/official/test_fill.py`, `tests/official/test_checker.py`

**Interfaces:**
- Consumes: `AgencyInfo`, `LayoutSpec`, `NATIONAL_FONTS`, `is_placeholder`, `salutations`, โมเดลจาก `schema`
- Produces (`fill`): `fill(doc: AnyDoc, agency: AgencyInfo | None = None) -> AnyDoc` (คืนสำเนา ไม่แก้ของเดิม); `forward_paragraphs(issuer: str, letters: list[ForwardedLetter]) -> list[str]`
- Produces (`checker`): `Finding(level: Literal["error","warning","info"], code: str, message: str, where: str = "", ref: str = "")`; `check_document(doc: AnyDoc, spec: LayoutSpec = DEFAULT_SPEC) -> list[Finding]`; `check_docx(path: Path, spec: LayoutSpec = DEFAULT_SPEC) -> list[Finding]`; `summarize(findings) -> dict` (keys `errors, warnings, infos, findings`)

- [ ] **Step 1:** เขียน `test_fill.py`:
  - หนังสือภายนอก `salutation`/`closing` เป็นค่าเริ่มต้นและ `to="นายกรัฐมนตรี"` → ได้ "กราบเรียน"/"ขอแสดงความนับถืออย่างยิ่ง"
  - ร่างระบุ `salutation="นมัสการ"`, `closing="ขอนมัสการด้วยความเคารพ"` → คงเดิม (ต่างจากต้นทาง)
  - `subject="เรื่อง ขออนุมัติ"` → "ขออนุมัติ"; `to="เรียน อธิบดี"` → "อธิบดี"; `signer.name="(นายสมมติ ตัวอย่าง)"` → ไม่มีวงเล็บ
  - มี `agency`: `agency`/`owner_unit`/`phone`/`email`/`signer` ที่เป็น `[ระบุ…]` หรือว่างถูกเติม ค่าที่ร่างระบุแล้วไม่ถูกทับ; `agency=None` → ไม่เปลี่ยน
  - หนังสือประทับตรามี `forwarded` 1, 2 ฉบับ → ย่อหน้าเดียวเชื่อมด้วย "และสำเนา" ตามด้วย "จึงเรียนมาเพื่อทราบ"; 3 ฉบับ → "จำนวน ๓ ฉบับ ดังนี้" แล้วแยกข้อ; ถ้าร่างมี `paragraphs` เองอยู่แล้ว → ไม่ทับ
  - Review Focus 5: ย่อหน้าที่มี `\n` ถูกแยก; รายการ "วันที่: … / เวลา: …" ถูกรวมเป็นความเรียง; วัตถุต้นฉบับไม่ถูกแก้
- [ ] **Step 2:** เขียน `test_checker.py`: สำหรับรหัสแต่ละตัวต่อไปนี้ มีกรณีที่รายงานและกรณีที่ไม่รายงาน (ตัวอย่างที่ผ่าน `fill` แล้วต้องไม่มีระดับ error เลย): `PLACEHOLDER`, `DATE_FORMAT`, `DATE_BE`, `DOC_NO`, `SUBJECT_LONG`, `BODY_CAUSE`, `BODY_PURPOSE`, `SALUTATION_PAIR`, `SALUTATION_RANK`, `OWNER_UNIT`, `ORDER_NO`, `ORDER_EFFECTIVE`, `REG_ARTICLES`, `REG_NAME`, `REG_EFFECTIVE`, `REG_KEEPER`, `BYLAW_AUTHORITY`, `ANN_BODY`, `CERT_OPENING`, `MIN_AGENDA`, `MIN_RESOLUTION`, `TEXT_PUA`, `TEXT_SARA_AM`, `TEXT_MIXED_DIGITS`, `TEXT_BRACKET`, `OPENING_REF_NAN`, `CONCLUSION_RANK`; ระดับ: `MEMO_PROPOSAL` และ `STAMP_CONCLUSION` เป็น `info`; `STAMP_NOTIFY_ONLY` รายงานเฉพาะเมื่อ `spec.stamp_notify_only=True`; Finding ระดับ error ทุกตัวมี `ref` ไม่ว่าง; ไม่มี Finding ใดที่ `code` ขึ้นต้น `PWA_`; ไม่มีข้อความใดมีชื่อหน่วยงานเฉพาะ; `SALUTATION_PAIR` กับคำขึ้นต้นที่ตารางไม่รู้จัก ("นมัสการ") เป็น `warning` ไม่ใช่ `error`; `check_docx` กับไฟล์จากตัวสร้าง → ไม่มี error; ไฟล์ที่แก้ขอบซ้ายเป็น 2 ซม. → `MARGIN`; ฟอนต์ Arial → `FONT`; ไฟล์ที่ไม่มีรูป → `GARUDA` (warning); บันทึกข้อความที่แก้หัวเป็น 16 pt → `MEMO_TITLE`
- [ ] **Step 3:** รันเทส คาดว่า FAIL
- [ ] **Step 4:** เขียน `fill.py` จากส่วน `_postprocess`/`forward_paragraphs` ของต้นทาง ด้วยกติกา "เติมเฉพาะค่าเริ่มต้นหรือ `[ระบุ…]`"; นำ `checker.py` มา ตัด `_check_pwa`, `date_be_rules`, การตรวจด้วย AI; เปลี่ยนระดับ/ข้อความตามเทส; ข้อความที่อ้างแนวปฏิบัติหน่วยงานเขียนใหม่เป็นคำแนะนำทั่วไป
- [ ] **Step 5:** รันเทส PASS; build + `--check`; commit `feat(official): เติมส่วนตายตัวและตัวตรวจร่าง/ไฟล์`

### Task 7: คำสั่ง `thaigov`

**Files:**
- Create: `src/thaidocs/official/cli.py`, `src/thaidocs/official/__main__.py`
- Test: `tests/official/test_cli.py`; เพิ่มใน `tests/test_build.py`

**Interfaces:**
- Consumes: ทุกโมดูลก่อนหน้า; จาก `thaidocs.cli`: `CliError`, `_utf8`, `_save`, `_segmenter`; `thaidocs.wrap.wrap(doc, seg)`
- Produces: `main(argv: list[str] | None = None) -> int`; คำสั่ง `types`, `schema TYPE`, `example TYPE [OUT.json]`, `render IN.json OUT.docx [--emblem FILE] [--profile FILE] [--font NAME] [--arabic-digits] [--wrap-words] [--words FILE ...] [--force]`, `check IN.json|IN.docx [--profile FILE]`; exit code: 0 สำเร็จ, 1 `check` พบระดับ error, 2 ข้อผิดพลาดการใช้งาน/ข้อมูลเข้า

- [ ] **Step 1:** เขียนเทส (เรียก `main([...])` ใช้ `capsys`, `tmp_path`, `HOME` ชี้ `tmp_path`):
  - `types` พิมพ์ 11 บรรทัด แต่ละบรรทัดมีรหัส ชื่อไทย และข้อของระเบียบ
  - `schema memo` มีชื่อ field ทุกตัวของ `MemoDoc` พร้อมคำอธิบายไทย; `schema xyz` → exit 2 พร้อมรายการชนิด
  - `example order out.json` → ไฟล์เป็น JSON ที่ `parse_doc` ผ่าน; ไม่ใส่ OUT → พิมพ์ทาง stdout
  - `render` ทั้ง 11 ตัวอย่าง → exit 0 ไฟล์เปิดได้ stdout มี "สร้างไฟล์" และจำนวน error/warning/info และบรรทัดบอกว่าใช้กรอบแทนตรา
  - `render` เมื่อ OUT มีอยู่ → exit 2; ใส่ `--force` → เขียนทับได้; OUT ไม่ลงท้าย `.docx` → exit 2
  - `--arabic-digits` → ไม่มีเลขไทยที่มาจากเลขในร่าง; `--font "TH SarabunPSK"` → run ใช้ฟอนต์นั้น; `--wrap-words` → `thaidocs.ooxml.is_wrapped(doc)` จริง
  - `--profile` ที่มี `agency` → ค่าเติมลงเอกสาร; profile เสีย → exit 2 ข้อความไทย
  - Review Focus 1: ไฟล์ไม่ใช่ JSON, JSON มี BOM (ต้องอ่านได้), JSON เป็นรายการไม่ใช่วัตถุ, ไม่มี `doc_type`, field ผิดชนิด → exit 2 (ยกเว้น BOM) stderr มีข้อความไทยชี้ field และไม่มีคำว่า "Traceback"
  - Review Focus 4: `--emblem` ชี้ไฟล์ที่ไม่มี, โฟลเดอร์, ไฟล์ข้อความ → exit 2
  - `check ร่าง.json` ของตัวอย่าง → exit 0; ร่างที่วันที่เป็น "14/11/2026" → exit 1 และพิมพ์ `DATE_FORMAT`; `check ไฟล์.docx` จาก `render` → exit 0; `check` กับนามสกุลอื่น → exit 2
  - `check ร่าง.json --profile p.json` ไม่รายงาน `PLACEHOLDER` ของ field ที่ profile เติมให้
  - จำลอง import pydantic ไม่ได้ → exit 2 ข้อความมี `pip install pydantic`
  - ใน `tests/test_build.py`: `test_official_script_runs` แตก `thai-official-letters.zip` ลง `tmp_path` รัน `python <แตก>/thai-official-letters/scripts/thaigov.py render <ตัวอย่างจาก example> out.docx` ด้วย `cwd=tmp_path` และ `PYTHONPATH` ว่าง → exit 0
- [ ] **Step 2:** รันเทส คาดว่า FAIL
- [ ] **Step 3:** เขียน `cli.py`: ลำดับของ `render` = อ่าน JSON (`utf-8-sig`) → `parse_doc` → `load_profile` → ปรับ spec ด้วย `--font`/`--arabic-digits` (`dataclasses.replace`) → `fill` → `resolve_emblem` → `DocxRenderer.render` → (ถ้า `--wrap-words`) เปิด bytes ด้วย python-docx แล้ว `wrap` → `_save` → `check_document` + `check_docx` → พิมพ์สรุปและรายการระดับ error/warning; รูปแบบ `schema TYPE` เป็นข้อความอ่านง่ายบรรทัดละ field (`ชื่อ (ชนิด, จำเป็น/ค่าเริ่มต้น): คำอธิบาย`) โมเดลย่อยแสดงเยื้องใต้ field แม่
- [ ] **Step 4:** รันเทส PASS; build + `--check`; รันเทสทั้ง repo; commit `feat(official): คำสั่ง thaigov (types, schema, example, render, check)`

### Task 8: เอกสาร skill

**Files:**
- Create: `plugins/thai-docs/skills/thai-official-letters/SKILL.md`, `references/doc-types.md`, `references/writing.md`, `references/layout-w2019.md`, `references/agency-profile.md`, `references/preview.md`
- Modify: `tests/test_build.py`

- [ ] **Step 1:** เพิ่มเทส: frontmatter ของ SKILL.md มี `name: thai-official-letters` และ `description` ≤ 1024 อักขระที่มีชื่อชนิดหนังสืออย่างน้อย 5 ชนิด; SKILL.md ≤ 200 บรรทัด; ทุกไฟล์ที่ SKILL.md อ้างใน `references/` มีอยู่จริง; ทุกคำสั่งและตัวเลือกใน `cli._parser()` ปรากฏใน SKILL.md; รหัส Finding ทุกตัวที่ `checker.py` ใช้ ปรากฏใน `references/doc-types.md` หรือ SKILL.md; ไม่มีคำว่า "Claude" "Codex" "Antigravity" ใน SKILL.md และ references
- [ ] **Step 2:** รันเทส คาดว่า FAIL
- [ ] **Step 3:** เขียน SKILL.md ตามหัวข้อ "ขั้นตอนที่ SKILL.md สอน AI" ของ spec ด้วยโครงเดียวกับ SKILL.md ของ `thai-docx` (ใช้เมื่อไร, เตรียมเครื่อง รวม `pip install python-docx lxml pydantic`, ขั้นตอน, ตารางคำสั่ง/ตัวเลือก, ตรวจก่อนส่ง, อ่านเพิ่ม) ต้องมีข้อความ "ห้ามแต่งชื่อ ตำแหน่ง เลขที่หนังสือ วันที่ ที่ผู้ใช้ไม่ได้ให้ ให้ใส่ `[ระบุ…]`" และ "บอกผู้ใช้ว่าต้องใส่ตราเอง ออกเลข และลงนามเอง"
- [ ] **Step 4:** เขียน references: `doc-types.md` (ตาราง 11 ชนิด: ใช้เมื่อไร ข้อของระเบียบ กระดาษ/ตรา ผู้ลงนาม + ตารางรหัสผลตรวจ ความหมาย วิธีแก้); `writing.md` เขียนใหม่จากระเบียบงานสารบรรณและกติกาที่ `checker.py` ใช้ (โครง 3 ภาค คำเปิด "ด้วย/ตาม" คำลงท้ายภาคสรุป การเขียน "เรื่อง" คำขึ้นต้น/คำลงท้ายชุดที่ตารางมี พร้อมบอกว่ากรณีอื่นให้ดูภาคผนวก 2 ของระเบียบ) **ห้ามคัดหรือเรียบเรียงจากคู่มือการเขียนของต้นทาง**; `layout-w2019.md` สร้างตารางจาก `LayoutSpec` (ค่า หน่วย ที่มา); `agency-profile.md` (ตัวอย่างไฟล์ ลำดับการหาตรา key ใน `DOCUMENTED_LAYOUT_KEYS`); `preview.md` คัดจากของ `thai-docx`
- [ ] **Step 5:** รันเทส PASS; build + `--check`; commit `docs(official): SKILL.md และ references ของ thai-official-letters`

### Task 9: ส่วนกลางของ repo และเทสความสะอาด

**Files:**
- Modify: `README.md`, `NOTICE`, `plugins/thai-docs/.claude-plugin/plugin.json`, `.claude-plugin/marketplace.json`, `.github/workflows/test.yml`, `.github/workflows/release.yml`, `site/index.html`, `tests/test_build.py`
- Create: `tests/test_hygiene.py`

- [ ] **Step 1:** เขียน `tests/test_hygiene.py`: ไล่ไฟล์จาก `git ls-files` (ข้ามไฟล์ไบนารี รายการคำใน `data/` และ `_khammoon/model/`) ต้องไม่พบ regex `/Users/[A-Za-z]`, `/home/[a-z]`, `[A-Z]:\\Users\\`; ที่อยู่อีเมลที่พบต้องลงท้าย `@example.com`, `@users.noreply.github.com` หรือ `@anthropic.com`; ไม่มีไฟล์ที่ชื่อลงท้าย ` 2` หรือ ` 2.<นามสกุล>` เพิ่มใน `tests/test_build.py`: รุ่นใน `pyproject.toml`, `plugin.json`, `marketplace.json` (ทั้งสองจุด) เป็น `0.2.0` ตรงกัน; `release.yml` แนบ zip ครบ 4 ไฟล์
- [ ] **Step 2:** รันเทส คาดว่า FAIL (รุ่นและ release.yml)
- [ ] **Step 3:** แก้ไฟล์: manifests รุ่น 0.2.0 และคำอธิบายรวมหนังสือราชการ; `test.yml` คงคำสั่งเดิม (extra `dev` มี pydantic แล้ว); `release.yml` แนบ `thai-official-letters.zip` และ `thai-official-letters-nokhammoon.zip`; `README.md` เพิ่มส่วน skill ใหม่ (ทำอะไรได้ ตัวอย่างคำสั่ง ติดตั้ง ข้อจำกัด: ไม่แจกตรา คำขึ้นต้น/คำลงท้ายเฉพาะชุดที่พบบ่อย ยังไม่ได้ตรวจกับ Word for Windows); `NOTICE` เพิ่มว่า `official/` ดัดแปลงจาก SARABUN.AI (Apache-2.0) และไม่ให้สิทธิ์ใช้ตราครุฑ/ตราหน่วยงาน; `site/index.html` เพิ่มส่วนหนังสือราชการ (แก้ไฟล์เท่านั้น ไม่เผยแพร่)
- [ ] **Step 4:** รันเทสทั้ง repo PASS; build + `--check`; `.venv/bin/python tools/build.py --zip` แล้วรายงานขนาด zip ทั้ง 4; commit `chore: รุ่น 0.2.0 เอกสารและ workflow ของ skill หนังสือราชการ`

### Task 10: ตรวจด้วยตา และตรวจรอบสุดท้าย (ผู้พัฒนา ไม่ push)

**Files:** ไม่มีไฟล์ใหม่ใน repo (ไฟล์ทดสอบอยู่ในโฟลเดอร์เอกสารของ Word แล้วลบ)

- [ ] **Step 1:** สร้างไฟล์จากตัวอย่าง 4 ชนิด (`external_letter`, `memo`, `stamp_letter`, `order`) ด้วย `thaigov render` แปลงเป็น PDF ด้วย Word for Mac ทีละไฟล์ (เบื้องหลัง ไฟล์ต้องอยู่ในโฟลเดอร์เอกสารของ Word) แล้วเปิดดูทุกหน้า
- [ ] **Step 2:** เทียบกับตัวอย่างแนบ ว 2019: ตำแหน่งตรา บรรทัด "ที่"/ส่วนราชการ/วันที่ เส้นประของบันทึกข้อความ คำลงท้ายและชื่อผู้ลงนามอยู่กึ่งกลางที่ 9.5 ซม. เลขหน้า "- ๒ -" ที่หน้าสอง ไม่มีเส้นหยักแดง จุดที่ไม่ตรงให้แก้ใน `render.py` พร้อมเทสที่ตรึงค่านั้น แล้ววนซ้ำ
- [ ] **Step 3:** ลบไฟล์ทดสอบทั้งหมดออกจากโฟลเดอร์เอกสารของ Word
- [ ] **Step 4:** ตรวจสุดท้าย: `.venv/bin/python -m pytest -q` ผ่าน; `tools/build.py --check` ผ่าน; `git status` สะอาด ไม่มีไฟล์ ` 2`; `git diff public-main --stat -- plugins/thai-docs/skills/thai-docx` ว่าง; ค้นด้วยมือ (`git grep`) หาชื่อหน่วยงาน/บุคคลจริงของต้นทางตามรายการในบันทึกในเครื่อง ไม่พบ
- [ ] **Step 5:** รายงานผู้ใช้: ผลเทส ขนาด zip ภาพหน้า PDF ที่ตรวจ สิ่งที่ยังไม่ได้ทดสอบ (Word for Windows, LibreOffice) แล้ว **หยุดรอผู้ใช้อนุมัติ** ก่อน push / release
