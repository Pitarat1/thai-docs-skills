"""สร้าง .docx จากโครงข้อมูลหนังสือราชการ

โครงสร้าง: ระเบียบสำนักนายกรัฐมนตรีว่าด้วยงานสารบรรณ พ.ศ. ๒๕๒๖ และที่แก้ไขเพิ่มเติม (ลำดับส่วนประกอบตามข้อ ๑๑–๒๕)
การพิมพ์:  แนวทางตาม นร ๐๑๐๖/ว ๒๐๑๙ (ค่าทั้งหมดอยู่ใน style.LayoutSpec)

ระยะห่างใช้แบบ "Before" ตามแนวทาง: ทุกส่วนเป็นระยะบรรทัดปกติ แล้วเพิ่ม Before ๖ หรือ ๑๒ พอยท์
ไม่ใส่บรรทัดว่างระหว่างภาค ยกเว้นช่วงลงลายมือชื่อ (๔ Enter) และส่วนราชการเจ้าของเรื่อง (๔ Enter)

ดัดแปลงจากโครงการต้นทางภายใต้ Apache License 2.0 ดู NOTICE ของ repo
ไม่พึ่ง pydantic/Pillow; fontTools ใช้เมื่อมี (วัดความกว้างข้อความ) ไม่มีก็ใช้ค่าประมาณ
"""
from __future__ import annotations

import io
import re
import unicodedata
from pathlib import Path
from typing import Iterable, Optional

from docx import Document
from docx.enum.text import (WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX, WD_LINE_SPACING, WD_TAB_ALIGNMENT, WD_TAB_LEADER,
                            WD_UNDERLINE)
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

from .. import layout
from ..digits import to_thai_digits
from .profile import PLACEHOLDER_EMBLEM
from .schema import DocType
from .style import DEFAULT_SPEC, LayoutSpec
from .textutil import PLACEHOLDER_RE, clean_text, protect_breaks, split_paragraphs

_SPLIT_PH = re.compile(f"({PLACEHOLDER_RE.pattern})")
CENTER = WD_ALIGN_PARAGRAPH.CENTER
DIST = "distribute"   # กระจายแบบไทย (thaiDistribute)


class DocxRenderer:
    def __init__(self, spec: LayoutSpec = DEFAULT_SPEC, emblem: Optional[Path] = None):
        self.spec = spec
        # ตรา: ไฟล์ที่ส่งมา > layout.emblem_png > กรอบแทนตรา (ผู้เรียกตรวจไฟล์ด้วย profile.resolve_emblem)
        src = emblem or spec.emblem_png
        self.emblem = Path(src) if src else None
        self.used_placeholder_emblem = False

    # ------------------------------------------------------------------ public
    def render(self, doc) -> bytes:
        """สร้างเอกสาร คืนเนื้อไฟล์ .docx"""
        handler = {
            "memo": self._memo, "external_letter": self._external, "stamp_letter": self._stamp,
            "order": self._order, "regulation": self._regulation, "bylaw": self._regulation,
            "announcement": self._announcement, "statement": self._statement, "news": self._statement,
            "certificate": self._certificate, "minutes": self._minutes,
        }.get(doc.doc_type.value)
        if handler is None:
            raise NotImplementedError(f"ยังไม่รองรับการสร้างหนังสือชนิด {doc.doc_type.value}")
        self.d = self._new_document()
        self.used_placeholder_emblem = False
        self._omit_signature = doc.doc_type.value in self.spec.omit_signature
        handler(doc)
        layout.finalize(self.d, self.spec.font, self.spec.body_pt)   # ต้องทำสุดท้าย: run ไทย + ภาษา + ลำดับ XML
        buf = io.BytesIO()
        self.d.save(buf)
        return buf.getvalue()

    def render_to(self, doc, path: str | Path) -> Path:
        path = Path(path)
        path.write_bytes(self.render(doc))
        return path

    # ------------------------------------------------------------- doc setup
    def _new_document(self):
        s = self.spec
        d = Document()
        sec = d.sections[0]
        sec.page_width, sec.page_height = Cm(s.page_w_cm), Cm(s.page_h_cm)
        sec.left_margin, sec.right_margin = Cm(s.margin_left_cm), Cm(s.margin_right_cm)
        sec.top_margin, sec.bottom_margin = Cm(s.margin_top_cm), Cm(s.margin_bottom_cm)
        sec.header_distance = Cm(s.header_distance_cm)

        normal = d.styles["Normal"]
        normal.font.name = s.font
        normal.font.size = Pt(s.body_pt)
        rPr = normal.element.get_or_add_rPr()
        self._set_rfonts(rPr)
        self._set_cs_size(rPr, s.body_pt)
        self._set_lang_th(rPr)
        pf = normal.paragraph_format
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        pf.line_spacing = 1.0          # [ว2019] ระยะบรรทัด ๑ เท่า

        if s.word_compat_mode:
            for cs in d.settings.element.iter(qn("w:compatSetting")):
                if cs.get(qn("w:name")) == "compatibilityMode":
                    cs.set(qn("w:val"), str(s.word_compat_mode))

        # เลขหน้า: หน้าแรกไม่มี หน้าถัดไป "- ๒ -" กึ่งกลางด้านบน
        sec.different_first_page_header_footer = True
        hp = sec.header.paragraphs[0]
        hp.alignment = CENTER
        pt = s.page_number_pt or s.body_pt
        pre, post = s.page_number_fmt.split("{}")
        if pre:
            self._run(hp, pre, size=pt)
        self._page_field(hp, pt)
        if post:
            self._run(hp, post, size=pt)
        # ขนาดเครื่องหมายย่อหน้าของหัวกระดาษ = ขนาดเลขหน้า (ความสูงหัวกระดาษมีผลต่อบรรทัดแรกของหน้า)
        mark = hp._p.get_or_add_pPr()
        mrpr = OxmlElement("w:rPr")
        for tag in ("w:sz", "w:szCs"):
            e = OxmlElement(tag)
            e.set(qn("w:val"), str(int(pt * 2)))
            mrpr.append(e)
        mark.append(mrpr)
        if s.page_number_pt:
            # ความสูงหัวกระดาษคงที่ = ขนาดเลขหน้า
            hp.paragraph_format.line_spacing = Pt(pt)
            hp.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
            # หัวกระดาษหน้าแรก (ว่าง) สูงเท่ากัน — ตำแหน่งบรรทัดแรกของหน้าแรกจึงตรงกับแม่แบบ
            fp = sec.first_page_header.paragraphs[0]
            fp.paragraph_format.line_spacing = Pt(pt)
            fp.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
            self._run(fp, "", size=pt)
        return d

    # --------------------------------------------------------------- helpers
    def _set_rfonts(self, rPr):
        rFonts = rPr.find(qn("w:rFonts"))
        if rFonts is None:
            rFonts = OxmlElement("w:rFonts")
            rPr.insert(0, rFonts)
        for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
            rFonts.set(qn(attr), self.spec.font)

    @staticmethod
    def _set_cs_size(rPr, pt: float):
        szCs = rPr.find(qn("w:szCs"))
        if szCs is None:
            szCs = OxmlElement("w:szCs")
            rPr.append(szCs)
        szCs.set(qn("w:val"), str(int(pt * 2)))

    @staticmethod
    def _set_lang_th(rPr):
        lang = OxmlElement("w:lang")
        lang.set(qn("w:bidi"), "th-TH")
        rPr.append(lang)

    def _run(self, p, text: str, *, size: Optional[float] = None, bold: bool = False,
             color: Optional[str] = None, underline: bool | str = False):
        """เพิ่มข้อความ (ตัดอักขระควบคุม แปลงเลขไทย) — [ระบุ...] ไฮไลต์สีเหลือง; underline="dotted" = ขีดเส้นใต้แบบจุด"""
        s = self.spec
        text = clean_text(text)
        if s.thai_numerals:
            text = to_thai_digits(text)
        text = protect_breaks(text)
        size = size or s.body_pt
        for chunk in _SPLIT_PH.split(text):
            if not chunk:
                continue
            r = p.add_run(chunk)
            r.font.name = s.font
            r.font.size = Pt(size)
            r.font.bold = bold
            if underline == "dotted":
                r.font.underline = WD_UNDERLINE.DOTTED
            elif underline:
                r.font.underline = True
            if color:
                r.font.color.rgb = RGBColor.from_string(color)
            rPr = r._element.get_or_add_rPr()
            self._set_rfonts(rPr)
            self._set_cs_size(rPr, size)
            if bold:
                rPr.append(OxmlElement("w:bCs"))
            if s.highlight_placeholders and PLACEHOLDER_RE.fullmatch(chunk):
                r.font.highlight_color = WD_COLOR_INDEX.YELLOW
        return p

    def _page_field(self, p, pt: Optional[float] = None):
        r = p.add_run()
        for tag, typ in (("w:fldChar", "begin"), ("w:instrText", None), ("w:fldChar", "end")):
            el = OxmlElement(tag)
            if typ:
                el.set(qn("w:fldCharType"), typ)
            else:
                el.set(qn("xml:space"), "preserve")
                el.text = " PAGE \\* ThaiArabic " if self.spec.thai_numerals else " PAGE "
            r._element.append(el)
        pt = pt or self.spec.body_pt
        r.font.size = Pt(pt)
        self._set_cs_size(r._element.get_or_add_rPr(), pt)

    def _par(self, *, align=None, left: float = 0, first: float = 0, hanging: float = 0,
             before: float = 0, tabs: Iterable[tuple] = (), keep_next: bool = False):
        p = self.d.add_paragraph()
        pf = p.paragraph_format
        if align == DIST:
            jc = OxmlElement("w:jc")
            jc.set(qn("w:val"), "thaiDistribute")
            p._p.get_or_add_pPr().append(jc)
        elif align is not None:
            p.alignment = align
        pf.left_indent = Cm(left)
        if hanging:
            pf.first_line_indent = Cm(-hanging)
        elif first:
            pf.first_line_indent = Cm(first)
        pf.space_before = Pt(before)
        pf.keep_with_next = keep_next
        for pos, kind, leader in tabs:
            pf.tab_stops.add_tab_stop(Cm(pos), kind, leader)
        return p

    def _line(self, text: str = "", **kw):
        bold, size, color = kw.pop("bold", False), kw.pop("size", None), kw.pop("color", None)
        underline = kw.pop("underline", False)
        p = self._par(**kw)
        if text:
            self._run(p, text, bold=bold, size=size, color=color, underline=underline)
        return p

    def _emblem_source(self) -> str:
        if self.emblem and self.emblem.is_file():
            return str(self.emblem)
        self.used_placeholder_emblem = True
        return str(PLACEHOLDER_EMBLEM)

    def _garuda(self, height_cm: float, p=None, *, floating: bool = False, center: bool = False,
                offset: Optional[tuple] = None):
        p = p or self._par(align=CENTER)
        run = p.add_run()
        run.add_picture(self._emblem_source(), height=Cm(height_cm))
        if floating and center:
            _inline_to_anchor(run, y_emu=33655, center=True)       # แผ่นแบบ: V offset 0.09 ซม.
        elif floating and offset:
            _inline_to_anchor(run, x_emu=offset[0], y_emu=offset[1])
        elif floating:
            # ล่างตราเสมอฐานบรรทัด "บันทึกข้อความ" (บรรทัด Exactly ๓๕ พอยท์ เตี้ยกว่าตรา ๑.๕ ซม.)
            _inline_to_anchor(run, y_emu=int(Pt(self.spec.memo_title_exact_pt) - Cm(height_cm)))
        return p

    def _classification(self, doc, *, inline_speed: bool = False):
        """ชั้นความลับ: กึ่งกลางขอบบนและขอบล่างทุกหน้า (ระเบียบว่าด้วยการรักษาความลับของทางราชการ)
        ชั้นความเร็ว: ตัวหนาสีแดง ไม่เล็กกว่า ๓๒ พอยท์ (ข้อ ๒๘)
          inline_speed=True → พิมพ์ในบรรทัดเดียวกับตรา (มุมบนซ้าย) ใช้กับกระดาษตราครุฑ"""
        s = self.spec
        if getattr(doc, "secrecy", None):
            sec = self.d.sections[0]
            for part in (sec.first_page_header, sec.first_page_footer, sec.header, sec.footer):
                p = part.add_paragraph() if part.paragraphs[0].text else part.paragraphs[0]
                p.alignment = CENTER
                self._run(p, doc.secrecy, bold=True, size=s.body_pt + 4, color=s.speed_color)
        if getattr(doc, "speed", None) and not inline_speed:
            self._line(doc.speed, bold=True, size=s.speed_pt, color=s.speed_color)

    def _garuda_header(self, doc=None):
        """หัวกระดาษตราครุฑตามแผ่นแบบทางการ:
        ตรา ๓ ซม. เป็นรูปลอยกึ่งกลางในย่อหน้าแรก + ย่อหน้าว่าง ๒ ย่อหน้า
        แล้วบรรทัด "ที่" (Before ๖ pt) จะอยู่ระดับขาครุฑ — ชั้นความเร็ว (ถ้ามี) พิมพ์ในย่อหน้าแรกชิดซ้าย"""
        s = self.spec
        if doc is not None:
            self._classification(doc, inline_speed=True)
        p = self._par()
        speed = getattr(doc, "speed", None) if doc is not None else None
        if speed:
            self._run(p, speed, bold=True, size=s.speed_pt, color=s.speed_color)
        self._garuda(s.garuda_external_cm, p, floating=True, center=True)
        for _ in range(s.header_blank_paras):
            self._par()
        return p

    def _body(self, paragraphs: Iterable[str], *, keep_tail: bool = False, first_before: Optional[float] = None):
        """ข้อความ: ย่อหน้า ๒.๕ ซม. กระจายแบบไทย ระยะปกติ + Before ๖ พอยท์ทุกย่อหน้า
        keep_tail: ย่อหน้าสุดท้ายอยู่หน้าเดียวกับคำลงท้าย/ลายมือชื่อ"""
        s = self.spec
        items = split_paragraphs(list(paragraphs))
        for i, text in enumerate(items):
            before = s.para_before_pt if (i or first_before is None) else first_before
            self._line(text, align=DIST, first=s.para_indent_cm, before=before,
                       keep_next=keep_tail and i == len(items) - 1)

    def _centered(self, text: str, **kw):
        """จัดกึ่งกลางที่แท็บกลาง signature_center_cm (แบบแผ่นแบบทางการ)"""
        return self._centered_at(text, self.spec.signature_center_cm, **kw)

    def _centered_at(self, text: str, center_cm: float, **kw):
        return self._line(f"\t{text}", tabs=[(center_cm, WD_TAB_ALIGNMENT.CENTER, WD_TAB_LEADER.SPACES)], **kw)

    def _text_cm(self, text: str, size: Optional[float] = None, bold: bool = False) -> float:
        return text_width_cm(to_thai_digits(text) if self.spec.thai_numerals else text, size or self.spec.body_pt,
                             bold=bold)

    def _fit_one_line(self, runs, text: str, avail_cm: float, size: Optional[float] = None):
        """บีบข้อความให้อยู่บรรทัดเดียว: ลดช่องไฟระหว่างอักษร (สูงสุด -๑ พอยท์) ถ้ายังล้นจึงลดขนาดอักษรทีละ ๐.๕ พอยท์
        (ใช้กับบรรทัด "ส่วนราชการ" ของบันทึกข้อความ ซึ่งไม่ควรขึ้นบรรทัดใหม่)"""
        size = size or self.spec.body_pt
        width = self._text_cm(text, size)
        if width <= avail_cm:
            return
        n = sum(1 for ch in text if not unicodedata.combining(ch)) or 1
        spacing_pt = max(-1.0, -(width - avail_cm) / n * 72 / 2.54)
        new_size = size
        while self._text_cm(text, new_size) + n * spacing_pt * 2.54 / 72 > avail_cm and new_size > size - 2:
            new_size -= 0.5
        for r in runs:
            rPr = r._element.get_or_add_rPr()
            sp = OxmlElement("w:spacing")
            sp.set(qn("w:val"), str(int(round(spacing_pt * 20))))
            rPr.append(sp)
            if new_size != size:
                r.font.size = Pt(new_size)
                self._set_cs_size(rPr, new_size)

    def _signature(self, signer, *, blank_lines: Optional[int] = None, center_cm: Optional[float] = None,
                   paren_cm: Optional[float] = None):
        """เว้น ๓ บรรทัด (๔ Enter) ให้ลงลายมือชื่อ → (ชื่อเต็ม) → ตำแหน่ง
        center_cm: ชื่อ/ตำแหน่งกึ่งกลางที่แนวนี้ (ค่าเริ่มต้น = แท็บกลาง signature_center_cm)
        paren_cm:  วงเล็บเปิดของชื่ออยู่ที่แนวนี้ แล้วตำแหน่งกึ่งกลางกับชื่อ (ตามคู่มือหน่วยงาน)
        ชื่อว่าง: ไม่พิมพ์บรรทัดชื่อ (ไม่มี "()" เปล่า)"""
        s = self.spec
        if getattr(self, "_omit_signature", False):     # ลงนามในระบบสารบรรณอิเล็กทรอนิกส์
            return
        for _ in range(s.signature_blank_lines if blank_lines is None else blank_lines):
            self._par(keep_next=True)
        positions = [x for x in signer.position if x.strip()]
        name = f"({signer.name.strip()})" if signer.name.strip() else ""
        if paren_cm is not None:
            if name:
                self._line(name, left=paren_cm, keep_next=bool(positions))
                center_cm = paren_cm + self._text_cm(name) / 2
            else:
                center_cm = s.signature_center_cm
            for i, line in enumerate(positions):
                self._centered_at(line, center_cm, keep_next=i < len(positions) - 1)
            return
        c = center_cm if center_cm is not None else s.signature_center_cm
        if name:
            self._centered_at(name, c, keep_next=bool(positions))
        for i, line in enumerate(positions):
            self._centered_at(line, c, keep_next=i < len(positions) - 1)

    def _issued(self, text: str):
        """สั่ง/ประกาศ/ให้ไว้ ณ วันที่ ... (+ ลายมือชื่อ)"""
        s = self.spec
        if s.issued_blank_line:
            self._par(keep_next=True)
        self._line(text, left=s.issued_date_indent_cm, before=0 if s.issued_blank_line else s.closing_before_pt,
                   keep_next=True)

    def _issuer_signature(self, signer):
        self._signature(signer, paren_cm=self.spec.signer_paren_cm)

    def _separator(self):
        s = self.spec
        if not s.separator_cm:
            self._line("-" * 26, align=CENTER)
            return
        side = (s.text_width_cm - s.separator_cm) / 2
        p = self._par(left=side)
        p.paragraph_format.right_indent = Cm(side)
        pPr = p._p.get_or_add_pPr()
        bdr = OxmlElement("w:pBdr")
        bottom = OxmlElement("w:bottom")
        for k, v in (("w:val", "single"), ("w:sz", "6"), ("w:space", "1"), ("w:color", "auto")):
            bottom.set(qn(k), v)
        bdr.append(bottom)
        pPr.insert(0, bdr)
        p.paragraph_format.line_spacing = Pt(8)
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY

    def _top_gap(self):
        """หนังสือที่ไม่มีตรา: บรรทัดแรกห่างขอบกระดาษบนตามคู่มือหน่วยงาน (no_emblem_top_cm)"""
        s = self.spec
        gap = s.no_emblem_top_cm - s.margin_top_cm
        return Cm(gap).pt if gap > 0 else 0

    def _labelled(self, label: str, value: str, *, hang_cm: float, before: float):
        """ป้ายกำกับ + ข้อความแบบแขวน (บรรทัดที่ ๒ ตรงกับข้อความบรรทัดแรก)"""
        return self._line(f"{label}\t{value}", align=DIST, left=hang_cm, hanging=hang_cm, before=before,
                          tabs=[(hang_cm, WD_TAB_ALIGNMENT.LEFT, WD_TAB_LEADER.SPACES)])

    def _numbered(self, label: str, items: list[str], hang_cm: float, before: float):
        multi = len(items) > 1
        for i, item in enumerate(items, 1):
            self._labelled(label if i == 1 else "", f"{i}. {item}" if multi else item,
                           hang_cm=hang_cm, before=before if i == 1 else 0)

    def _contact_block(self, unit: str, phone: str, fax: Optional[str] = None, email: Optional[str] = None,
                       blank_lines: Optional[int] = None):
        """ส่วนราชการเจ้าของเรื่อง / โทร. / โทรสาร / ไปรษณีย์อิเล็กทรอนิกส์ — ชิดซ้าย"""
        s = self.spec
        for _ in range(s.owner_unit_blank_lines if blank_lines is None else blank_lines):
            self._par()
        for line in unit.split("\n"):
            self._line(line)
        self._line(f"โทร. {phone}")
        if fax:
            self._line(f"โทรสาร {fax}")
        if email:
            self._line(f"ไปรษณีย์อิเล็กทรอนิกส์ {email}")

    def _title_block(self, lines: list[str], *, emblem: bool = True):
        """หัวหนังสือสั่งการ/ประชาสัมพันธ์: ตรากลาง + ชื่อเรื่องกึ่งกลาง + เส้นคั่น
        ชื่อเรื่องตัวธรรมดาหรือตัวหนาตาม title_bold บรรทัดละ ๑ Enter เรื่องยาวขึ้นบรรทัดใหม่ตรงอักษรแรกของเรื่อง"""
        s = self.spec
        if emblem:
            self._garuda(s.garuda_external_cm)
        for i, t in enumerate(lines):
            before = (s.block_before_pt if emblem else self._top_gap()) if i == 0 else 0
            self._line(t, align=CENTER, bold=s.title_bold, before=before)
        self._separator()

    def _order_body(self, paragraphs: list[str], effective: Optional[str] = None):
        """คำสั่ง/ระเบียบ/ข้อบังคับ: order_para_first_only → ย่อหน้าแรก +๖ พอยท์ ย่อหน้าถัดไป ๑ Enter, "ทั้งนี้ ตั้งแต่" +๖ พอยท์"""
        s = self.spec
        if not s.order_para_first_only:
            self._body([*paragraphs, *([effective] if effective else [])])
            return
        for i, t in enumerate(split_paragraphs(list(paragraphs))):
            self._line(t, align=DIST, first=s.para_indent_cm, before=s.para_before_pt if i == 0 else 0)
        if effective:
            self._line(effective, align=DIST, first=s.para_indent_cm, before=s.para_before_pt)

    # ------------------------------------------------------------ doc types
    def _external(self, x):
        """ข้อ ๑๑ หนังสือภายนอก (แบบที่ ๑) — กระดาษตราครุฑ"""
        s = self.spec
        self._garuda_header(x)
        agency = x.agency or ["[ระบุส่วนราชการเจ้าของหนังสือ]"]
        self._line(f"ที่ {x.doc_no}\t{agency[0]}", before=s.block_before_pt,
                   tabs=[(s.agency_tab_cm, WD_TAB_ALIGNMENT.LEFT, WD_TAB_LEADER.SPACES)])
        for line in agency[1:]:
            self._line(line, left=s.agency_tab_cm)
        self._line(x.date, left=s.date_indent_cm, before=s.block_before_pt)
        self._labelled("เรื่อง", x.subject, hang_cm=s.subject_hang_cm, before=s.block_before_pt)
        self._labelled(x.salutation, x.to, hang_cm=s.subject_hang_cm, before=s.block_before_pt)
        if x.references:
            self._numbered("อ้างถึง", x.references, hang_cm=s.subject_hang_cm, before=s.block_before_pt)
        if x.attachments:
            self._numbered("สิ่งที่ส่งมาด้วย", x.attachments, hang_cm=s.attachment_hang_cm, before=s.block_before_pt)
        self._body(x.body.paragraphs(), keep_tail=True)
        if s.closing_mode == "start_center":
            # ๒ Enter → คำลงท้าย อักษรแรกตรงแนวกึ่งกลางตรา → ๔ Enter → ชื่อ/ตำแหน่งกึ่งกลางกับคำลงท้าย
            self._par(keep_next=True)
            self._line(x.closing, left=s.center_cm, keep_next=True)
            self._signature(x.signer, center_cm=s.center_cm + self._text_cm(x.closing) / 2)
        else:
            self._centered(x.closing, before=s.closing_before_pt, keep_next=True)
            self._signature(x.signer)
        self._contact_block(x.owner_unit, x.phone, x.fax, x.email)
        if x.copies_to:
            self._line("สำเนาส่ง", before=s.block_before_pt)
            for c in x.copies_to:
                self._line(c, left=1.25)

    def _memo(self, m):
        """ข้อ ๑๒ หนังสือภายใน (แบบที่ ๒) — กระดาษบันทึกข้อความ
        ค่าเริ่มต้น = แผ่นแบบบันทึกข้อความของ ว ๒๐๑๙ / ค่า memo_* ใน LayoutSpec ปรับตามแม่แบบหน่วยงาน"""
        s = self.spec
        W = s.text_width_cm
        if s.memo_margin_top_cm is not None:
            self.d.sections[0].top_margin = Cm(s.memo_margin_top_cm)
        self._classification(m, inline_speed=s.memo_speed_tab_cm is not None)
        # ตรามุมซ้าย (รูปลอย) + "บันทึกข้อความ" ๒๙ พอยท์ ตัวหนา ระยะบรรทัด Exactly
        tabs = []
        if s.memo_speed_tab_cm is not None:
            tabs.append((s.memo_speed_tab_cm, WD_TAB_ALIGNMENT.LEFT, WD_TAB_LEADER.SPACES))
        if s.memo_title_tab_cm is not None:
            tabs.append((s.memo_title_tab_cm, WD_TAB_ALIGNMENT.LEFT, WD_TAB_LEADER.SPACES))
        else:
            tabs.append((W / 2, WD_TAB_ALIGNMENT.CENTER, WD_TAB_LEADER.SPACES))
        p = self._par(tabs=tabs, before=s.memo_title_before_pt)
        pf = p.paragraph_format
        pf.space_after = Pt(s.memo_title_after_pt)
        pf.line_spacing = Pt(s.memo_title_exact_pt)
        pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        self._garuda(s.garuda_memo_cm, p, floating=True, offset=s.memo_emblem_offset_emu)
        if s.memo_speed_tab_cm is not None:
            self._run(p, "\t")
            if m.speed:
                self._run(p, m.speed, size=s.speed_pt, bold=s.memo_speed_bold, color=s.speed_color)
        r0 = len(p.runs)
        self._run(p, "\tบันทึกข้อความ", size=s.memo_title_pt, bold=True)
        if s.memo_title_char_spacing_pt:
            for r in p.runs[r0:]:
                sp = OxmlElement("w:spacing")
                sp.set(qn("w:val"), str(int(s.memo_title_char_spacing_pt * 20)))
                r._element.get_or_add_rPr().append(sp)

        # เส้นประหัวบันทึก = เส้นใต้แบบจุดของข้อความและแท็บ (ไม่ใช้จุดนำของแท็บ ซึ่งวางสูงกว่าเส้นใต้จนเห็นเป็น ๒ ชั้น)
        dot = WD_TAB_LEADER.SPACES
        dept, phone = m.department, None
        if s.memo_phone_tab_cm is not None:
            mm = re.search(r"\s*(โทร\.?.*)$", m.department)
            if mm:
                dept, phone = m.department[:mm.start()].strip(), mm.group(1).strip()
        if phone:
            p = self._par(tabs=[(s.memo_phone_tab_cm, WD_TAB_ALIGNMENT.LEFT, dot), (W, WD_TAB_ALIGNMENT.RIGHT, dot)])
            self._run(p, s.unit_word, size=s.memo_label_pt, bold=True)
            self._run(p, f"  {dept}\t{phone}\t", underline="dotted")
        else:
            p = self._par(tabs=[(W, WD_TAB_ALIGNMENT.RIGHT, dot)])
            self._run(p, s.unit_word, size=s.memo_label_pt, bold=True)
            r0 = len(p.runs)
            self._run(p, f"  {dept}\t", underline="dotted")
            avail = W - self._text_cm(s.unit_word, s.memo_label_pt, bold=True) - 0.3   # เหลือที่ให้เส้นประท้ายบรรทัด
            self._fit_one_line(p.runs[r0:], f"  {dept}", avail)
        p = self._par(tabs=[(s.memo_date_tab_cm, WD_TAB_ALIGNMENT.LEFT, dot), (W, WD_TAB_ALIGNMENT.RIGHT, dot)])
        self._run(p, "ที่", size=s.memo_label_pt, bold=True)
        self._run(p, f"  {m.doc_no}\t", underline="dotted")
        self._run(p, "วันที่", size=s.memo_label_pt, bold=True)
        self._run(p, f"  {m.date}\t", underline="dotted")
        p = self._par(tabs=[(W, WD_TAB_ALIGNMENT.RIGHT, dot)])
        self._run(p, "เรื่อง", size=s.memo_label_pt, bold=True)
        self._run(p, f"  {m.subject}\t", underline="dotted")

        self._labelled(m.salutation, m.to, hang_cm=s.salutation_hang_cm, before=s.block_before_pt)  # คำขึ้นต้น + Before ๖
        if m.sections:
            # บันทึกเสนอเรื่องแบบมีหัวข้อ: ๑. เรื่องเดิม ๒. ข้อเท็จจริง ๓. ข้อกฎหมาย ๔. ข้อพิจารณา ๕. ข้อเสนอ
            # หัวข้อ: เลขที่ ๒.๕ ซม. + Before ๖ พอยท์ / ข้อความ: ย่อหน้า
            first = s.memo_section_first_cm if s.memo_section_first_cm is not None else s.para_indent_cm + 0.75
            for i, sec in enumerate(m.sections, 1):
                last = i == len(m.sections)
                self._line(f"{i}.\t{sec.heading}", left=s.para_indent_cm + 0.63, hanging=0.63, before=s.para_before_pt,
                           keep_next=True, tabs=[(s.para_indent_cm + 0.63, WD_TAB_ALIGNMENT.LEFT, WD_TAB_LEADER.SPACES)])
                texts = split_paragraphs(sec.paragraphs)
                for j, t in enumerate(texts):
                    self._line(t, align=DIST, first=first, keep_next=last and j == len(texts) - 1)
        else:
            self._body(m.body.paragraphs(), keep_tail=True)
        # ๔ Enter หลังภาคสรุป → (ชื่อ) / ตำแหน่ง
        paren = s.memo_date_tab_cm if s.memo_signature_mode == "paren" else None
        self._signature(m.signer, paren_cm=paren)

    def _stamp(self, x):
        """ข้อ ๑๔ หนังสือประทับตรา (แบบที่ ๓) — กระดาษตราครุฑ"""
        s = self.spec
        self._garuda_header(x)
        self._line(f"ที่ {x.doc_no}", before=s.block_before_pt)
        self._labelled("ถึง", x.to, hang_cm=1.0, before=s.stamp_to_before_pt)
        for _ in range(s.stamp_blank_after_to):                        # ว ๒๐๑๙: ๒ Enter + Before ๖
            self._par()
        self._body(x.paragraphs, keep_tail=True)
        # ๔ Enter → ชื่อส่วนราชการ + วัน เดือน ปี กึ่งกลางที่ ๙.๕ ซม.
        # ตราชื่อส่วนราชการ (วงนอก ๔.๕ / วงใน ๓.๕ ซม. หมึกแดง) ประทับให้เส้นผ่าศูนย์กลางผ่านชื่อส่วนราชการ
        # + หัวหน้าส่วนราชการระดับกอง/ผู้ได้รับมอบหมาย ลงชื่อย่อกำกับตรา
        for _ in range(3):
            self._par(keep_next=True)
        if s.stamp_initials:
            # ตราหน่วยงาน (หมึกแดง) ขอบนอกวางบนแนวกึ่งกลาง + ลงชื่อย่อกำกับตราทางขวา
            c = s.center_cm + 1.4
            p = self._centered_at(x.issuer, c, keep_next=True)
            p.paragraph_format.tab_stops.add_tab_stop(Cm(c + 3.2), WD_TAB_ALIGNMENT.LEFT, WD_TAB_LEADER.SPACES)
            self._run(p, "\t(ลงชื่อย่อกำกับตรา)")
            self._centered_at(x.date, c)
        else:
            self._centered(x.issuer, keep_next=True)
            self._centered(x.date)
        self._contact_block(x.owner_unit, x.phone, x.fax, email=x.email)

    def _order(self, o):
        """ข้อ ๑๖ คำสั่ง (แบบที่ ๔)"""
        s = self.spec
        self._title_block([f"คำสั่ง{o.issuer}", f"ที่ {o.doc_no}", f"เรื่อง  {o.subject}"])
        if o.committees:
            self._order_body(o.paragraphs)
            for i, c in enumerate(o.committees, 1):
                self._committee(str(i), c)
            self._order_body([], o.effective)
        else:
            self._order_body(o.paragraphs, o.effective)
        self._issued(f"สั่ง  ณ  วันที่  {o.date}")
        self._issuer_signature(o.signer)

    def _committee(self, no: str, c):
        """คำสั่งแต่งตั้ง: หัวข้อคณะ → ตารางรายชื่อ (ไม่มีเส้น) → หน้าที่และอำนาจ → คณะย่อย"""
        s = self.spec
        self._line(f"{no}. {c.title}", left=s.para_indent_cm, before=s.para_before_pt, keep_next=True)
        for t in split_paragraphs(list(c.paragraphs)):
            self._line(t, align=DIST, first=s.para_indent_cm)
        self._members(no, c.members)
        self._duties(c.duties)
        for j, g in enumerate(c.subgroups, 1):
            sub = f"{no}.{j}"
            self._line(f"{sub} {g.title}", left=s.para_indent_cm, before=s.para_before_pt, keep_next=True)
            self._members(sub, g.members)
            self._duties(g.duties)

    def _members(self, no: str, members: list):
        """ตารางรายชื่อ ๓ คอลัมน์: เลขข้อ | ชื่อ/ตำแหน่ง/สังกัด | บทบาท — แก้ไขต่อใน Word ได้ง่าย"""
        if not members:
            return
        s = self.spec
        widths = (s.para_indent_cm + 1.5, s.text_width_cm - s.para_indent_cm - 1.5 - 4.5, 4.5)
        t = self.d.add_table(rows=0, cols=3)
        t.autofit = False
        for k, m in enumerate(members, 1):
            row = t.add_row()
            trPr = row._tr.get_or_add_trPr()
            trPr.append(OxmlElement("w:cantSplit"))
            lines = [m.name, *(x for x in (m.position, m.affiliation) if x)]
            cells = row.cells
            for cell, w in zip(cells, widths):
                cell.width = Cm(w)
            self._run(cells[0].paragraphs[0], f"{no}.{k}")
            cells[0].paragraphs[0].paragraph_format.left_indent = Cm(s.para_indent_cm - 0.19)
            for i, line in enumerate(lines):
                cp = cells[1].paragraphs[0] if i == 0 else cells[1].add_paragraph()
                self._run(cp, line)
            self._run(cells[2].paragraphs[0], m.role)

    def _duties(self, duties: list[str]):
        if not duties:
            return
        s = self.spec
        self._line("หน้าที่และอำนาจ", left=s.para_indent_cm, before=s.para_before_pt, keep_next=True, underline=True)
        for n, d in enumerate(duties, 1):
            self._line(f"{n})\t{d}", align=DIST, left=s.para_indent_cm + 0.75, hanging=0.75,
                       tabs=[(s.para_indent_cm + 0.75, WD_TAB_ALIGNMENT.LEFT, WD_TAB_LEADER.SPACES)])

    def _regulation(self, r):
        """ข้อ ๑๗ ระเบียบ / ข้อ ๑๘ ข้อบังคับ (แบบที่ ๕–๖)"""
        s = self.spec
        word = "ระเบียบ" if r.doc_type == DocType.REGULATION else "ข้อบังคับ"
        lines = [f"{word}{r.issuer}", f"ว่าด้วย{r.title}"]
        if r.edition:
            lines.append(f"(ฉบับที่ {r.edition})")
        lines.append(f"พ.ศ. {r.year}")
        self._title_block(lines)
        self._body([r.preamble])
        art_before = 0 if s.order_para_first_only else s.para_before_pt
        for i, a in enumerate(r.articles, 1):
            if a.chapter:
                self._line(a.chapter, align=CENTER, bold=s.title_bold, before=s.closing_before_pt, keep_next=True)
            self._line(f"ข้อ {i}  {a.text}", align=DIST, first=s.para_indent_cm, before=art_before)
        self._issued(f"ประกาศ  ณ  วันที่  {r.date}")
        self._issuer_signature(r.signer)

    def _announcement(self, a):
        """ข้อ ๒๐ ประกาศ / แจ้งความ (แบบที่ ๗)"""
        s = self.spec
        self._title_block([f"{a.word}{a.issuer}", f"เรื่อง  {a.subject}"])
        self._body(a.paragraphs)
        self._issued(f"{a.word}  ณ  วันที่  {a.date}")
        self._issuer_signature(a.signer)

    def _statement(self, st):
        """ข้อ ๒๑ แถลงการณ์ / ข้อ ๒๒ ข่าว (แบบที่ ๘–๙) — ไม่มีลายมือชื่อ"""
        s = self.spec
        word = "แถลงการณ์" if st.doc_type == DocType.STATEMENT else "ข่าว"
        lines = [f"{word}{st.issuer}", f"เรื่อง  {st.subject}"]
        if st.number:
            lines.append(f"ฉบับที่ {st.number}")
        self._title_block(lines, emblem=s.news_emblem or st.doc_type == DocType.STATEMENT)
        self._body(st.paragraphs, keep_tail=True)
        if s.statement_issuer_cm is not None:
            # ๓ Enter → ชื่อหน่วยงานเริ่มห่างแนวกึ่งกลาง ๒ ซม. → วัน เดือน ปี กึ่งกลางใต้ชื่อ
            for _ in range(s.statement_blank_lines):
                self._par(keep_next=True)
            self._line(st.issuer, left=s.statement_issuer_cm, keep_next=True)
            self._centered_at(st.date, s.statement_issuer_cm + self._text_cm(st.issuer) / 2)
        else:
            self._par(keep_next=True)
            self._centered(st.issuer, before=s.closing_before_pt, keep_next=True)
            self._centered(st.date)

    def _certificate(self, c):
        """ข้อ ๒๔ หนังสือรับรอง (แบบที่ ๑๐)"""
        s = self.spec
        self._garuda_header()
        agency = c.agency or ["[ระบุส่วนราชการเจ้าของหนังสือ]"]
        self._line(f"{s.cert_no_label} {c.doc_no}\t{agency[0]}", before=s.block_before_pt,
                   tabs=[(s.agency_tab_cm, WD_TAB_ALIGNMENT.LEFT, WD_TAB_LEADER.SPACES)])
        for line in agency[1:]:
            self._line(line, left=s.agency_tab_cm)
        if s.cert_title:
            self._line("หนังสือรับรอง", align=CENTER, bold=True, before=s.closing_before_pt)
            self._body(c.body, first_before=s.closing_before_pt)
        else:
            self._par()                                               # แบบที่ ๑๐: ๒ Enter → ข้อความ
            self._body(c.body, first_before=0)
        self._issued(f"ให้ไว้  ณ  วันที่  {c.date}")
        self._issuer_signature(c.signer)
        if c.photo_holder:
            # รูปถ่าย ๔x๖ ซม. ประทับตราที่มุมล่างขวาคาบรูป + ลายมือชื่อผู้ได้รับการรับรองใต้รูป
            self._line("[ติดรูปถ่าย ๔ x ๖ ซม.]", before=s.closing_before_pt)
            self._par()
            self._line(f"({c.photo_holder})")
            self._line("ลายมือชื่อผู้ได้รับการรับรอง")

    def _minutes(self, m):
        """ข้อ ๒๕ รายงานการประชุม (แบบที่ ๑๑)"""
        s = self.spec
        date = m.date if m.date.startswith("เมื่อ") else f"เมื่อ{m.date}"
        for i, t in enumerate((m.title, m.session, date, f"ณ {m.place}")):
            self._line(t, align=CENTER, bold=s.minutes_title_bold, before=self._top_gap() if i == 0 else 0)
        self._separator()
        ul = s.minutes_label_underline
        lbl = dict(bold=not ul, underline=ul)

        def people(label, names):
            if not names:
                return
            self._line(label, before=s.block_before_pt, keep_next=True, **lbl)
            for i, n in enumerate(names, 1):
                self._line(f"{i}. {n}", left=s.minutes_list_indent_cm)

        people("ผู้มาประชุม", m.attendees)
        people("ผู้ไม่มาประชุม", m.absentees)
        people("ผู้เข้าร่วมประชุม", m.others)
        p = self._par(before=s.closing_before_pt if not ul else s.block_before_pt)
        self._run(p, "เริ่มประชุมเวลา", **lbl)
        self._run(p, f" {m.start_time} น.")
        for i, item in enumerate(m.agenda, 1):
            p = self._par(before=s.closing_before_pt, keep_next=True,
                          tabs=[(3.5, WD_TAB_ALIGNMENT.LEFT, WD_TAB_LEADER.SPACES)])
            self._run(p, f"ระเบียบวาระที่ {i}\t{item.title}", bold=True)
            self._body(item.discussion)
            p = self._par(align=DIST, first=s.para_indent_cm, before=s.para_before_pt)
            self._run(p, "มติที่ประชุม", bold=True)
            self._run(p, f"  {item.resolution}")
        p = self._par(before=s.closing_before_pt if not ul else s.para_before_pt)
        self._run(p, "เลิกประชุมเวลา", **lbl)
        self._run(p, f" {m.end_time} น.")
        if s.minutes_recorder_cm is not None:
            # แบบที่ ๑๑: ๔ Enter → (ลงชื่อ) เริ่มห่างแนวกึ่งกลาง ๒ ซม. → (ชื่อ) → ตำแหน่ง → ผู้จดรายงานการประชุม
            for _ in range(s.signature_blank_lines):
                self._par(keep_next=True)
            x0 = s.minutes_recorder_cm
            sig = "(ลงชื่อ)"
            name = f"({m.recorder.name})"
            c = x0 + max(self._text_cm(sig), self._text_cm(name)) / 2
            self._line(sig, left=x0, keep_next=True)
            self._centered_at(name, c, keep_next=True)
            pos = [x for x in m.recorder.position if x.strip()]
            if not any("ผู้จดรายงาน" in x for x in pos):
                pos.append("ผู้จดรายงานการประชุม")
            for i, line in enumerate(pos):
                self._centered_at(line, c, keep_next=i < len(pos) - 1)
        else:
            self._signature(m.recorder, blank_lines=2)


# ความกว้างข้อความ: อ่าน advance จากไฟล์ฟอนต์ TH Sarabun New ด้วย fontTools (ถ้ามี) ไม่งั้นใช้ค่าเฉลี่ย
_FONT_DIRS = ("~/.fonts", "~/Library/Fonts", "/Library/Fonts", "/usr/share/fonts/truetype/thai",
              "/usr/share/fonts/truetype/tlwg", "C:/Windows/Fonts")
_FONT_CACHE: dict = {}
_UPPER_LOWER = set("\u0e31\u0e34\u0e35\u0e36\u0e37\u0e38\u0e39\u0e3a\u0e47\u0e48\u0e49\u0e4a\u0e4b\u0e4c\u0e4d")


def _font_metrics(bold: bool):
    key = "b" if bold else "m"
    if key not in _FONT_CACHE:
        _FONT_CACHE[key] = None
        try:
            from fontTools.ttLib import TTFont
        except ImportError:
            return None
        name = "THSarabunNew Bold.ttf" if bold else "THSarabunNew.ttf"
        for d in _FONT_DIRS:
            path = Path(d).expanduser() / name
            if path.is_file():
                try:
                    f = TTFont(str(path))
                    _FONT_CACHE[key] = (f.getBestCmap(), f["hmtx"].metrics, f["head"].unitsPerEm)
                    break
                except Exception:  # pragma: no cover - ไฟล์ฟอนต์เสีย
                    pass
    return _FONT_CACHE[key]


def text_width_cm(text: str, pt: float = 16, bold: bool = False) -> float:
    """ความกว้างข้อความโดยประมาณ (ใช้จัด "กึ่งกลางกับคำลงท้าย/ชื่อ" และบีบบรรทัดหน่วยงานของบันทึกข้อความ)"""
    m = _font_metrics(bold)
    if m is None and bold:                             # ไม่มีไฟล์ตัวหนา: ใช้ตัวธรรมดา + ๕%
        return text_width_cm(text, pt) * 1.05
    if m:
        cmap, hm, upm = m
        units = sum(hm[cmap[ord(ch)]][0] for ch in text if ord(ch) in cmap)
        return units / upm * pt * 2.54 / 72
    base = sum(1 for ch in text if not unicodedata.combining(ch) and ch not in _UPPER_LOWER)
    return base * 0.31 * pt / 16


def _inline_to_anchor(run, x_emu: int = 0, y_emu: int = 0, center: bool = False):
    """แปลงรูป inline เป็นรูปลอย ตำแหน่งแนวนอนเทียบขอบเนื้อหา แนวตั้งเทียบย่อหน้า"""
    from lxml import etree

    inline = run._r.find(".//" + qn("wp:inline"))
    ext = inline.find(qn("wp:extent"))
    docpr = inline.find(qn("wp:docPr"))
    graphic = inline.find(qn("a:graphic"))
    nsdecl = ('xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
              'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"')
    pos_h = ('<wp:positionH relativeFrom="margin"><wp:align>center</wp:align></wp:positionH>' if center else
             f'<wp:positionH relativeFrom="margin"><wp:posOffset>{x_emu}</wp:posOffset></wp:positionH>')
    anchor = etree.fromstring(
        f'<wp:anchor {nsdecl} distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" '
        f'behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
        f'<wp:simplePos x="0" y="0"/>{pos_h}'
        f'<wp:positionV relativeFrom="paragraph"><wp:posOffset>{y_emu}</wp:posOffset></wp:positionV>'
        f'<wp:extent cx="{ext.get("cx")}" cy="{ext.get("cy")}"/>'
        f'<wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>'
        f'<wp:docPr id="{docpr.get("id")}" name="{docpr.get("name")}"/>'
        f'<wp:cNvGraphicFramePr/></wp:anchor>')
    anchor.append(graphic)
    inline.getparent().replace(inline, anchor)
