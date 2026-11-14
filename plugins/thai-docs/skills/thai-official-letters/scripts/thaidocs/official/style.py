"""ค่าการพิมพ์ตามแนวทางการพิมพ์หนังสือราชการภาษาไทยด้วยคอมพิวเตอร์
(หนังสือสำนักนายกรัฐมนตรี ที่ นร ๐๑๐๖/ว ๒๐๑๙ ลงวันที่ ๓๐ พฤศจิกายน ๒๕๕๓ + คำอธิบายการพิมพ์ + แผ่นแบบ/ตัวอย่าง
 หนังสือภายนอก ภายใน ประทับตรา — ดาวน์โหลดต้นฉบับได้จากเว็บไซต์สำนักนายกรัฐมนตรี opm.go.th/th/legal/content/37674)
และระเบียบสำนักนายกรัฐมนตรีว่าด้วยงานสารบรรณ พ.ศ. ๒๕๒๖ และที่แก้ไขเพิ่มเติม

ทุกค่าที่ renderer ใช้อยู่ที่นี่ที่เดียว (ไม่พึ่ง pydantic/Pillow/fontTools) — แต่ละค่าระบุที่มา:
  [ว2019]   แนวทางตาม นร ๐๑๐๖/ว ๒๐๑๙
  [ระเบียบ] ระเบียบงานสารบรรณ (เลขข้อ)
  [แผ่นแบบ] วัดจากไฟล์แผ่นแบบ (.doc) ทางการของสำนักนายกรัฐมนตรีที่แนบกับ ว ๒๐๑๙
  [ปฏิบัติ] ค่าที่ใช้กันทั่วไปในคู่มือหน่วยงาน (แนวทางไม่ได้กำหนดตัวเลข) — ปรับได้ตามคู่มือหน่วยงาน
"""
from dataclasses import dataclass
from typing import Optional


# ฟอนต์แห่งชาติ ๑๓ ฟอนต์ (SIPA/กรมทรัพย์สินทางปัญญา) ที่ ว ๒๐๑๙ ให้ใช้แทนฟอนต์เดิม
NATIONAL_FONTS = {
    "TH SarabunPSK", "TH Sarabun New", "TH Charmonman", "TH Krub", "TH Kodchasal",
    "TH K2D July8", "TH Mali Grade6", "TH Chakra Petch", "TH Baijam", "TH KoHo",
    "TH Charm of AU", "TH Srisakdi", "TH Niramit AS", "TH Fah kwang",
    "TH SarabunIT๙",   # TH SarabunPSK ฉบับแสดงเลขอารบิกเป็นเลขไทย
}


@dataclass(frozen=True)
class LayoutSpec:
    font: str = "TH Sarabun New"          # [ว2019] "รูปแบบตัวพิมพ์ไทยสารบรรณ (TH Sarabun PSK)"
    body_pt: float = 16                  # [ว2019] "ขนาด ๑๖ พอยท์" (หนังสือนำ ว ๒๐๑๙)
    memo_title_pt: float = 29            # [ว2019] "บันทึกข้อความ" ตัวหนา ๒๙ พอยท์
    memo_title_exact_pt: float = 35      # [ว2019] ระยะบรรทัด Exactly ๓๕ พอยท์
    memo_label_pt: float = 20            # [ว2019] ส่วนราชการ / ที่ / วันที่ / เรื่อง ตัวหนา ๒๐ พอยท์
    speed_pt: float = 32                 # [ระเบียบ ข้อ ๒๘] ชั้นความเร็ว ตัวหนา ไม่เล็กกว่า ๓๒ พอยท์ สีแดง
    speed_color: str = "FF0000"

    page_w_cm: float = 21.0              # [ปฏิบัติ] A4
    page_h_cm: float = 29.7
    margin_left_cm: float = 3.0          # [ว2019] ขอบซ้าย ๓ ซม.
    margin_right_cm: float = 2.0         # [ว2019] ขอบขวา ๒ ซม. (ไม้บรรทัด ๐–๑๖ ซม.)
    margin_top_cm: float = 1.5           # [ว2019] ครุฑห่างขอบบนประมาณ ๑.๕ ซม. (แผ่นแบบตั้งไว้ ๑.๖)
    margin_bottom_cm: float = 2.0        # [ปฏิบัติ]
    header_distance_cm: float = 0.8

    garuda_external_cm: float = 3.0      # [ว2019] กระดาษตราครุฑ สูง ๓ ซม.
    garuda_memo_cm: float = 1.5          # [ว2019] กระดาษบันทึกข้อความ สูง ๑.๕ ซม.

    para_indent_cm: float = 2.5          # [ว2019] ย่อหน้าภาคเหตุ/ความประสงค์/สรุป ๒.๕ ซม.
    para_before_pt: float = 6            # [ว2019] ระหว่างภาค: ระยะบรรทัดปกติ + Before ๖ พอยท์
    block_before_pt: float = 6           # [ว2019] เรื่อง/คำขึ้นต้น/อ้างถึง/สิ่งที่ส่งมาด้วย: + Before ๖ พอยท์
    closing_before_pt: float = 12        # [ว2019] คำลงท้าย: + Before ๑๒ พอยท์
    signature_blank_lines: int = 3       # [ว2019] ชื่อเต็มผู้ลงนาม: เว้น ๓ บรรทัด (๔ Enter) จากคำลงท้าย/ภาคสรุป
    owner_unit_blank_lines: int = 3      # [ว2019] ส่วนราชการเจ้าของเรื่อง: เว้น ๓ บรรทัด (๔ Enter) จากตำแหน่ง

    agency_tab_cm: float = 10.0          # [แผ่นแบบ] ส่วนราชการเจ้าของหนังสือ/ที่ตั้ง (๗ แท็บ + เว้นวรรค ≈ ๑๐ ซม.)
    center_cm: float = 8.0               # [แผ่นแบบ] แนวกึ่งกลางครุฑ (ครุฑกว้าง ๒.๖ ซม. วางที่ ๖.๗ ซม.)
    date_indent_cm: float = 8.0          # [แผ่นแบบ] วัน เดือน ปี เริ่มที่แนวกึ่งกลางครุฑ (๖ แท็บ + เว้นวรรค)
    signature_center_cm: float = 9.5     # [แผ่นแบบ] คำลงท้าย / (ชื่อเต็ม) / ตำแหน่ง จัดกึ่งกลางที่แท็บกลาง ๙.๕ ซม.
    issued_date_indent_cm: float = 5.0   # [ปฏิบัติ] "สั่ง/ประกาศ/ให้ไว้ ณ วันที่" (หนังสือที่ไม่มีแผ่นแบบ)
    subject_hang_cm: float = 1.5         # [ตัวอย่าง] ข้อความ "เรื่อง" บรรทัดที่ ๒ ตรงกับบรรทัดแรก
    attachment_hang_cm: float = 3.0      # [ตัวอย่าง] ข้อความ "สิ่งที่ส่งมาด้วย" บรรทัดที่ ๒
    salutation_hang_cm: float = 1.27     # [แผ่นแบบ] คำขึ้นต้นบันทึกข้อความ hanging ๗๒๐ twips
    memo_date_tab_cm: float = 8.0        # [แผ่นแบบ] เส้นประช่อง "ที่" ยาวถึง ~๘ ซม. แล้วต่อด้วย "วันที่"
    stamp_to_before_pt: float = 8        # [แผ่นแบบ] หนังสือประทับตรา "ถึง" Before ๘ พอยท์ แล้วเว้น ๑ บรรทัด (๒ Enter)

    thai_numerals: bool = True           # [ปฏิบัติ] เลขไทย (ระเบียบไม่บังคับ — ปิดได้)
    highlight_placeholders: bool = True

    # ---- ค่าที่แตกต่างตามคู่มือหน่วยงาน (ค่าเริ่มต้น = แนวทางสำนักนายกรัฐมนตรี)
    header_blank_paras: int = 2          # [แผ่นแบบ] ย่อหน้าว่างหลังย่อหน้าตรา ก่อนบรรทัด "ที่" (บางคู่มือหน่วยงานใช้ ๑: "ที่" อยู่ระดับกลางตรา)
    emblem_png: Optional[str] = None     # ไฟล์ตรา (None = ใช้ตราจากไฟล์ตั้งค่า/ตัวเลือกบรรทัดคำสั่ง/กรอบแทนตรา)
    emblem_name: str = "ตราครุฑ"
    unit_word: str = "ส่วนราชการ"         # ป้ายบันทึกข้อความ / "ส่วนราชการเจ้าของเรื่อง"
    page_number_fmt: str = "- {} -"
    page_number_pt: Optional[float] = None   # None = body_pt
    word_compat_mode: Optional[int] = None   # None = ค่าของ python-docx (14) / 15 = Word 2013+
    title_bold: bool = True              # ชื่อคำสั่ง/ระเบียบ/ประกาศ ตัวหนา
    separator_cm: float = 0              # 0 = เส้นประ "-" ๒๖ ตัว / >0 = เส้นคั่นยาวเท่านี้ (ซม.) กึ่งกลาง
    closing_mode: str = "tab"            # tab = คำลงท้าย/ชื่อ กึ่งกลางที่แท็บกลาง signature_center_cm
                                         # start_center = คำลงท้ายเริ่มที่แนวกึ่งกลางตรา แล้วชื่อ/ตำแหน่ง กึ่งกลางกับคำลงท้าย
    signer_paren_cm: Optional[float] = None   # วงเล็บเปิด "(" ของชื่อเต็มอยู่ที่แนวนี้ (คำสั่ง/ระเบียบ/ประกาศ/บันทึก/รับรอง)
    issued_blank_line: bool = False      # "สั่ง/ประกาศ/ให้ไว้ ณ วันที่" หลัง ๒ Enter (แทน Before ๑๒ พอยท์)
    order_para_first_only: bool = False  # คำสั่ง/ระเบียบ: Before ๖ เฉพาะย่อหน้าแรก (+ "ทั้งนี้ ตั้งแต่") ที่เหลือ ๑ Enter
    statement_issuer_cm: Optional[float] = None  # แถลงการณ์/ข่าว: ชื่อหน่วยงานเริ่มที่แนวนี้ (None = กึ่งกลาง)
    statement_blank_lines: int = 1
    news_emblem: bool = True             # ข่าว มีตรา
    no_emblem_top_cm: float = 0          # หนังสือที่ไม่มีตรา: บรรทัดแรกห่างขอบกระดาษบน (0 = ตามขอบบน)
    cert_title: bool = True              # หนังสือรับรอง มีบรรทัด "หนังสือรับรอง"
    cert_no_label: str = "เลขที่"
    stamp_blank_after_to: int = 1        # หนังสือประทับตรา: บรรทัดว่างหลัง "ถึง"
    stamp_initials: bool = False         # พิมพ์ "(ลงชื่อย่อกำกับตรา)" ข้างตรา
    minutes_title_bold: bool = True
    minutes_label_underline: bool = False
    minutes_list_indent_cm: float = 1.0
    minutes_recorder_cm: Optional[float] = None  # ผู้จดรายงานการประชุม เริ่มที่แนวนี้ (None = แท็บกลาง)
    # ---- บันทึกข้อความตามแม่แบบหน่วยงาน (None/0 = ใช้ค่าแผ่นแบบ ว ๒๐๑๙)
    memo_margin_top_cm: Optional[float] = None
    memo_title_before_pt: float = 0
    memo_title_after_pt: float = 0
    memo_title_tab_cm: Optional[float] = None      # None = "บันทึกข้อความ" กึ่งกลาง
    memo_title_char_spacing_pt: float = 0
    memo_speed_tab_cm: Optional[float] = None      # ชั้นความเร็วในบรรทัด "บันทึกข้อความ"
    memo_speed_bold: bool = True
    memo_emblem_offset_emu: Optional[tuple] = None # (x, y) ตำแหน่งตราเทียบขอบซ้าย/ย่อหน้าหัว
    memo_phone_tab_cm: Optional[float] = None      # "โทร" แยกไปแท็บนี้ในบรรทัดหน่วยงาน
    memo_signature_mode: str = "center"            # center = กึ่งกลางแท็บ signature_center_cm / paren = "(" ที่แนววันที่
    memo_section_first_cm: Optional[float] = None  # บันทึก ๒/๒: ย่อหน้าข้อความใต้หัวข้อ
    memo_footer_logo_png: Optional[str] = None     # โลโก้ท้ายบันทึกข้อความ (ทุกหน้า)
    memo_footer_logo_w_cm: float = 7.8
    date_be_rules: bool = False          # ตรวจการใช้ "พ.ศ." ในวันที่ตามคู่มือหน่วยงาน
    doc_no_prefix: Optional[str] = None  # รหัสตัวพยัญชนะหน้าเลขที่หนังสือ (เช่น "มท")
    stamp_notify_only: bool = False      # หนังสือประทับตราใช้แจ้งเพื่อทราบเท่านั้น (แนวปฏิบัติของบางหน่วยงาน)
    omit_signature: tuple[str, ...] = ()  # ชนิดหนังสือที่ไม่พิมพ์ชื่อ/ตำแหน่งผู้ลงนาม (ลงนามในระบบสารบรรณอิเล็กทรอนิกส์)

    @property
    def text_width_cm(self) -> float:
        return self.page_w_cm - self.margin_left_cm - self.margin_right_cm


DEFAULT_SPEC = LayoutSpec()

# key ใน "layout" ของไฟล์ตั้งค่าหน่วยงานที่เอกสารผู้ใช้อธิบายไว้ (key อื่นใน LayoutSpec ยังใช้ได้ แต่ไม่รับรองความเสถียร)
DOCUMENTED_LAYOUT_KEYS = (
    "font", "body_pt", "margin_left_cm", "margin_right_cm", "margin_top_cm", "margin_bottom_cm",
    "thai_numerals", "unit_word", "omit_signature", "page_number_fmt", "emblem_png",
    "stamp_notify_only", "highlight_placeholders",
)
