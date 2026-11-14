"""โครงข้อมูลหนังสือราชการ ๑๑ ชนิด ตามระเบียบสำนักนายกรัฐมนตรีว่าด้วยงานสารบรรณ พ.ศ. ๒๕๒๖
และที่แก้ไขเพิ่มเติม

ลำดับ field ในแต่ละ model = ลำดับส่วนประกอบตามข้อของระเบียบ (อ้างไว้ใน docstring)
ผู้ใช้หรือ AI เติมเฉพาะ field; ตัวจัดหน้าจัดรูปแบบให้

โมดูลนี้เป็นโมดูลเดียวที่นำเข้า pydantic
"""
from __future__ import annotations

import json
import re
from enum import Enum
from pathlib import Path
from typing import Literal, Optional, Union

from pydantic import BaseModel, ConfigDict, Field, ValidationError, model_validator


class DocType(str, Enum):
    MEMO = "memo"                        # หนังสือภายใน / บันทึกข้อความ  (ข้อ ๑๒)
    EXTERNAL_LETTER = "external_letter"  # หนังสือภายนอก               (ข้อ ๑๑)
    STAMP_LETTER = "stamp_letter"        # หนังสือประทับตรา             (ข้อ ๑๓–๑๔)
    ORDER = "order"                      # คำสั่ง                      (ข้อ ๑๖)
    REGULATION = "regulation"            # ระเบียบ                     (ข้อ ๑๗)
    BYLAW = "bylaw"                      # ข้อบังคับ                    (ข้อ ๑๘)
    ANNOUNCEMENT = "announcement"        # ประกาศ / แจ้งความ            (ข้อ ๒๐)
    STATEMENT = "statement"              # แถลงการณ์                   (ข้อ ๒๑)
    NEWS = "news"                        # ข่าว                        (ข้อ ๒๒)
    CERTIFICATE = "certificate"          # หนังสือรับรอง                (ข้อ ๒๔)
    MINUTES = "minutes"                  # รายงานการประชุม              (ข้อ ๒๕)


DOC_TYPE_TH = {
    DocType.MEMO: "บันทึกข้อความ",
    DocType.EXTERNAL_LETTER: "หนังสือภายนอก",
    DocType.STAMP_LETTER: "หนังสือประทับตรา",
    DocType.ORDER: "คำสั่ง",
    DocType.REGULATION: "ระเบียบ",
    DocType.BYLAW: "ข้อบังคับ",
    DocType.ANNOUNCEMENT: "ประกาศ",
    DocType.STATEMENT: "แถลงการณ์",
    DocType.NEWS: "ข่าว",
    DocType.CERTIFICATE: "หนังสือรับรอง",
    DocType.MINUTES: "รายงานการประชุม",
}

Speed = Literal["ด่วนที่สุด", "ด่วนมาก", "ด่วน"]          # ข้อ ๒๘
Secrecy = Literal["ลับที่สุด", "ลับมาก", "ลับ"]


class _Base(BaseModel):
    model_config = ConfigDict(extra="forbid")


class Signer(_Base):
    """ข้อ ๑๑.๑๐–๑๑.๑๑: ลงลายมือชื่อ + พิมพ์ชื่อเต็มในวงเล็บใต้ลายมือชื่อ + ตำแหน่ง"""
    name: str = Field("[ระบุชื่อผู้ลงนาม]", description="ชื่อเต็ม ไม่ต้องใส่วงเล็บ ไม่ต้องใส่คำนำหน้าก็ได้")
    position: list[str] = Field(default_factory=lambda: ["[ระบุตำแหน่ง]"],
                                description="ตำแหน่ง แยกบรรทัด")


class Body(_Base):
    """ข้อความ: ภาคเหตุ → ภาคความประสงค์ → ภาคสรุป  (หลายประเด็นให้แยกเป็นข้อ)"""
    cause: list[str] = Field(..., description="ภาคเหตุ ขึ้นต้น 'ด้วย'/'ตาม'/'เนื่องด้วย'")
    purpose: list[str] = Field(..., description="ภาคความประสงค์")
    conclusion: str = Field(..., description="ภาคสรุป เช่น 'จึงเรียนมาเพื่อโปรดพิจารณา'")

    def paragraphs(self) -> list[str]:
        return [*self.cause, *self.purpose, self.conclusion]


class _Classified(_Base):
    speed: Optional[Speed] = Field(None, description="ชั้นความเร็ว (ข้อ ๒๘)")
    secrecy: Optional[Secrecy] = Field(None, description="ชั้นความลับ ตามระเบียบว่าด้วยการรักษาความลับของทางราชการ")


# ------------------------------------------------------------------- หนังสือโต้ตอบ
class ExternalLetterDoc(_Classified):
    """ข้อ ๑๑ หนังสือภายนอก — ใช้กระดาษตราครุฑ"""
    doc_type: Literal[DocType.EXTERNAL_LETTER] = DocType.EXTERNAL_LETTER
    doc_no: str = Field(..., description="๑๑.๑ ที่: รหัสตัวพยัญชนะ+เลขประจำเจ้าของเรื่อง/เลขทะเบียนหนังสือส่ง เช่น 'ตย [ระบุเลขหน่วยงาน]/[ระบุ]'")
    agency: list[str] = Field(..., description="๑๑.๒ ส่วนราชการเจ้าของหนังสือ + ที่ตั้ง แยกบรรทัด")
    date: str = Field(..., description="๑๑.๓ วัน เดือน ปี: ตัวเลขวันที่ ชื่อเต็มของเดือน ตัวเลขปีพุทธศักราช")
    subject: str = Field(..., description="๑๑.๔ เรื่อง: ย่อที่สุด")
    salutation: str = Field("เรียน", description="๑๑.๕ คำขึ้นต้นตามฐานะผู้รับ (ภาคผนวก ๒)")
    to: str = Field(..., description="ตำแหน่ง/ชื่อผู้รับ ไม่ต้องมีคำขึ้นต้น")
    references: list[str] = Field(default_factory=list, description="๑๑.๖ อ้างถึง: หนังสือที่ติดต่อกันมาครั้งหลังสุด")
    attachments: list[str] = Field(default_factory=list, description="๑๑.๗ สิ่งที่ส่งมาด้วย")
    body: Body                                                      # ๑๑.๘
    closing: str = Field("ขอแสดงความนับถือ", description="๑๑.๙ คำลงท้าย (ภาคผนวก ๒)")
    signer: Signer = Signer()                                        # ๑๑.๑๐–๑๑.๑๑
    owner_unit: str = Field("[ระบุส่วนราชการเจ้าของเรื่อง]", description="๑๑.๑๒ ส่วนราชการเจ้าของเรื่อง")
    phone: str = Field("[ระบุ]", description="๑๑.๑๓ โทร.")
    fax: Optional[str] = None
    email: Optional[str] = Field(None, description="ไปรษณีย์อิเล็กทรอนิกส์ (ระบบสารบรรณอิเล็กทรอนิกส์ ฉบับที่ ๔)")
    copies_to: list[str] = Field(default_factory=list, description="๑๑.๑๔ สำเนาส่ง (ถ้ามี)")


class MemoSection(_Base):
    """หัวข้อบันทึกเสนอเรื่อง (บันทึกเสนอเรื่องแบบมีหัวข้อ): เรื่องเดิม / ข้อเท็จจริง / ข้อกฎหมาย / ข้อพิจารณา / ข้อเสนอ"""
    heading: str
    paragraphs: list[str]


MEMO_SECTION_HEADINGS = ["เรื่องเดิม", "ข้อเท็จจริง", "ข้อกฎหมาย", "ข้อพิจารณา", "ข้อเสนอ"]


class MemoDoc(_Classified):
    """ข้อ ๑๒ หนังสือภายใน — ใช้กระดาษบันทึกข้อความ (อ้างถึง/สิ่งที่ส่งมาด้วยให้เขียนในข้อความ)"""
    doc_type: Literal[DocType.MEMO] = DocType.MEMO
    department: str = Field(..., description="๑๒.๑ ส่วนราชการ: เจ้าของเรื่อง + โทร.")
    doc_no: str = Field(..., description="๑๒.๒ ที่")
    date: str = Field(..., description="๑๒.๓ วันที่")
    subject: str = Field(..., description="๑๒.๔ เรื่อง")
    salutation: str = "เรียน"                                         # ๑๒.๕
    to: str
    body: Optional[Body] = None                                      # ๑๒.๖ (แบบข้อความต่อเนื่อง)
    sections: list[MemoSection] = Field(default_factory=list,
                                        description="บันทึกเสนอเรื่องแบบมีหัวข้อ: ใช้แทน body เมื่อเสนอเรื่องเพื่อพิจารณา")
    signer: Signer = Signer()                                        # ๑๒.๗

    @model_validator(mode="after")
    def _need_body(self):
        if self.body is None and not self.sections:
            raise ValueError("ต้องมี body (แบบข้อความต่อเนื่อง) หรือ sections (บันทึกเสนอเรื่องแบบมีหัวข้อ)")
        return self


class ForwardedLetter(_Base):
    """หนังสือต้นเรื่องที่ส่งสำเนาต่อ (หนังสือประทับตราแจ้งเพื่อทราบ)"""
    agency: str = Field(..., description="ส่วนราชการเจ้าของหนังสือต้นเรื่อง เช่น 'กระทรวงตัวอย่าง'")
    speed: Optional[str] = Field(None, description="ชั้นความเร็วของหนังสือต้นเรื่อง (ถ้ามี) เช่น 'ด่วนที่สุด'")
    doc_no: str = Field("[ระบุเลขที่หนังสือ]", description="เลขที่หนังสือต้นเรื่อง เช่น 'ตย ๐๒๐๑/ว ๑๒๓'")
    date: str = Field("[ระบุวันที่หนังสือ]", description="วันที่ของหนังสือต้นเรื่อง เช่น '๑๐ ตุลาคม ๒๕๖๙'")
    subject: str = Field(..., description="ชื่อเรื่องของหนังสือต้นเรื่อง")


class StampLetterDoc(_Classified):
    """ข้อ ๑๓–๑๔ หนังสือประทับตรา — ใช้แทนการลงชื่อของหัวหน้าส่วนราชการระดับกรมขึ้นไป
    เฉพาะเรื่องไม่สำคัญ (ขอรายละเอียด ส่งสำเนา ตอบรับทราบ เตือนเรื่องที่ค้าง ฯลฯ)
    หัวหน้าส่วนราชการระดับกองหรือผู้ได้รับมอบหมาย ลงชื่อย่อกำกับตรา; ตราประทับด้วยหมึกแดง"""
    doc_type: Literal[DocType.STAMP_LETTER] = DocType.STAMP_LETTER
    doc_no: str = Field(..., description="๑๔.๑ ที่")
    to: str = Field(..., description="๑๔.๒ ถึง: ส่วนราชการ/หน่วยงาน/บุคคล")
    forwarded: list[ForwardedLetter] = Field(default_factory=list,
                                             description="หนังสือต้นเรื่องที่ส่งสำเนาไปเพื่อทราบ — ระบบเขียนข้อความให้เอง")
    paragraphs: list[str] = Field(default_factory=list, description="๑๔.๓ ข้อความ (เว้นว่างได้เมื่อมี forwarded)")
    issuer: str = Field(..., description="๑๔.๔ ส่วนราชการที่ส่งหนังสือออก")
    date: str = Field(..., description="๑๔.๖ วัน เดือน ปี")
    owner_unit: str = Field("[ระบุส่วนราชการเจ้าของเรื่อง]", description="๑๔.๗")
    phone: str = Field("[ระบุ]", description="๑๔.๘ โทร. หรือที่ตั้ง")
    fax: Optional[str] = None
    email: Optional[str] = None

    @model_validator(mode="after")
    def _need_text(self):
        if not self.paragraphs and not self.forwarded:
            raise ValueError("ต้องมี paragraphs หรือ forwarded")
        return self


# ------------------------------------------------------------------- หนังสือสั่งการ
class CommitteeMember(_Base):
    """องค์ประกอบคณะกรรมการ/คณะทำงาน ๑ คน (หรือ ๑ ตำแหน่ง เช่น "หัวหน้าส่วนราชการทุกแห่ง")"""
    name: str = Field(..., description="ชื่อ-สกุล พร้อมคำนำหน้า หรือชื่อตำแหน่งเมื่อแต่งตั้งโดยตำแหน่ง")
    position: Optional[str] = Field(None, description="ตำแหน่ง (เว้นว่างถ้า name เป็นชื่อตำแหน่งอยู่แล้ว)")
    affiliation: Optional[str] = Field(None, description="สังกัด เช่น โรงพยาบาลตัวอย่าง")
    role: str = Field(..., description="บทบาท เช่น ประธาน / รองประธาน / กรรมการ / คณะทำงาน / กรรมการและเลขานุการ")


class CommitteeGroup(_Base):
    """คณะย่อยภายใต้คณะหลัก (เลข ๒.๑ ๒.๒ ...)"""
    title: str = Field(..., description="ชื่อคณะย่อย ไม่ต้องใส่เลขข้อ")
    members: list[CommitteeMember] = Field(default_factory=list)
    duties: list[str] = Field(default_factory=list, description="หน้าที่และอำนาจ แยกข้อ ไม่ต้องใส่เลขข้อ")


class Committee(_Base):
    """คณะกรรมการ/คณะทำงาน (เลข ๑. ๒. ...) — มีสมาชิกเอง หรือแบ่งเป็นคณะย่อย หรือเป็นข้อความสั่งการ"""
    title: str = Field(..., description="ชื่อคณะ หรือหัวข้อ ไม่ต้องใส่เลขข้อ")
    members: list[CommitteeMember] = Field(default_factory=list)
    duties: list[str] = Field(default_factory=list, description="หน้าที่และอำนาจ แยกข้อ ไม่ต้องใส่เลขข้อ")
    subgroups: list[CommitteeGroup] = Field(default_factory=list)
    paragraphs: list[str] = Field(default_factory=list, description="ข้อความสั่งการของข้อนี้ (ถ้าไม่ใช่รายชื่อ)")


class OrderDoc(_Base):
    """ข้อ ๑๖ คำสั่ง — ใช้กระดาษตราครุฑ"""
    doc_type: Literal[DocType.ORDER] = DocType.ORDER
    issuer: str = Field(..., description="๑๖.๑ ส่วนราชการหรือตำแหน่งที่ออกคำสั่ง")
    doc_no: str = Field(..., description="๑๖.๒ ที่: เลขที่ (เริ่ม ๑ ทุกปีปฏิทิน)/ปีพุทธศักราช เช่น '๑๒๓/๒๕๖๙'")
    subject: str = Field(..., description="๑๖.๓ เรื่อง")
    paragraphs: list[str] = Field(..., description="๑๖.๔ ข้อความ: เหตุ + อำนาจ (ถ้ามี) + สิ่งที่สั่ง")
    committees: list[Committee] = Field(default_factory=list,
                                        description="คำสั่งแต่งตั้ง: คณะกรรมการ/คณะทำงาน ตามลำดับข้อ ๑. ๒. ...")
    effective: str = Field("ทั้งนี้ ตั้งแต่บัดนี้เป็นต้นไป", description="๑๖.๔ วันใช้บังคับ")
    date: str = Field(..., description="๑๖.๕ สั่ง ณ วันที่: '๒๑ ตุลาคม พ.ศ. ๒๕๖๙'")
    signer: Signer = Signer()                                        # ๑๖.๖–๑๖.๗


class Article(_Base):
    text: str = Field(..., description="ข้อความของข้อ ไม่ต้องพิมพ์ 'ข้อ n' ระบบใส่ให้")
    chapter: Optional[str] = Field(None, description="ชื่อหมวด ถ้าข้อนี้เริ่มหมวดใหม่ เช่น 'หมวด ๑ บททั่วไป'")


class RegulationDoc(_Base):
    """ข้อ ๑๗ ระเบียบ / ข้อ ๑๘ ข้อบังคับ — ใช้กระดาษตราครุฑ
    ข้อ ๑ ชื่อ, ข้อ ๒ วันใช้บังคับ, ข้อสุดท้าย ผู้รักษาการ"""
    doc_type: Literal[DocType.REGULATION, DocType.BYLAW] = DocType.REGULATION
    issuer: str = Field(..., description="๑๗.๑ ชื่อส่วนราชการที่ออก")
    title: str = Field(..., description="๑๗.๒ ว่าด้วย ... (ไม่ต้องพิมพ์คำว่า 'ว่าด้วย')")
    edition: Optional[str] = Field(None, description="๑๗.๓ ฉบับที่ (ฉบับแรกไม่ต้องระบุ)")
    year: str = Field(..., description="๑๗.๔ พ.ศ.")
    preamble: str = Field(..., description="๑๗.๕ เหตุผลโดยย่อ + อำนาจ เช่น 'โดยที่เป็นการสมควร... อาศัยอำนาจตาม... จึงวางระเบียบไว้ ดังต่อไปนี้'")
    articles: list[Article] = Field(..., description="๑๗.๖ ข้อ ๑ ชื่อ / ข้อ ๒ วันใช้บังคับ / ... / ข้อสุดท้าย ผู้รักษาการ")
    date: str = Field(..., description="๑๗.๗ ประกาศ ณ วันที่")
    signer: Signer = Signer()


class AnnouncementDoc(_Base):
    """ข้อ ๒๐ ประกาศ (หรือ 'แจ้งความ' เมื่อกฎหมายกำหนด) — ใช้กระดาษตราครุฑ"""
    doc_type: Literal[DocType.ANNOUNCEMENT] = DocType.ANNOUNCEMENT
    word: Literal["ประกาศ", "แจ้งความ"] = "ประกาศ"
    issuer: str = Field(..., description="๒๐.๑ ส่วนราชการที่ออกประกาศ")
    subject: str = Field(..., description="๒๐.๒ เรื่อง")
    paragraphs: list[str] = Field(..., description="๒๐.๓ ข้อความ: เหตุผล + สิ่งที่ประกาศ")
    date: str = Field(..., description="๒๐.๔ ประกาศ ณ วันที่")
    signer: Signer = Signer()


# ------------------------------------------------------------------- หนังสือประชาสัมพันธ์
class StatementDoc(_Base):
    """ข้อ ๒๑ แถลงการณ์ / ข้อ ๒๒ ข่าว — ใช้กระดาษตราครุฑ ไม่มีลายมือชื่อ"""
    doc_type: Literal[DocType.STATEMENT, DocType.NEWS] = DocType.STATEMENT
    issuer: str = Field(..., description="๒๑.๑ ส่วนราชการที่ออก")
    subject: str = Field(..., description="๒๑.๒ เรื่อง")
    number: Optional[str] = Field(None, description="๒๑.๓ ฉบับที่ (ถ้ามีหลายฉบับในเรื่องเดียวกัน)")
    paragraphs: list[str] = Field(..., description="๒๑.๔ ข้อความ")
    date: str = Field(..., description="๒๑.๖ วัน เดือน ปี")


# ------------------------------------------------------------------- หนังสือที่เจ้าหน้าที่ทำขึ้น
class CertificateDoc(_Base):
    """ข้อ ๒๔ หนังสือรับรอง — ใช้กระดาษตราครุฑ"""
    doc_type: Literal[DocType.CERTIFICATE] = DocType.CERTIFICATE
    doc_no: str = Field(..., description="๒๔.๑ เลขที่ เช่น '[ระบุ]/๒๕๖๙'")
    agency: list[str] = Field(..., description="๒๔.๒ ส่วนราชการเจ้าของหนังสือ (+ ที่ตั้ง)")
    body: list[str] = Field(..., description="๒๔.๓ ข้อความ ขึ้นต้นว่า 'หนังสือฉบับนี้ให้ไว้เพื่อรับรองว่า' "
                                             "+ ชื่อเต็ม ตำแหน่ง สังกัด + สิ่งที่รับรอง")
    date: str = Field(..., description="๒๔.๔ ให้ไว้ ณ วันที่")
    signer: Signer = Signer()                                        # ๒๔.๕–๒๔.๖
    photo_holder: Optional[str] = Field(None, description="๒๔.๗ ชื่อผู้ได้รับการรับรอง (เรื่องสำคัญ ติดรูป ๔x๖ ซม.)")


class AgendaItem(_Base):
    title: str
    discussion: list[str] = Field(default_factory=list)
    resolution: str = "[ระบุมติที่ประชุม]"


class MinutesDoc(_Base):
    """ข้อ ๒๕ รายงานการประชุม"""
    doc_type: Literal[DocType.MINUTES] = DocType.MINUTES
    title: str = Field(..., description="๒๕.๑ รายงานการประชุม + ชื่อคณะที่ประชุม/ชื่อการประชุม")
    session: str = Field("ครั้งที่ [ระบุ]", description="๒๕.๒")
    date: str = Field("[ระบุวันที่ประชุม]", description="๒๕.๓ เมื่อ")
    place: str = Field("[ระบุสถานที่]", description="๒๕.๔ ณ")
    attendees: list[str] = Field(default_factory=lambda: ["[ระบุรายชื่อผู้มาประชุม]"],
                                 description="๒๕.๕ ผู้มาประชุม (ผู้แทน: ระบุว่ามาแทนผู้ใด)")
    absentees: list[str] = Field(default_factory=list, description="๒๕.๖ ผู้ไม่มาประชุม (พร้อมเหตุผล ถ้ามี)")
    others: list[str] = Field(default_factory=list, description="๒๕.๗ ผู้เข้าร่วมประชุม")
    start_time: str = Field("[ระบุ]", description="๒๕.๘ เวลาเริ่มประชุม เขียนเฉพาะเวลา เช่น '๐๙.๓๐' ไม่ต้องมี 'น.' เพราะระบบต่อให้เอง")
    agenda: list[AgendaItem]                                          # ๒๕.๙
    end_time: str = Field("[ระบุ]", description="๒๕.๑๐ เวลาเลิกประชุม เขียนเฉพาะเวลา เช่น '๑๑.๓๐' ไม่ต้องมี 'น.' เพราะระบบต่อให้เอง")
    recorder: Signer = Signer(name="[ระบุผู้จดรายงานการประชุม]", position=["ผู้จดรายงานการประชุม"])  # ๒๕.๑๑


AnyDoc = Union[MemoDoc, ExternalLetterDoc, StampLetterDoc, OrderDoc, RegulationDoc,
               AnnouncementDoc, StatementDoc, CertificateDoc, MinutesDoc]

MODEL_BY_TYPE: dict[DocType, type[_Base]] = {
    DocType.MEMO: MemoDoc,
    DocType.EXTERNAL_LETTER: ExternalLetterDoc,
    DocType.STAMP_LETTER: StampLetterDoc,
    DocType.ORDER: OrderDoc,
    DocType.REGULATION: RegulationDoc,
    DocType.BYLAW: RegulationDoc,
    DocType.ANNOUNCEMENT: AnnouncementDoc,
    DocType.STATEMENT: StatementDoc,
    DocType.NEWS: StatementDoc,
    DocType.CERTIFICATE: CertificateDoc,
    DocType.MINUTES: MinutesDoc,
}


# ข้อของระเบียบสำนักนายกรัฐมนตรีว่าด้วยงานสารบรรณ ที่กำหนดชนิดหนังสือนั้น
DOC_TYPE_REF: dict[DocType, str] = {
    DocType.MEMO: "ข้อ ๑๒",
    DocType.EXTERNAL_LETTER: "ข้อ ๑๑",
    DocType.STAMP_LETTER: "ข้อ ๑๓",
    DocType.ORDER: "ข้อ ๑๖",
    DocType.REGULATION: "ข้อ ๑๗",
    DocType.BYLAW: "ข้อ ๑๘",
    DocType.ANNOUNCEMENT: "ข้อ ๒๐",
    DocType.STATEMENT: "ข้อ ๒๑",
    DocType.NEWS: "ข้อ ๒๒",
    DocType.CERTIFICATE: "ข้อ ๒๔",
    DocType.MINUTES: "ข้อ ๒๕",
}


# ------------------------------------------------------------------- อ่านและตรวจข้อมูล
class SchemaError(ValueError):
    """ข้อมูลไม่ตรงโครง; messages = รายการข้อความไทยรูป '<ตำแหน่ง field>: <ปัญหา>'"""

    def __init__(self, messages: list[str]):
        self.messages = list(messages)
        super().__init__("\n".join(self.messages))


_TYPE_TH = {
    "string_type": "ต้องเป็นข้อความ",
    "list_type": "ต้องเป็นรายการ",
    "tuple_type": "ต้องเป็นรายการ",
    "dict_type": "ต้องเป็นกลุ่มข้อมูล (วัตถุ)",
    "model_type": "ต้องเป็นกลุ่มข้อมูล (วัตถุ)",
    "model_attributes_type": "ต้องเป็นกลุ่มข้อมูล (วัตถุ)",
    "int_type": "ต้องเป็นจำนวนเต็ม",
    "int_parsing": "ต้องเป็นจำนวนเต็ม",
    "bool_type": "ต้องเป็นจริง/เท็จ",
    "bool_parsing": "ต้องเป็นจริง/เท็จ",
    "float_type": "ต้องเป็นตัวเลข",
}


def _loc(loc: tuple) -> str:
    return ".".join(str(x) for x in loc) if loc else "(ทั้งเอกสาร)"


def _convert(err: dict) -> str:
    kind = err.get("type", "")
    where = _loc(err.get("loc", ()))
    if kind == "missing":
        return f"{where}: ต้องระบุ"
    if kind == "extra_forbidden":
        return f"{where}: ไม่รู้จัก field นี้"
    if kind in ("literal_error", "enum"):
        expected = (err.get("ctx") or {}).get("expected", "")
        values = re.findall(r"'([^']+)'", expected)
        return f"{where}: ต้องเป็นหนึ่งใน {', '.join(values) if values else expected}"
    if kind == "value_error":
        inner = (err.get("ctx") or {}).get("error")
        return f"{where}: {inner if inner is not None else err.get('msg', '')}"
    if kind in _TYPE_TH:
        return f"{where}: {_TYPE_TH[kind]}"
    return f"{where}: {err.get('msg', 'ค่าไม่ถูกต้อง')}"


def parse_doc(data: dict) -> AnyDoc:
    """ตรวจ dict ตามชนิด doc_type แล้วคืนโมเดล; ผิดโครงโยน SchemaError"""
    types = ", ".join(t.value for t in DocType)
    if not isinstance(data, dict):
        raise SchemaError(["(ทั้งเอกสาร): ต้องเป็นกลุ่มข้อมูล (วัตถุ) ที่มี field doc_type"])
    raw = data.get("doc_type")
    if raw is None:
        raise SchemaError([f"doc_type: ต้องระบุ (ชนิดที่รับ: {types})"])
    try:
        dt = DocType(raw)
    except ValueError:
        raise SchemaError([f"doc_type: ต้องเป็นหนึ่งใน {types} (ได้รับ {raw!r})"]) from None
    try:
        return MODEL_BY_TYPE[dt].model_validate(data)
    except ValidationError as e:
        raise SchemaError([_convert(x) for x in e.errors()]) from None


def load_example(doc_type: DocType) -> dict:
    """อ่านไฟล์ตัวอย่างของชนิดนั้น (examples/<ค่า doc_type>.json ข้างโมดูลนี้)"""
    path = Path(__file__).parent / "examples" / f"{DocType(doc_type).value}.json"
    return json.loads(path.read_text(encoding="utf-8"))
