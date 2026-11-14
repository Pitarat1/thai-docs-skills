# แปลง docx เป็น PDF เพื่อดูผล

ดูผลด้วยตาทุกครั้งก่อนส่ง: ตารางล้น หัวข้อค้างท้ายหน้า ช่องลงชื่อแยกหน้า หัวตารางซ้ำ ข้อความแตกบรรทัด
คำขาดกลางคำ ตัวอักษรเพี้ยน เลือกวิธีตามเครื่องที่มี (ผลที่ตรงกับสิ่งที่ผู้ใช้จะเห็นที่สุดคือ Word เอง)

## 1. Word for Mac (AppleScript)

Word for Mac อยู่ใน sandbox เปิด/บันทึกได้เฉพาะไฟล์ในโฟลเดอร์ของมันเอง จึงต้องคัดลอกไฟล์เข้าไปก่อน

```bash
WD="$HOME/Library/Containers/com.microsoft.Word/Data/Documents"
cp "ไฟล์_v2.docx" "$WD/preview.docx"
osascript <<EOF
tell application "Microsoft Word"
    open POSIX file "$WD/preview.docx"
    set d to active document
    save as d file name "$WD/preview.pdf" file format format PDF
    close d saving no
end tell
EOF
mv "$WD/preview.pdf" ./ไฟล์_v2.pdf && rm "$WD/preview.docx"
```

- แปลงทีละไฟล์ และลบไฟล์ทดสอบในโฟลเดอร์ของ Word เมื่อเสร็จ
- ครั้งแรก macOS จะถามสิทธิ์ให้โปรแกรมที่รันคำสั่งควบคุม Word ให้ผู้ใช้กดอนุญาตเอง
- ถ้าผู้ใช้เปิด Word ทำงานอื่นอยู่ คำสั่งนี้ปิดเฉพาะเอกสารที่เปิดเอง (`close d saving no`)
- เป็นวิธีเดียวที่ตรวจผลการตัดคำ (`wrap`) แล้ว

## 2. LibreOffice

```bash
soffice --headless --convert-to pdf --outdir ตรวจ ไฟล์_v2.docx
pdftoppm -r 60 -png ตรวจ/ไฟล์_v2.pdf ตรวจ/หน้า        # ภาพทีละหน้า (poppler)
pdftotext ตรวจ/ไฟล์_v2.pdf - | grep -c "(ต่อ)"         # นับคำสำคัญ เช่น หัวตาราง "(ต่อ)"
```

- บน macOS คำสั่งอยู่ที่ `/Applications/LibreOffice.app/Contents/MacOS/soffice`
- จำนวนหน้าและจุดขึ้นบรรทัดเป็นค่าประมาณ ใน Word อาจต่างเล็กน้อย; ผลของ `wrap` ใน LibreOffice ยังไม่ได้ตรวจ
- ถ้าเครื่องไม่มีฟอนต์ TH Sarabun New ข้อความจะถูกแทนด้วยฟอนต์อื่น ขนาดและการขึ้นบรรทัดจะผิด → ใช้วิธีที่ 3
  (ตรวจด้วย `fc-match "TH Sarabun New"`)
- ต่อภาพหลายหน้าเป็นแผ่นเดียว (contact sheet) แล้วดูทีเดียวจะเร็วกว่าเปิดทีละหน้า

## 3. ฟอนต์แทน (เครื่องที่ไม่มี TH Sarabun New เช่น Linux/คอนเทนเนอร์)

สร้างฟอนต์แทนจาก Sarabun (Google Fonts, OFL) ซึ่ง metric ใกล้กัน แล้วตั้งชื่อเป็น "TH Sarabun New"
เพื่อให้ LibreOffice ใช้แทน ใช้ดูผลเท่านั้น ไม่ฝังลงในเอกสาร

```bash
mkdir -p ~/fontdl && cd ~/fontdl
npm pack @fontsource/sarabun@5.0.13 && tar xzf fontsource-sarabun-5.0.13.tgz
pip install fonttools brotli
python3 - <<'EOF'
from fontTools.ttLib import TTFont
from fontTools.merge import Merger
import os
os.makedirs(os.path.expanduser('~/.fonts'), exist_ok=True)
for w, style in (('400', 'Regular'), ('700', 'Bold')):
    paths = []
    for s in ('thai', 'latin'):
        f = TTFont(f'package/files/sarabun-{s}-{w}-normal.woff2'); f.flavor = None
        f.save(f'{s}{w}.ttf'); paths.append(f'{s}{w}.ttf')
    m = Merger().merge(paths)
    for rec in m['name'].names:
        if rec.nameID in (1, 16): rec.string = 'TH Sarabun New'
        if rec.nameID in (2, 17): rec.string = style
        if rec.nameID == 4: rec.string = 'TH Sarabun New ' + style
        if rec.nameID == 6: rec.string = 'THSarabunNew-' + style
    m['head'].unitsPerEm = 1390   # ย่อให้ขนาดใกล้ TH Sarabun New (ค่าประมาณ)
    m.save(os.path.expanduser(f'~/.fonts/THSarabunNew-{style}.ttf'))
EOF
fc-cache -f && fc-match "TH Sarabun New"
```

จากนั้นแปลงด้วยวิธีที่ 2 ตามปกติ ถ้าเอกสารใช้ฟอนต์อื่น (เช่น `fix --font "TH SarabunPSK"`)
ให้เปลี่ยนชื่อใน nameID 1/4/6/16 และชื่อไฟล์ให้ตรงกับฟอนต์นั้น
