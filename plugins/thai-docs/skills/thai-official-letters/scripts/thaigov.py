"""thaigov: สร้างและตรวจหนังสือราชการ (จุดเรียกใช้ของ skill thai-official-letters)

รัน: python scripts/thaigov.py --help
แพ็กเกจ thaidocs อยู่ในโฟลเดอร์เดียวกับไฟล์นี้ (คัดลอกจาก src/ ด้วย tools/build.py ห้ามแก้ที่นี่)
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from thaidocs.official.cli import main  # noqa: E402

raise SystemExit(main())
