#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""สร้าง report_latest.js จากไฟล์ Report_sale_YYYYMMDD.xls ที่วันที่ใหม่สุดในโฟลเดอร์แอป

แอป (index.html) อ่านไฟล์ .xls ในโฟลเดอร์เองไม่ได้ เพราะเบราว์เซอร์ไม่อนุญาตให้หน้าเว็บอ่านไฟล์ข้างเคียง
สคริปต์นี้จึงห่อเนื้อหาไฟล์รายงานล่าสุดเป็นไฟล์ .js ที่แอปโหลดได้ แล้วแอปจะนำเข้าเองตอนเปิด

วิธีใช้ (รันทุกครั้งหลังวางไฟล์ Report_sale ใหม่):
    python3 scripts/build_report_latest.py            # สร้างไฟล์ แล้วเปิดแอปในเครื่องให้นำเข้าและซิงก์ขึ้น Sheet
    python3 scripts/build_report_latest.py --no-open  # สร้างไฟล์อย่างเดียว ไม่เปิดเบราว์เซอร์

สำคัญ: การนำเข้าเกิดเฉพาะเมื่อเปิดแอปจากไฟล์ในเครื่อง (index.html ในโฟลเดอร์นี้) เท่านั้น
แอปที่เปิดจากเว็บ github.io ไม่มีไฟล์รายงาน (ห้ามขึ้น GitHub) จะได้ข้อมูลผ่านการซิงก์ Google Sheet หลังจากนั้น

ไม่ต้องติดตั้งอะไรเพิ่ม · ไม่แก้และไม่ลบไฟล์ Report_sale · report_latest.js มีข้อมูลลูกค้า ห้ามขึ้น GitHub
"""
import json
import re
import subprocess
import sys
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
OUTPUT = ROOT / "report_latest.js"
NAME_RE = re.compile(r"^Report_sale_(\d{4})(\d{2})(\d{2})(?:[-_](\d+))?\.xls$", re.I)


def newest_report():
    found = []
    for path in ROOT.glob("Report_sale_*.xls"):
        m = NAME_RE.match(path.name)
        if not m:
            continue
        y, mo, d, seq = m.groups()
        try:
            datetime(int(y), int(mo), int(d))
        except ValueError:
            continue
        found.append(((y + mo + d, int(seq or 0)), path))
    if not found:
        return None
    found.sort()
    return found[-1][1]


def main():
    src = newest_report()
    if src is None:
        print("ไม่พบไฟล์ Report_sale_YYYYMMDD.xls ในโฟลเดอร์ " + str(ROOT), file=sys.stderr)
        return 1
    text = src.read_text(encoding="utf-8", errors="replace")
    count = len(re.findall(r">\s*\d+\.\s*PI-", text))
    if count == 0:
        print("ไฟล์ %s ไม่มีรายการ PI — ไม่ได้สร้าง report_latest.js" % src.name, file=sys.stderr)
        return 1
    m = NAME_RE.match(src.name)
    payload = {
        "name": src.name,
        "date": "%s-%s-%s" % (m.group(1), m.group(2), m.group(3)),
        "builtAt": datetime.now().strftime("%Y-%m-%dT%H:%M:%S"),
        "count": count,
        "html": text,
    }
    body = json.dumps(payload, ensure_ascii=False).replace("</", "<\\/")
    OUTPUT.write_text(
        "// สร้างอัตโนมัติจาก %s โดย scripts/build_report_latest.py — มีข้อมูลลูกค้า ห้ามขึ้น GitHub\n"
        "window.PI_REPORT_FILE = %s;\n" % (src.name, body),
        encoding="utf-8",
    )
    print("สร้าง report_latest.js จาก %s แล้ว (%d ใบ PI)" % (src.name, count))
    if "--no-open" not in sys.argv and sys.platform == "darwin":
        # เปิดแอปจากไฟล์ในเครื่อง: แอปจะนำเข้ารายงานนี้เองแล้วซิงก์ขึ้น Google Sheet (เว็บ/iPhone เห็นหลังดึงข้อมูล)
        subprocess.run(["open", str(ROOT / "index.html")], check=False)
        print("เปิดแอปในเครื่องแล้ว — รอให้ป้ายมุมขวาบนขึ้น \"ซิงก์แล้ว\"")
    return 0


if __name__ == "__main__":
    sys.exit(main())
