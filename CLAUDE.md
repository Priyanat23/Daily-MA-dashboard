# Daily MA Dashboard — คำสั่งมาตรฐานสำหรับ routine อัตโนมัติ

Repo นี้ใช้รัน skill `daily-ma-dashboard-v2` (และ `daily-ma-dashboard-xlsx` คู่กัน) ที่อยู่ใน
`.claude/skills/` เพื่อสร้าง Daily MA Dashboard ให้แผนก PCRC จาก Google Sheet ทุกวัน

## ทุกครั้งที่รัน routine "Daily MA Dashboard" (ไม่ว่าจะรันแบบ scheduled/อัตโนมัติ หรือมีคนสั่งเอง)
ต้องทำสิ่งต่อไปนี้เพิ่มเติมจากขั้นตอนใน SKILL.md เสมอ โดยไม่ต้องถามผู้ใช้ก่อน:

1. **รันแบบไม่ต้องมีคนเปิดคอมพิวเตอร์/มือถือ** — routine นี้ต้องรันจบครบทุกขั้นตอนแบบอัตโนมัติเองได้
   เสมอ (ตาม instructions ใน SKILL.md ข้อ "ทำตาม 5 ขั้นตอนนี้เท่านั้น...ไม่ต้องถามผู้ใช้เพิ่ม") ห้าม
   หยุดรอถามคำถามระหว่างทาง

2. **หลังสร้างไฟล์ `outputs/daily_ma_dashboard.html` เสร็จแล้ว ให้ publish ไฟล์เดียวกันนี้เป็น Artifact
   ด้วย tool `Artifact` เสมอ** (นอกเหนือจาก commit/push เข้า repo ตามปกติ):
   - ใช้ `file_path` เดิมทุกครั้ง (`outputs/daily_ma_dashboard.html`) เพื่อให้ redeploy ทับ URL เดิม
     ไม่สร้างลิงก์ใหม่รายวัน — ผู้ใช้จะได้ลิงก์เดียวที่เปิดดูข้อมูลล่าสุดได้เสมอ
   - Artifact จะ private โดย default — ผู้ใช้จะเป็นคนกด "share/public" เองภายหลังเพื่อส่งต่อทาง email
     (ไม่ต้อง publish หรือเปิด public ให้เองในระหว่าง routine)
   - ใส่ `favicon` เป็นอิโมจิเดิมทุกครั้ง (เช่น 📊) ไม่เปลี่ยนไปมา
   - ตั้ง `title` ใน `<title>` ให้คงที่ (เช่น "Daily MA Dashboard — PCRC") ไม่ใส่วันที่ในชื่อ เพื่อไม่ให้
     ชื่อเปลี่ยนทุกวันใน gallery
   - ก่อน publish ต้อง Write ไฟล์ HTML เนื้อหาตาม `body` skeleton ที่ Artifact tool ต้องการ (ไม่มี
     `<!DOCTYPE>`/`<html>`/`<head>`/`<body>` ครอบซ้ำ) — ใช้ path เดียวกับที่บันทึกลง
     `outputs/daily_ma_dashboard.html` ได้เลยถ้าไฟล์เข้ากันได้ ไม่ต้องสร้างไฟล์คู่ขนาน

3. **สรุปผลท้าย routine ต้องแนบลิงก์ Artifact ของ HTML dashboard เสมอ** (นอกเหนือจากตัวเลข %MA/Revenue
   และชื่อผู้ถูกมอบหมายงาน) เพื่อให้ผู้ใช้กดเปิด/กด public ได้ทันทีจากมือถือโดยไม่ต้องเปิด repo

ไฟล์ `.xlsx` (`outputs/daily_ma_dashboard.xlsx`) ไม่ต้อง publish เป็น Artifact — commit/push เข้า repo
ตามปกติพอ (Artifact รองรับเฉพาะ HTML/Markdown)
