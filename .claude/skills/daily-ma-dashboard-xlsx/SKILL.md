---
name: daily-ma-dashboard-xlsx
description: >
  สร้างไฟล์ Excel (.xlsx) ของ Daily MA Dashboard แผนก PCRC ที่จำลองโครงสร้าง+ข้อมูลของ Daily MA Dashboard
  HTML (skill daily-ma-dashboard-v2) แบบขับเคลื่อนด้วยสูตร (ไม่ hardcode ผลลัพธ์) ครบ 3 ชีต: Dashboard
  (header, period box, KPI cards 5 ใบ รวม Incentive, กราฟ, breakdown แบบการ์ดแนวนอน 8 ช่อง (Shop Online
  แยก 1/2), ตารางรายสัปดาห์, ตารางเทียบสัปดาห์อ้างอิง, การวิเคราะห์ผล, แบนเนอร์มอบหมายงาน), DailyMA
  (ตาราง Target/Actual/Gap รายวัน + กราฟ), RawData (ข้อมูลต้นทาง + config)

  ใช้ทันทีเมื่อผู้ใช้พูดว่า: "ทำ Excel ของ MA dashboard", "excel เหมือน dashboard", "xlsx MA dashboard",
  "ทำไฟล์ Excel daily MA", "แปลง dashboard เป็น excel", "excel เหมือน html 100%" ไม่ว่าจะพูดสั้นแค่ไหน
  ใช้คู่กับ daily-ma-dashboard-v2 เสมอ (อ่านข้อมูลจาก Google Sheet เดียวกัน)

  **Google Sheet เริ่มต้น:** FILE_ID `1q2ZYX9lTMLZoJu_HpLf93dCPHxBY3T2IawPf2fkfuAA`
---

# Daily MA Dashboard (Excel) — Skill สำหรับสร้างไฟล์ .xlsx

## ภาพรวม

Skill นี้สร้างไฟล์ Excel ที่จำลอง Daily MA Dashboard (เดิมเป็น HTML จาก skill `daily-ma-dashboard-v2`)
ให้ตรงกันทั้งโครงสร้างและข้อมูล **แต่ทุกตัวเลขในไฟล์ Excel เป็นสูตรที่คำนวณจากข้อมูลดิบ** ไม่ใช่ค่าที่ฝังตายตัว
เพื่อให้เปิดใน Excel แล้วแก้ไขข้อมูล หรือเปลี่ยนวันที่ปัจจุบัน (`TodayDay`) แล้วทุกอย่างคำนวณใหม่ให้เองอัตโนมัติ

ไฟล์ผลลัพธ์มี 3 ชีต (เรียงจากซ้าย):
1. **Dashboard** — เหมือนหน้า "📊 Dashboard" ของ HTML ทุกส่วน (รวม KPI card "Incentive ประจำวัน" และ
   breakdown การ์ดแนวนอนที่แยก Shop Online 1/2 เหมือน HTML เป๊ะ)
2. **DailyMA** — เหมือนหน้า "📅 DailyMA" ของ HTML ทุกส่วน (8 คอลัมน์สายธุรกิจ รวม Shop Online 1/2 แยกกัน)
3. **RawData** — ข้อมูลต้นทางรายวัน (ไม่มีใน HTML แต่จำเป็นสำหรับ Excel ให้สูตรอ้างอิงได้) + ค่า config
   (`TodayDay`, `RevenueMonth`, `IncentiveTarget`, `IncentiveActual`)

## วิธีใช้ (ทำตามขั้นตอนนี้)

### ขั้นตอนที่ 1 — อ่านข้อมูลจาก Google Sheet
ใช้ `Google Drive:read_file_content` หรือ `download_file_content` กับ FILE_ID ด้านบน (เหมือนขั้นตอนแรกของ
skill `daily-ma-dashboard-v2`) — **อ่านทั้งชีต `2.DailyMA` (ข้อมูลรายวัน) และชีต `1.Dashboard` (สำหรับ
Incentive Target/Actual ที่เซลล์ C3/C4 — ดู field `INCENTIVE_TARGET`/`INCENTIVE_ACTUAL` ด้านล่าง)**
ถ้าผู้ใช้เพิ่งสร้าง Dashboard HTML ในแชทนี้ไปแล้ว ให้ใช้ข้อมูลเดียวกันนั้นได้เลยโดยไม่ต้องอ่านชีตซ้ำ

### ขั้นตอนที่ 2 — คัดลอกสคริปต์และแก้ไขค่า config
คัดลอก `scripts/build_dashboard.py` ไปไว้ที่ workspace (เช่น `/home/claude/build_dashboard.py`)
แล้วแก้ไข**เฉพาะ**ส่วน `=== CONFIG: EDIT HERE ===` ที่อยู่ด้านบนของไฟล์ ให้ตรงกับข้อมูลที่อ่านได้:

| ตัวแปร | ความหมาย |
|---|---|
| `OUTPUT_PATH` | ที่อยู่ไฟล์ผลลัพธ์ |
| `DEPT_NAME` | ชื่อหน่วยงาน (ปกติคือ "PCRC") |
| `WEEK_LABEL` | สัปดาห์ปัจจุบัน เช่น `"27–30"` |
| `DATE_RANGE_LABEL` | ช่วงวันที่เต็ม เช่น `"29 มิถุนายน – 26 กรกฎาคม 2569"` — **ใช้ได้แม้ช่วงคาบเกี่ยว 2 เดือน** |
| `DATE_RANGE_SHORT` | ช่วงวันที่แบบย่อไม่มีปี เช่น `"29 มิ.ย.–26 ก.ค."` ใช้ในคอลัมน์แคบของ comp table เท่านั้น |
| `MONTH_LABEL` | ข้อความอ้างอิงทั่วไป **ไม่ได้ใช้ต่อท้ายวันที่ในสูตรแล้ว** (ดูหัวข้อ "ช่วงข้อมูลคาบเกี่ยว 2 เดือน" ด้านล่าง) — ใส่ไว้เผื่ออ่านง่ายเฉยๆ |
| `COMPARE_LABEL` | ป้ายชื่อสัปดาห์ที่ใช้เทียบ เช่น `"สัปดาห์ 10–13 (มี.ค. 2569)"` |
| `TODAY_DAY` | **ลำดับที่** ของวันนี้ในลิสต์ `DAYS` (นับจาก 1 เสมอ ไม่ใช่วันที่ปฏิทิน) — วันล่าสุดที่มี Actual = `TODAY_DAY - 1` |
| `REVENUE_MONTH` | Revenue สายธุรกิจของงวด (บาท) |
| `INCENTIVE_TARGET` / `INCENTIVE_ACTUAL` | Incentive ประจำวัน (บาท) จากชีต `1.Dashboard` เซลล์ **C3** (Target) / **C4** (Actual) — คนละแหล่งข้อมูลกับ `DAYS` |
| `DAILY_REVENUE` | dict `{dayNum: revenue}` ของวันที่ทราบ revenue เจาะจง (ปกติคือวันล่าสุด/เมื่อวาน) |
| `WEEKS` | list ของ (เลขสัปดาห์, ช่วงวันที่แบบสั้น) ใช้ในตารางสรุปรายสัปดาห์ |
| `DAYS` | **ข้อมูลหลัก** — list ครบทุกวันของงวด รูปแบบ `(dayNum, week, 'D mmm. YYYY', holiday_bool, target_dict, actual_dict)` — `dayNum` เป็นลำดับ 1..N ของงวด (ไม่ใช่วันที่ปฏิทิน), `date` ต้องรวมปี พ.ศ. เสมอ (ดูหัวข้อคาบเกี่ยว 2 เดือนด้านล่าง) — dict มี key: `jrts, ap, ar, rc, so1, so2, pr, sv, pct` (**8 ฟิลด์ ห้ามรวม so1/so2 เป็นค่าเดียว**) |
| `REF_COMPARISON` | ข้อมูล Actual ของสัปดาห์อ้างอิง (label, period, `values` เป็น list 8 ค่าตามลำดับ `jrts, ap, ar, rc, so1, so2, pr, sv`) |
| `ANALYSIS_TEXT` | ข้อความวิเคราะห์ 3 คอลัมน์ (`cause`, `source`, `solution`) แต่ละคอลัมน์เป็น list ของ bullet (string) |

**ห้ามแก้โครงสร้างสคริปต์ส่วนอื่น** (สูตร Excel, กราฟ, conditional formatting, ตำแหน่งแถว) เพราะถูก
ออกแบบให้ตรงกับ HTML แล้ว ดูหัวข้อ "กฎสำคัญที่ต้องรู้" ด้านล่างก่อนแก้ไขส่วนใดที่ไม่แน่ใจ ตำแหน่ง**คอลัมน์**ใน
RawData/Dashboard/DailyMA คำนวณอัตโนมัติจากจำนวน field ใน `FIELDS` อยู่แล้ว (ดูหัวข้อ "สถาปัตยกรรมคอลัมน์
แบบไดนามิก") จึงไม่ต้องคำนวณเลขคอลัมน์เองถ้าไม่ได้เพิ่ม/ลด field

### ขั้นตอนที่ 3 — รันสคริปต์และตรวจสอบ error
```bash
python3 build_dashboard.py
python3 /mnt/skills/public/xlsx/scripts/recalc.py /home/claude/daily_ma_dashboard.xlsx 60
```
ผลลัพธ์ `total_errors` **ต้องเป็น 0 เสมอ**ก่อนส่งให้ผู้ใช้ ถ้ามี error ให้ดู `error_summary` หา cell ที่ผิด
แก้ไข formula ตรงนั้นแล้วรันซ้ำ (สาเหตุที่พบบ่อยที่สุดคือสูตร text ที่ใช้ `&"..."&` ต่อกันแล้วเครื่องหมาย `"` ไม่ครบคู่)

### ขั้นตอนที่ 4 — (แนะนำ) ตรวจสอบภาพก่อนส่ง
แปลงเป็น PDF/PNG เพื่อดูว่าตัวเลขไม่ล้นช่อง (`###`), ป้ายวันที่ไม่ถูกตัด, และกราฟไม่ทับซ้อนกัน:
```bash
soffice --headless --convert-to pdf daily_ma_dashboard.xlsx
```
แล้วแปลงหน้าแรกเป็นรูปด้วย `pdf2image` เพื่อดูคร่าวๆ ถ้าเจอ `###` ให้ขยาย column width ของคอลัมน์นั้น
ถ้าเจอข้อความป้ายกำกับ (เช่น label ของ comp table คอลัมน์ A/B) ถูกตัด ให้ขยาย `dash_widths["A"]`/`["B"]`
แทนการย่อข้อความ (ดู `dash_widths` ใกล้ต้นส่วน Dashboard sheet)

### ขั้นตอนที่ 5 — Present ไฟล์
คัดลอกไฟล์ไปที่ `/mnt/user-data/outputs/` แล้วเรียก `present_files`

## กฎสำคัญที่ต้องรู้ (อย่าแก้ logic เหล่านี้โดยไม่จำเป็น)

### สถาปัตยกรรมคอลัมน์แบบไดนามิก (สำคัญที่สุด — อ่านก่อนแก้อะไรก็ตามที่เกี่ยวกับคอลัมน์)
สคริปต์คำนวณตำแหน่งคอลัมน์ของ RawData/Dashboard/DailyMA จาก **`N_FIELDS = len(FIELDS)`** เสมอ ไม่มีการ
hardcode ตัวอักษรคอลัมน์ (`L`, `M`, `U`, `V`, `W` ฯลฯ) อีกต่อไป — ตัวแปรที่คำนวณไว้แล้วให้ใช้ตรงๆ:
- ใน RawData: `TGT_START/TGT_END/TGT_TOTAL_COL/TGT_PCT_COL` และ `ACT_START/ACT_END/ACT_TOTAL_COL/ACT_PCT_COL`
  และ `REV_COL` (พร้อมตัวอักษรคู่กัน `*_LTR`) — ใช้แทนทุกจุดที่เคยอ้างอิง RawData column ตรงๆ
- ใน Dashboard sheet: `COMP_FIELD_MAXCOL/COMP_TOTAL_COL/COMP_PCT_COL` (comp table) และตัวแปรใน scope ของ
  weekly table (`pct_col_idx`) คำนวณจาก `len(target_cols)` เช่นกัน
- ใน DailyMA sheet: `DM_FIELD_START/DM_TOTAL_COL/DM_PCT_COL`
**ถ้าจะเพิ่ม/ลดจำนวน field ในอนาคต แก้แค่ `FIELDS`/`FIELD_LABELS` ในตอน CONFIG แล้วโครงสร้างคอลัมน์ทั้งหมด
จะขยับตามอัตโนมัติ** ยกเว้น 2 จุดที่ยัง fix ไว้ตายตัวเพราะผูกกับ layout ภาพ (การ์ด breakdown สมมติว่าแบ่งได้
ลงตัวเป็นกลุ่มละ 4 — ถ้า field ไม่ใช่เลขคู่ของ 4 ต้องปรับ `bd_groups`/`bd_col_ranges` เอง) และตำแหน่งแถว
Breakdown/Comp table ที่ยัง fix (ดูหัวข้อ "ตำแหน่งแถว" ด้านล่าง)

### KPI cards — 5 ใบ (ไม่ใช่ 4)
ใบแรกสุดคือ **"Incentive ประจำวัน (บาท)"** ใช้ defined name `IncentiveActual`/`IncentiveTarget` (ไม่ได้
ผูกกับ `TodayDay` หรือข้อมูลรายวันใน `DAYS` เลย เพราะแหล่งข้อมูลจริงคือ `1.Dashboard!C3/C4` ของ Google
Sheet ต้นทาง คนละก้อนกับข้อมูลรายวัน) ตามด้วย Gross MA รายเดือน, Daily Gross MA วันล่าสุด, %MA/Revenue
วันล่าสุด, %MA/Revenue สะสมเดือน — คอลัมน์การ์ดคือ `A:C, D:F, G:I, J:L, M:P` (ไม่ใช่ 4 กลุ่มเท่ากันแบบเดิม)
**ถ้าลืมใส่ Incentive card ผู้ใช้จะทักว่า "ข้อมูล Incentive หายไป" เหมือนที่เคยเกิดขึ้นมาแล้ว**

### Breakdown เป็นการ์ดแนวนอน 2 แถว × 4 (ไม่ใช่ตารางแนวตั้ง)
ออกแบบให้เหมือนการ์ด `breakdown-grid` ของ HTML ให้มากที่สุด: แถวบน 4 ฟิลด์แรก (JRTS, Association Product,
Association Refer, Reciprocation), แถวล่าง 4 ฟิลด์ที่เหลือ (Shop Online 1, Shop Online 2, Project,
Service) — พอดี 8 ช่องเพราะ Shop Online แยกเป็น 1/2 แล้ว (ถ้าเคยเห็นเวอร์ชันเก่าที่มีช่องว่างเหลือ 1 ช่อง
นั่นเป็นของเวอร์ชันที่ยังรวม Shop Online เป็นค่าเดียว — เวอร์ชันปัจจุบันไม่มีช่องว่างแล้ว) การ์ดแต่ละใบมี
4 แถวย่อย: ชื่อฟิลด์ / Actual (สี+ไฟเขียวแดงตามเงื่อนไข) / Target / Gap badge (▲/▼)

### Target แบบ "เต็มงวด" vs "ถึงเมื่อวาน" — ใช้คนละที่กันโดยตั้งใจ
- **Period box / KPI cards (แถวบนสุด)** ใช้ Target **เต็มงวด** (`SUM(RawData!{TGT_TOTAL_LTR}2:...)` ทุกวันรวมอนาคต)
  เพราะ HTML dashboard เทียบ MTD Actual กับ Target เป้าทั้งงวดเสมอ
- **Breakdown cards (8 สายธุรกิจ)** ใช้ Target **ถึงเมื่อวานเท่านั้น** (`SUMIFS(...,"<"&TodayDay)`)
  เพราะ HTML ใช้ `pastDays` filter เฉพาะวันที่ผ่านมาแล้ว
- **ตารางสรุปรายสัปดาห์** ไม่กรองเลย ใช้ทั้งสัปดาห์รวมวันอนาคตด้วย (ตรงกับ HTML ที่ไม่ filter pastDays ในตารางนี้)
- **ตารางเทียบสัปดาห์อ้างอิง** แถว "Target" ใช้เต็มงวด, แถว "Actual สะสม" ใช้ถึงเมื่อวาน, ส่วนแถวอ้างอิงงวดอื่น
  (`REF_COMPARISON`) เป็นค่า hardcode เพราะเป็นข้อมูลคนละชุด/คนละช่วงเวลา — ทุกแถวหาร %MA ด้วย `RevenueMonth`
  เดียวกันเสมอ (ตรวจสอบแล้วว่าตรงกับตัวเลขใน Google Sheet ต้นฉบับ)

### ช่วงข้อมูลคาบเกี่ยว 2 เดือน (เช่น สัปดาห์ 27–30 = 29 มิ.ย.–26 ก.ค.) — บทเรียนสำคัญ
งวดรายงานอิงตามสัปดาห์ (`WEEKS`) ไม่ใช่เดือนปฏิทิน จึงมักคาบเกี่ยวข้าม 2 เดือนได้เสมอ **ห้ามสร้างข้อความวันที่
ด้วยการต่อเลขวันดิบกับ `MONTH_LABEL` คงที่ตัวเดียว** (เช่น `(TodayDay-1)&" "&MONTH_LABEL`) เพราะจะผิดทันทีที่
งวดข้ามเดือน — วิธีที่ถูกต้องคือดึงข้อความวันที่จริงจาก `RawData!C` ด้วย `INDEX/MATCH` เทียบกับ `dayNum` เสมอ
(ดูตัวอย่างในสูตรของ header A3, period box I5, banner มอบหมายงาน, และ label ของ comp table แถว
`COMP_ACT_ROW`) — field `date` ใน `DAYS` ต้องใส่ปี พ.ศ. เต็มเสมอ (`'D mmm. YYYY'`) เพราะจุดอ้างอิงพวกนี้ไม่มี
ตัวช่วยเติมปีให้แล้ว `DATE_RANGE_SHORT` มีไว้เฉพาะจุดที่คอลัมน์แคบ (เช่น comp table คอลัมน์ B กว้างแค่ ~18)
ไม่พอใส่ข้อความเต็มรูปแบบ "D month YYYY–D month YYYY"

### สีเขียว/แดง (green/red font)
Actual ดีกว่า Target (มากกว่า, น้อยติดลบกว่า) = เขียว, แย่กว่า = แดง ใช้ conditional formatting แบบ
`FormulaRule` เทียบเซลล์ Actual กับเซลล์ Target ที่ตำแหน่งคงที่ — ห้ามตัดสินจากเครื่องหมายบวก/ลบของตัวเลขเอง
(Gross MA ส่วนใหญ่ติดลบอยู่แล้ว ติดลบน้อยกว่า = ดีกว่า = เขียว)

### วันที่ยังไม่มีข้อมูล (อนาคต/วันนี้)
ในตาราง DailyMA แถว Actual/Gap ของวันที่ `dayNum >= TodayDay` จะเป็นค่าว่าง (สูตร `IF(...,value,"")`)
แทนข้อความ "ยังไม่มีข้อมูล" แบบ merge cell ของ HTML (ข้อจำกัดของ Excel ที่ merge แบบ data-driven ทำยาก)

### Column width ของตารางตัวเลข/ข้อความ
คอลัมน์ D:P ใน Dashboard sheet ต้อง**กว้างพอ** (ดู `dash_widths`) ไม่งั้นตัวเลขติดลบหลักแสนพร้อมวงเล็บ
(เช่น `(332,168.72)`) จะกลายเป็น `###` — คอลัมน์ A/B ต้องกว้างพอสำหรับ label ยาวๆ ของ comp table ด้วย
(ปัจจุบันตั้ง A=32, B=18 เพื่อรองรับ label แบบ "Target สัปดาห์ 27–30 (29 มิถุนายน – 26 กรกฎาคม 2569)") —
ถ้าผู้ใช้ขอเพิ่มทศนิยม/field ใหม่/label ยาวขึ้น ให้เพิ่มความกว้างตาม อย่าลดเนื้อหาข้อความเพื่อให้พอดี

### Number format ตามมาตรฐาน
ลบ = วงเล็บ `(123.45)` ไม่ใช่ `-123.45`, เปอร์เซ็นต์ = ทศนิยม 2 ตำแหน่ง `0.00%`, ค่าที่ผู้ใช้กรอกเอง (input)
ใช้สีน้ำเงิน (`0000FF`), สูตร/ผลคำนวณใช้สีดำปกติ — ตามมาตรฐานของ xlsx skill

### ตำแหน่งแถวแบบตายตัว (อย่าขยับโดยไม่อัปเดต reference)
Breakdown card อยู่แถว 39-46 (2 กลุ่ม × 4 แถวย่อย, ดูตัวแปร `BREAKDOWN_HDR_ROW`), Weekly table หัวอยู่แถว 50,
Comp table หัวอยู่แถว 66 (แถวข้อมูล 67-69) เสมอ — กราฟแท่ง Components ใน Dashboard sheet **อ้างอิงแถว
66-69 ตรงๆ** (ไม่ใช่ dynamic) ถ้าย้ายตำแหน่งตารางนี้ต้องอัปเดต `COMP_HDR_ROW`/`COMP_TGT_ROW`/`COMP_ACT_ROW`
ในสคริปต์ให้ตรงกันด้วย (คอลัมน์ของตารางเหล่านี้เป็น dynamic ตาม `N_FIELDS` แล้ว แต่ตำแหน่งแถวยังคงตายตัว)

### จำนวนวันในงวดไม่เท่ากับ 28 วัน
ถ้างวดนั้นมีจำนวนวันต่างจาก 28 (เช่น เดือนที่มี 30/31 วัน หรือขอบเขตตามสัปดาห์ที่ไม่ลงตัว) ให้เพิ่ม/ลด entry
ใน `DAYS` ตามจริง — โค้ดส่วน RawData และ DailyMA table คำนวณแถวต่อเนื่องจากความยาวของ `DAYS` อัตโนมัติ
(ใช้ `RD_LAST = 1 + len(DAYS)`) แต่ Breakdown/Weekly/Comp table ใน Dashboard sheet เป็นตำแหน่งแถวคงที่ตาม
ด้านบน (ไม่ได้ scale ตามจำนวนวัน) — เพียงพอสำหรับกรณีปกติ 28-31 วัน ตาราง Weekly ยังคงอิง `WEEKS` (4 สัปดาห์)
ถ้างวดมีมากกว่า/น้อยกว่า 4 สัปดาห์ต้องขยาย/ลดจำนวนแถวของตารางนี้เอง (โค้ดจะ loop ตามความยาวของ `WEEKS`
อัตโนมัติ แต่พื้นที่ที่เผื่อไว้ระหว่าง `WEEK_TBL_HDR_ROW` ถึง comp table แถว 65 อาจไม่พอถ้าสัปดาห์เกิน 4)

## โครงสร้างไฟล์ที่สร้าง (สรุปอย่างย่อ)

**Dashboard:** แถว 1-3 header, 5-11 period box ×2, 13-16 KPI ×5 (รวม Incentive), 18-35 กราฟ ×2,
38-46 breakdown (การ์ดแนวนอน 2×4), 49-62 weekly table, 65-69 comp table, 72-76 analysis, 79 banner

**DailyMA:** แถว 1-2 header, 4-6 KPI ×4, 8-9 table header, 10+ ตาราง 3 แถว/วัน (Target/Actual/Gap,
8 คอลัมน์สายธุรกิจ + Total + %MA/Rev), ตามด้วยกราฟเส้น + กราฟแท่งสะสม

**RawData:** แถว 1 header (25 คอลัมน์: 4 fixed + 8 target + total + pct + 8 actual + total + pct +
daily_revenue), แถว 2 ถึง `1+len(DAYS)` ข้อมูลรายวัน, ด้านล่างเป็น Config (`TodayDay`, `RevenueMonth`,
`Incentive Target`, `Incentive Actual`)

## ข้อจำกัดที่ทราบอยู่แล้ว (ไม่ต้องพยายามแก้ให้สมบูรณ์แบบ 100%)
- กราฟแท่งสะสมรายวันใน DailyMA ใช้ข้อมูลครบทุกวันของงวด (ไม่กรองเฉพาะวันที่ผ่านมาแล้วแบบ HTML)
- อีโมจิ (📊📅✅⚠️) อาจไม่แสดงผลใน LibreOffice headless แต่แสดงผลปกติเมื่อเปิดด้วย Excel/Google Sheets จริง
- Breakdown card ยังเป็นตำแหน่งแถวคงที่ (39-46) ไม่ scale ตามจำนวน field อัตโนมัติถ้า field ไม่ใช่เลขคู่ของ 4
