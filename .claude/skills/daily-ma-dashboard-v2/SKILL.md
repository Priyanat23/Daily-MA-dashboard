---
name: daily-ma-dashboard-v2
description: >
  สร้าง Daily MA Dashboard ครบชุด ทั้ง HTML (2-tab) และ Excel (.xlsx) ในคำสั่งเดียว สำหรับแผนก PCRC
  จาก Google Sheet ผ่าน Google Drive MCP — KPI Cards, Gap Analysis, กราฟ, ตาราง Target vs Actual,
  วิเคราะห์ผล, มอบหมายงาน (Available Time) ในตัว เช็ค %Gross MA/Revenue เทียบ Target ก่อนเสมอ ถ้า
  Actual ≥ Target ข้ามมอบหมายงาน ไม่ต้องอ่าน Sheet ที่สอง

  ใช้ทันทีเมื่อพูดว่า "daily MA dashboard", "daily MA", "daily MA report", "สร้าง dashboard",
  "อัปเดต dashboard", "dashboard MA ประจำวัน", "daily gross MA", "MA report" — ออกไฟล์ 2 ไฟล์เสมอ
  (HTML+Excel) เว้นแต่ระบุชัดว่าต้องการไฟล์เดียว

  Sheet หลัก FILE_ID `1q2ZYX9lTMLZoJu_HpLf93dCPHxBY3T2IawPf2fkfuAA` (1.Dashboard, 2.DailyMA, 3.BOM)
  Sheet ที่สอง (Available Time) FILE_ID `1INCAalbtHzflFgN5mZBsHcVEvOc5kIzY06c2icl6tCo`
---

# Daily MA Dashboard v2 — Skill สำหรับสร้าง HTML Dashboard (พร้อมส่วนมอบหมายงานในตัว)

## วิธีใช้ (ทำตาม 5 ขั้นตอนนี้เท่านั้น — รันครบทุกขั้นตอนแบบอัตโนมัติ ไม่ต้องถามผู้ใช้เพิ่ม)

> ⚡ **เพื่อความเร็ว**: อ่านเฉพาะ Google Sheet หลักก่อนเสมอในขั้นตอนที่ 1 แล้วเช็ค %Gross MA/Revenue
> ของสายธุรกิจ (ขั้นตอนที่ 2) **ก่อน** ตัดสินใจว่าต้องเปิด Google Sheet ที่สอง (Available Time)
> หรือไม่ — **ถ้า %Gross MA/Revenue Actual ≥ Target แล้ว ไม่ต้องมอบหมายงาน และห้าม
> download/อ่าน Google Sheet ที่สองเลยในรอบรันนี้** (ประหยัด 1 รอบ download ทั้งไฟล์) เปิด Google
> Sheet ที่สองเฉพาะกรณีที่จำเป็นต้องมอบหมายงานจริงเท่านั้น (ดูขั้นตอนที่ 3)

### ขั้นตอนที่ 1 — อ่านข้อมูล Google Sheet หลักก่อนเสมอ (ยังไม่ต้องแตะ Sheet ที่สอง)
```
1. Google Drive:download_file_content (exportMimeType=...spreadsheetml.sheet) กับ FILE_ID หลัก
   (1q2ZYX9lTMLZoJu_HpLf93dCPHxBY3T2IawPf2fkfuAA) แล้ว decode base64 → เปิดด้วย openpyxl
   อ่านชีต 1.Dashboard, 2.DailyMA (ทุกแถวรายวันของเดือนปัจจุบัน), 3.BOM

Google Sheet ที่สอง (1INCAalbtHzflFgN5mZBsHcVEvOc5kIzY06c2icl6tCo — สำหรับ Available Time) **ยังไม่
ต้องอ่านในขั้นนี้** จะอ่านเฉพาะเมื่อพบว่าต้องมอบหมายงานจริงตามเงื่อนไขในขั้นตอนที่ 2/3 ด้านล่างเท่านั้น
```

### ขั้นตอนที่ 2 — หา TODAY_DAYNUM และวันล่าสุดที่มี Actual
```
⚠️ พบจากการรันจริง (ก.ค. 2569): ข้อมูลรายวันใน `2.DailyMA` **ไม่ได้เริ่มที่วันที่ 1 ของเดือนเสมอไป**
— บางรอบเป็นช่วง 4 สัปดาห์แบบ rolling window ที่ข้ามเดือน (เช่น "29 มิ.ย.–26 ก.ค. 2569" สัปดาห์ 27–30,
ดูได้จากชื่อไฟล์ที่มักมีคำว่า "W<เลขสัปดาห์เริ่ม>-W<เลขสัปดาห์จบ>" กำกับ) **ห้ามสมมติว่า `dayNum`
เท่ากับ "เลขวันของเดือน" เด็ดขาด**

วิธีหา TODAY_DAYNUM ที่ถูกต้อง:
1. Parse วันที่จริงของแต่ละแถว (คอลัมน์ B ของแถว Target) เป็นค่าคงที่ `dayNum` แบบ**ลำดับตำแหน่งในตาราง
   นับจาก 1** (แถวแรกสุด = dayNum 1, แถวถัดไป = dayNum 2, ... ไม่ใช่เลขวันที่ปฏิทิน)
2. หาว่า "วันนี้จริง" (วันที่รัน skill นี้) ตรงกับแถวไหนในตาราง โดยเทียบวันที่ปฏิทินจริง (วัน-เดือน-ปี)
   กับคอลัมน์ B ของแต่ละแถว Target แล้วใช้ `dayNum` ของแถวนั้นเป็น `TODAY_DAYNUM`
3. ถ้าวันนี้จริงเกินขอบเขตข้อมูลในตาราง (เช่น ตารางมีถึงแค่วันที่ X แต่วันนี้เกินวันที่ X ไปแล้ว) หรือ
   วันนี้ยังไม่ปรากฏในตาราง (ข้อมูลยังไม่อัปเดตมาถึงช่วงนี้) ให้แจ้งเตือนผู้ใช้แทนการเดา/ประมาณค่า

TODAY_DAYNUM ที่ได้ (ตำแหน่งในตาราง) เป็นค่าคงที่ตายตัวในโค้ด HTML/JS ที่สร้างขึ้น ห้ามใช้
new Date().getDate() ในเบราว์เซอร์เด็ดขาด เพราะไฟล์ HTML เป็น snapshot ของวันที่สร้าง ถ้าเปิดไฟล์วันหลัง
แล้วใช้ new Date() ของเบราว์เซอร์ วันที่ "วันนี้/เมื่อวาน" จะเลื่อนผิดจากตอนคำนวณ

yesterdayDayNum = TODAY_DAYNUM - 1 (ใช้ตรงๆ ไม่ข้ามวันหยุด — Target มีทุกวันแม้วันหยุด)
ใช้แถวของ yesterdayDayNum ใน DAYS array เป็น "latestDay" สำหรับเช็ค %MA/Revenue Actual vs Target
```

### 🚦 เช็คเกท (gate check) ทันทีหลังได้ `latestDay` — ก่อนแตะ Google Sheet ที่สองเสมอ
เทียบ `latestDay.actual.pct` กับ `latestDay.target.pct` (%Gross MA/Revenue ของสายธุรกิจ) ทันที:
- **`latestDay.actual.pct >= latestDay.target.pct`** (ดีกว่าหรือเท่า Target) → **ไม่ต้องมอบหมายงาน
  เลย** ข้ามขั้นตอนที่ 3 ทั้งหมด (ไม่ต้องอ่าน Google Sheet ที่สอง ไม่ต้องคำนวณ Available Time/
  ELIGIBLE_RANKING) ไปแสดงแค่ "✅ ดีกว่า Target" ในส่วนมอบหมายงาน แล้วข้ามไปขั้นตอนที่ 4 ได้เลย
- **`latestDay.actual.pct < latestDay.target.pct`** (แย่กว่า Target) → ต้องมอบหมายงาน ไปต่อที่
  ขั้นตอนที่ 3 (ซึ่งจะเช็คต่อว่าเป็นวันหยุดหรือไม่ ก่อนตัดสินใจว่าต้องอ่าน Google Sheet ที่สองหรือไม่)

### โครงสร้างชีตจริง (อ้างอิงจากการรันจริง — ใช้ mapping นี้ตรงๆ ไม่ต้องไล่หาเองใหม่ทุกครั้ง)

**ชีต `2.DailyMA`** (แหล่งข้อมูลหลักสำหรับ DAYS array — ใช้ชีตนี้เป็นหลัก ไม่ใช่ `1.Dashboard`):
- แถว 8 และ 13 = header ซ้ำกัน คอลัมน์ D–M ตามลำดับคือ: JRTS, Association Product, Association Refer,
  Reciprocation, **Shop Online 1, Shop Online 2 (แยกกัน 2 คอลัมน์ ห้ามรวมเป็นตัวเดียว)**, Project,
  Service, Gross MA (=`total`), %Gross MA/Revenue (=`pct`)
- แถว 9–10 = สรุปเปรียบเทียบสัปดาห์อ้างอิง (Actual สัปดาห์ 10–13 มี.ค.) vs Target สัปดาห์ปัจจุบัน → ใช้
  เป็นค่า `refActual`/`curTarget` ตรงๆ ใน `buildCompTable()`
- แถว 14 เป็นต้นไป = ข้อมูลรายวัน **⚠️ ไม่ได้เริ่มที่วันที่ 1 ของเดือนเสมอไป — อาจเป็นช่วง 4 สัปดาห์แบบ
  rolling window ที่ข้ามเดือนก็ได้ (พบจริงจากการรัน ก.ค. 2569: เริ่ม 29 มิ.ย. ถึง 26 ก.ค. 2569 สัปดาห์
  27–30) ห้ามสมมติวันเริ่มต้นเด็ดขาด ให้อ่านวันที่จริงจากคอลัมน์ B ของแถวแรกเสมอ** **ทุกวันมี 3 แถว
  ติดกันเสมอ** (Target, Actual, Gap) คอลัมน์ A (สัปดาห์ที่) มีค่าแค่แถวแรกของแต่ละสัปดาห์ (แถวอื่น None
  — ต้อง forward-fill ค่าสัปดาห์ล่าสุดที่เจอไปยังแถว/วันถัดไปจนกว่าจะเจอค่าใหม่) คอลัมน์ B (วันที่) มีค่า
  แค่แถว Target ของแต่ละวัน — วิธี parse: loop หาแถวที่คอลัมน์ C = `'Target'` เพื่อระบุวันใหม่ แล้วอ่าน
  คอลัมน์ D–M ของแถวนั้น (Target) และแถว+1 (Actual) พร้อมกันเป็นคู่ กำหนด `dayNum` เป็นลำดับตำแหน่งที่
  parse ได้ (1, 2, 3, ... ) ไม่ใช่เลขวันที่ปฏิทิน (ดูขั้นตอนที่ 2)
- สรุปสะสมทั้งช่วง (Target/Actual/Gap) อยู่ถัดจากแถวข้อมูลรายวันแถวสุดท้าย **⚠️ ห้าม hardcode เลขแถว
  ตายตัว (เช่น "แถว 100–104") เพราะตำแหน่งเลื่อนได้ตามจำนวนวันในช่วงนั้น (28/30/31 วัน) — ให้หาแบบ
  dynamic แทน**: หลังจาก parse ข้อมูลรายวันครบทุกแถว (จนแถวคอลัมน์ C ไม่ใช่ Target/Actual/Gap อีกต่อไป)
  ให้มองหาแถวถัดไปที่คอลัมน์ D มีข้อความ `'Gross MA Back Office (บาท) สะสม'` แล้วอ่าน 3 แถวถัดจากนั้น
  เป็น Target/Actual/Gap ของยอดสะสมทั้งช่วง **ใช้ตรวจทาน (validate) เสมอก่อน build จริง**: sum ของ
  Target ทุกวันใน DAYS array ต้องเท่ากับแถว Target ที่หาเจอเป๊ะ, sum ของ Actual ต้องเท่ากับแถว Actual
  เป๊ะ — ถ้าไม่ตรงแสดงว่า parse ผิดแถว/คอลัมน์ (cross-check ได้อีกชั้นกับชีต `1.Dashboard` หัวข้อ
  "ข้อมูลสะสม" ซึ่งมีค่าเดียวกัน)
- ⚠️ **ยอดสะสม (Actual) ในแถวสรุปนี้อาจรวมข้อมูลบางส่วนของ "วันนี้" (TODAY_DAYNUM) ที่บันทึกเข้ามาแล้ว
  ด้วย** แม้วันนี้จะยังไม่จบวัน (พบจริง: มีค่า Shop Online 1 ของวันนี้ติดมาแล้วบางส่วน) ขณะที่
  `buildBreakdown()` ฝั่ง JS จะกรองเฉพาะ `dayNum<TODAY_DAYNUM` (ไม่รวมวันนี้) — ทำให้ตัวเลข breakdown
  รายช่องทางในหน้าเว็บกับยอดสะสมที่แสดงใน Period box/KPI card ต่างกันเล็กน้อยได้ **นี่คือพฤติกรรมปกติ
  ตามการออกแบบ ไม่ใช่ bug** ให้ใช้ยอดสะสมจากแถวสรุป (Target/Actual/Gap) ตรงๆ สำหรับ Period box, KPI
  card และ `curActual` ใน `buildCompTable()` เสมอ ไม่ต้องคำนวณจาก DAYS array ใหม่

**ชีต `1.Dashboard`**: ใช้แค่ cross-check ไม่ใช่แหล่งข้อมูลหลัก — แถว 6–8 (เปรียบเทียบสัปดาห์อ้างอิง),
แถว 12–14 (ข้อมูลของ "เมื่อวาน" ตรงกับ latestDay เสมอ), แถว 18–20 (สะสมเดือน) — เลขแถวจริงอาจขยับ
ขึ้นอยู่กับว่ามีการเพิ่มแถว Incentive ด้านบนหรือไม่ (ดูหัวข้อ Incentive ด้านล่าง) **ให้ยึดข้อความใน
คอลัมน์ B/C เป็นหลักในการหาตำแหน่งแถวจริง อย่ายึดเลขแถวตายตัว**

**Incentive ประจำวัน (บาท)** — อยู่ในชีต `1.Dashboard` เซลล์ `C3` = Target, เซลล์ `C4` = Actual
(แถว B3="Target", B4="Actual" กำกับไว้) **ต้องอ่านค่าใหม่ทุกครั้งที่รัน ห้าม hardcode เลขแถว 3/4
ตายตัวถ้าโครงสร้างชีตเปลี่ยน — ให้เช็คว่า A3="Incentive ประจำวัน (บาท)" ก่อนอ่าน C3/C4 เสมอ** ใช้ค่า
นี้สร้างเป็น KPI card ใบแรกสุดในหน้า Dashboard (ก่อนการ์ด Gross MA) เทียบ Actual(C4) vs Target(C3)
ตามกฎสีตัวเลข/ไฟกระพริบปกติ (Actual > Target = ดีกว่า = เขียว)

**ชีต `3.BOM`**: คอลัมน์ A–D (มีค่าแค่แถว 2–3) = `BOM_TASKS` จริง (EmpId/EmpName="ทุกคน") ส่วนคอลัมน์
F–H (แถว 2 เป็นต้นไป) เป็น**รายชื่อพนักงานทั้งแผนกแยกตามแผนกย่อย** (ธุรการ/วิจัย-ผลิต-บรรจุ/Q.C.) —
เป็นข้อมูลคนละก้อนในชีตเดียวกัน **ห้ามนำมาปนกับ BOM_TASKS** ใช้คอลัมน์นี้แทนหารหัสพนักงานเวลาต้อง join
ชื่อกับรหัส (เช่น "ปริยณัฐ รุจิชัยวิสิฐ" = รหัส 22360001)

**ชีตอื่นที่อาจพบในไฟล์ (ไม่ต้องใช้)**: ไฟล์จริงอาจมีชีตเพิ่มเติมนอกจาก 3 ชีตหลักข้างต้น เช่น
`4.มอบหมายงาน`, `Gross MA`, `Incentive`, `ข้อมูลจากพี่ปุ้ม`, `คำนวณ JRTS`, `กราฟ`,
`คำนวณ Target กล่องอื่น`, `JRTS PI`, `PI ใน JRTS` — **skill นี้ใช้แค่ `1.Dashboard`, `2.DailyMA`,
`3.BOM` เท่านั้น ไม่ต้องเปิด/อ่านชีตอื่นเลย** แม้จะมีอยู่ในไฟล์ก็ตาม

### กฎตรวจวันหยุด (holiday) — ห้ามเดาจากวันธรรมดา/เสาร์-อาทิตย์ เพียงอย่างเดียว
**`holiday = true` เมื่อ Target JRTS ของวันนั้นในชีต `2.DailyMA` (คอลัมน์ D แถว Target) = 0 เท่านั้น** —
นี่คือ flag ที่สูตรในชีตตั้งไว้ให้แล้ว ครอบคลุมทั้งวันเสาร์-อาทิตย์ และวันหยุดนักขัตฤกษ์ไทยที่ตรงกับ
วันธรรมดาด้วย (พบจริงจากการรัน: วันเฉลิมพระชนมพรรษาสมเด็จพระราชินี 3 มิ.ย. เป็นวันพุธ แต่ target jrts=0
เพราะเป็นวันหยุดนักขัตฤกษ์) **ห้ามคำนวณ holiday จาก day-of-week เพียงอย่างเดียวเด็ดขาด เพราะจะพลาด
วันหยุดนักขัตฤกษ์ที่ตรงกับวันธรรมดา** ฟิลด์อื่น (ap, ar, rc, so1, so2) ยังมีค่า target ปกติแม้เป็นวันหยุด
(target.total ของวันหยุดมักเป็นค่าคงที่ ~87.75–133.02 บาท จาก ap+so1+so2 หรือ ap+ar+rc+so1+so2 รวมกัน
แล้วแต่ข้อมูลจริงของชีต — ต้องอ่านค่าจริงเสมอ ไม่ hardcode)

### กฎปี พ.ศ. — ห้าม hardcode ปี พ.ศ. ในข้อความที่แสดงผล
ทุกจุดที่แสดงปี พ.ศ. (header meta, compTable label, section-title ฯลฯ) ต้อง**คำนวณจากปี ค.ศ. จริงของ
ข้อมูลในชีต + 543 เสมอ** ห้ามลอกปี พ.ศ. จาก template ตัวอย่างมาใช้ตรงๆ (เคยพบ bug จริง: template
ตัวอย่างใช้ "2568" สำหรับช่วงมีนาคม ทั้งที่ข้อมูลจริงในชีตคือมีนาคม ค.ศ. 2026 = พ.ศ. 2569 — ต้องตรวจปี
พ.ศ. ให้ตรงกับปี ค.ศ. ของวันที่ในชีตจริงทุกครั้ง ไม่ใช่เชื่อข้อความ placeholder)

### ทางลัดเมื่อวันนี้เป็นวันหยุด (กรณี A ในขั้นตอนที่ 3)
**ขั้นตอนนี้เข้าถึงได้เฉพาะเมื่อผ่านเกทเช็คด้านบนแล้วว่า `latestDay.actual.pct <
latestDay.target.pct` (ต้องมอบหมายงานจริง)** — ถ้าผ่านเงื่อนไขนั้นมาแล้ว ให้เช็คต่อว่า **วันนี้
(TODAY_DAYNUM) เป็นวันหยุด** หรือไม่ (เช็คจาก target jrts ของชีตหลักได้เลยตามกฎข้างต้น โดยไม่ต้องเปิด
Google Sheet ที่สองก่อน) → ถ้าเป็นวันหยุด **ไม่ต้อง download/อ่าน Google Sheet ที่สอง (Available
Time) เลยในรอบรันนี้** เพราะกรณี A มอบหมายให้ "ปริยณัฐ รุจิชัยวิสิฐ" คนเดียวตายตัว ข้ามการคำนวณ
Available Time/ELIGIBLE_RANKING ทั้งหมด (ดูขั้นตอนที่ 3 กรณี A) — ประหยัดขั้นตอนการอ่านไฟล์ที่ 2
ไปได้เลยในกรณีนี้ **เปิด/อ่าน Google Sheet ที่สองเฉพาะกรณี B เท่านั้น** (ต้องมอบหมายงานจริง และวันนี้
ไม่ใช่วันหยุด)

### ขั้นตอนที่ 3 — มอบหมายงาน (เข้าถึงเฉพาะเมื่อผ่านเกทเช็คแล้วว่า actual < target เท่านั้น)

เมื่อเข้าขั้นตอนนี้ (แปลว่าผ่านเกทเช็คด้านบนมาแล้วว่า `latestDay.actual.pct <
latestDay.target.pct`) ให้คำนวณ Gap ก่อนเสมอ (ใช้ร่วมกันทั้ง 2 กรณีด้านล่าง) — **คำนวณ
เป็นยอดเงินเท่านั้น ห้ามแปลงเป็นนาที/ห้ามหารด้วย Department rate อีกต่อไป**:
1. **ยอดที่ขาด (Gap Gross MA บาท) รวมทั้งวัน** = `latestDay.target.total - latestDay.actual.total`
   (ยอดเดียว ไม่แยกราย channel)
2. **ยอดที่ต้องทำให้ได้ (Net Revenue ที่ต้องทำเพิ่มเพื่อปิด Gap)** = ยอดที่ขาด ÷ 0.66 (Gross MA Ratio
   คงที่)

จากนั้น **เช็คว่า "วันที่มอบหมายงาน" คือ TODAY_DAYNUM (วันนี้ — วันที่รัน skill นี้ ไม่ใช่เมื่อวาน)
ตรงกับวันหยุดหรือไม่** (ดู `holiday` flag ของ entry ใน DAYS ที่ `dayNum === TODAY_DAYNUM`) เพื่อ
เลือกว่าจะมอบหมายงานแบบไหน — นี่คือกฎตัดสินใจหลัก:

#### กรณี A — วันนี้ (วันที่มอบหมายงาน) ตรงกับวันหยุด
มอบหมายให้ **"ปริยณัฐ รุจิชัยวิสิฐ" คนเดียวเท่านั้น** รับยอดที่ขาดและยอดที่ต้องทำให้ได้จากข้อ 1-2 ไป
ทั้งหมด โดยใช้ **เฉพาะ BOM-PCRC-0001-D1 (ติดตามลูกค้าใน Facebook (Lead, no show, repurchase))**
เท่านั้น — **ข้ามการคำนวณ Available Time / ELIGIBLE_RANKING ของขั้นที่ 4-7 ด้านล่างไปเลยในกรณีนี้**
(วันหยุดไม่มีข้อมูล Planning ของพนักงานคนอื่นตามปกติ) ห้ามมอบหมายให้คนอื่นเพิ่มในวันหยุด และห้ามใช้
BOM No. อื่นนอกจาก D1

#### รายชื่อพนักงานที่มีสิทธิ์รับมอบหมายงานในวันทำงาน (ASSIGNMENT_WHITELIST) — ใช้เฉพาะกรณี B เท่านั้น
**ในวันทำงาน (กรณี B) ให้มอบหมายงานได้เฉพาะพนักงานที่มีชื่ออยู่ใน whitelist นี้เท่านั้น** ห้ามมอบหมาย
ให้คนนอกรายชื่อนี้เด็ดขาดไม่ว่า Available Time จะสูงแค่ไหนก็ตาม (whitelist นี้ไม่เกี่ยวกับกรณี A ซึ่ง
มอบหมายให้ "ปริยณัฐ รุจิชัยวิสิฐ" คนเดียวตามเดิมเสมอ):
```
ASSIGNMENT_WHITELIST = [
  "วิลาวรรณ์ ก้อนกลิ่น", "หนึ่งนภา ตามระเบียบ", "กรกนก อรรถวิลัย", "ศุภลักษณ์ จั่นเพ็ชร์",
  "เรวัตร เจริญวงษ์", "วิจิตรา พันธ์แสง", "วัททลิกา สิริมิ่งขวัญ", "ชุติมา ทิพยกุล",
  "กัญญรส น้อยดัด", "ธาวิน เภาทอง", "อนุพงศ์ แสนกล้า", "เข็มพร พรมราช", "ศิริลักษณ์ ภู่สวรรค์ชัย"
]
```
Join รายชื่อนี้กับข้อมูล Planning/JRTS ด้วย**รหัสพนักงาน** ไม่ใช่ชื่อดิบ (หารหัสจาก `3.BOM` คอลัมน์
F–H หรือจาก Planning/JRTS sheet เอง) เพื่อกันปัญหาเขียนชื่อเพี้ยน/มีคำนำหน้าต่างกัน

#### กรณี B — วันนี้ (วันที่มอบหมายงาน) ไม่ตรงกับวันหยุด
**มอบหมายให้พนักงานอันดับ 1 ใน ELIGIBLE_RANKING เพียงคนเดียวเสมอ** (ไม่แบ่ง/ไม่ไล่มอบหมายหลายคน) รับ
ยอดที่ขาดและยอดที่ต้องทำให้ได้จากข้อ 1-2 ไปเต็มจำนวนทั้งหมด โดยใช้ **เฉพาะ BOM-PCRC-0001-A1 (โทร
ติดตามลูกค้า Association เพื่อซื้อสินค้า)** เท่านั้น — ห้ามใช้ BOM No. อื่นนอกจาก A1 ในกรณีนี้
**ห้ามมอบหมายให้ "ปริยณัฐ รุจิชัยวิสิฐ" เด็ดขาด** (เธอรับงานเฉพาะวันหยุดเท่านั้น — ต้องกรองชื่อเธอออก
จาก ELIGIBLE_RANKING ในข้อ 5 เสมอ แม้เธอจะมีแถว Planning วันนี้และ JobType="1-Office" ก็ตาม) และ
**ELIGIBLE_RANKING ต้องมีเฉพาะพนักงานที่อยู่ใน ASSIGNMENT_WHITELIST ด้านบนเท่านั้น** — คนที่มี
Available Time สูงแต่ไม่อยู่ใน whitelist ต้องถูกตัดออกจาก ranking ทั้งหมด (ไม่ใช่แค่ข้าม อันดับ 1):

4. **Available Time รายบุคคล** ของพนักงานในชีต `Planning เดือน<M>` วันที่ = TODAY_DAYNUM (วันนี้) —
   ใช้เพื่อ**จัดอันดับ (ranking) เท่านั้น** ไม่ได้ใช้แบ่งยอดเงิน/แปลงเป็นนาทีอีกต่อไป:
   - เวลาคงเหลือ = `max(0, 420 - TimePlanning รวมของคนนั้นในวันนี้)`
   - หา "วันที่มี Actual ล่าสุด" ในชีต `เดือน<M>` (JRTS) ก่อนวันนี้ — ถ้าวันก่อนหน้าไม่มีแถว JRTS เลย
     (วันหยุด) ให้ไล่ย้อนวันก่อนหน้าต่อไปจนเจอวันที่มีแถวจริง (ห้ามสมมติว่าเป็น "1 วันก่อนหน้า" ตายตัว)
   - BM Excess = `max(0, sum(PERFTIME ที่ PERFLEVEL=1) - 0.95×sum(TIMEMIN))` ของวันนั้น
   - ME Excess = `max(0, sum(PERFTIME ที่ PERFLEVEL=2) - 0.05×sum(TIMEMIN))` ของวันนั้น
   - Available Time = เวลาคงเหลือ + BM Excess + ME Excess
   - **กรองเฉพาะคนที่มีแถวในชีต Planning วันนี้และ JobType="1-Office" เท่านั้น และตัด "ปริยณัฐ
     รุจิชัยวิสิฐ" ออกเสมอไม่ว่าเงื่อนไขอื่นจะเข้าหรือไม่** — คนที่ไม่มีแถว Planning วันนี้เลย
     (ตรวจ JobType ไม่ได้) ให้ตัดออกจาก ranking และระบุชื่อไว้เป็นหมายเหตุ
     (อย่าสมมติว่า TimePlanning=0 แปลว่าว่างทั้งวัน)
   - **กรองเพิ่มอีกชั้น: ตัดคนที่ไม่อยู่ใน `ASSIGNMENT_WHITELIST` ออกจาก ranking ทั้งหมด** (กรองหลังจาก
     กรอง JobType="1-Office" และตัดปริยณัฐแล้ว) — ทำแม้ Available Time ของคนนั้นจะสูงที่สุดก็ตาม
   - Normalize ชื่อ/join ด้วยรหัสพนักงาน (ส่วนก่อน `-` ตัวแรกใน EMPLOYEE/Employee) ไม่ใช่ชื่อดิบ
5. เรียง Available Time มากไปน้อย → ได้ `ELIGIBLE_RANKING` (ต้องไม่มี "ปริยณัฐ รุจิชัยวิสิฐ" อยู่ในนี้
   และต้องมีเฉพาะคนใน `ASSIGNMENT_WHITELIST` เท่านั้น) — เก็บ list ไว้ทั้งหมดเป็นข้อมูลอ้างอิง แต่
   **ใช้มอบหมายจริงเฉพาะอันดับ 1 (`ELIGIBLE_RANKING[0]`) เท่านั้น** ถ้าไม่มีใครใน whitelist ผ่านเงื่อนไข
   เลย (list ว่าง) ให้แสดงข้อความแจ้งเตือนว่าไม่มีพนักงานที่มีสิทธิ์รับมอบหมายงานแทนการมอบหมาย
6. งานที่มอบหมาย = **BOM-PCRC-0001-A1 เท่านั้น** (ดึงคำอธิบายจากชีต `3.BOM` มาแสดงประกอบ) — ไม่ต้อง
   ดึง/แสดง BOM No. อื่น
7. คนที่ได้รับมอบหมาย = `ELIGIBLE_RANKING[0]` **คนเดียว** รับยอดที่ขาด (ข้อ 1) และยอดที่ต้องทำให้ได้
   (ข้อ 2) ไปเต็มจำนวนทั้งหมด (ไม่มีการแบ่ง/ทบยอดกับคนอื่น)

### ขั้นตอนที่ 4 — แปลงข้อมูลให้ตรงกับ DAYS array และค่าสรุป
ดึงข้อมูลจาก Google Sheet แล้วแทนค่าในตำแหน่งที่ระบุใน Template ด้านล่าง:

**ข้อมูลที่ต้องแทนค่า:**
| ตำแหน่ง | ข้อมูลที่ใส่ |
|---------|------------|
| `TODAY_DAYNUM` | เลขวันปัจจุบันจริง (ดูขั้นตอนที่ 2) |
| `DAYS array` | ข้อมูลรายวันทั้งหมด (ทุก entry ต้องมี `dayNum`) |
| `GROSS_MA_RATIO` | 0.66 เสมอ (ค่าคงที่ ไม่เปลี่ยน — ใช้แปลง Gap บาท → Net Revenue บาท เท่านั้น ไม่มีการแปลงเป็นนาที/Deptrate อีกต่อไป) |
| `BOM_TASK`, `ELIGIBLE_RANKING` | ผลลัพธ์จากขั้นตอนที่ 3 (คำนวณใหม่ทุกครั้งจากข้อมูลจริงวันนั้น — `BOM_TASK` ใช้แค่ BOM-PCRC-0001-A1 เท่านั้นสำหรับกรณีไม่ใช่วันหยุด, `ELIGIBLE_RANKING` ต้องกรองเฉพาะคนใน `ASSIGNMENT_WHITELIST` เท่านั้น) |
| Header meta | สัปดาห์, ช่วงเวลา, วันล่าสุด, Revenue |
| Period boxes | MTD Gross MA Actual/Target/Gap, %MA Actual/Target |
| KPI cards | ค่าสรุป 5 ใบ: (1) Incentive ประจำวัน จากชีต `1.Dashboard` C4 vs C3 — ดูหัวข้อ Incentive ด้านบน, (2) **Gap (Actual − Target) วันล่าสุด (บาท)** = `latestDay.actual.total - latestDay.target.total` (แสดง Actual/Target กำกับไว้ใต้ค่า Gap — **ไม่ใช้ Gross MA สะสม/รายเดือนแล้ว**), (3) Daily Gross MA วันล่าสุด (บาท), (4) %MA/Revenue วันล่าสุด, (5) %MA/Revenue สะสม |
| compTable data | refActual, curTarget, curActual |
| Analysis text | วิเคราะห์สาเหตุ/ที่มา/แนวทาง 3 ข้อ จากข้อมูลจริง |

### ขั้นตอนที่ 5 — สร้างไฟล์ HTML
```python
# บันทึกลงไฟล์
with open('/mnt/user-data/outputs/daily_ma_dashboard.html', 'w') as f:
    f.write(html_content)
```
**ยังไม่ต้องเรียก `present_files` ตรงนี้** — ให้ทำขั้นตอนที่ 6 (Excel) ให้เสร็จก่อน แล้ว present ทั้ง 2
ไฟล์พร้อมกันในขั้นตอนที่ 6 (เว้นแต่ผู้ใช้ระบุชัดเจนว่าต้องการแค่ HTML เท่านั้น ให้ข้ามขั้นตอนที่ 6 แล้ว
present ไฟล์ HTML ไฟล์เดียวได้เลย)

### ขั้นตอนที่ 6 — สร้างไฟล์ Excel (.xlsx) คู่กัน แล้ว present ทั้ง 2 ไฟล์พร้อมกัน
**ค่าเริ่มต้นของ skill นี้คือต้องออกทั้ง HTML และ Excel เสมอ** เว้นแต่ผู้ใช้จะขอแค่ไฟล์เดียวชัดเจน
(เช่น "แค่ html พอ") ให้ทำตามนี้:

1. **ห้ามอ่าน Google Sheet ซ้ำ** — ใช้ข้อมูล DAYS array, Incentive C3/C4, และค่าที่แปลงไว้แล้วจาก
   ขั้นตอนที่ 1–4 มาป้อนต่อให้สคริปต์ Excel ได้เลย (ข้อมูลชุดเดียวกัน)
2. คัดลอกสคริปต์จาก skill `daily-ma-dashboard-xlsx` มาไว้ที่ workspace:
   ```bash
   cp /mnt/skills/user/daily-ma-dashboard-xlsx/scripts/build_dashboard.py /home/claude/build_dashboard.py
   ```
3. แก้ไขเฉพาะส่วน `=== CONFIG: EDIT HERE ===` ด้านบนของ `build_dashboard.py` ให้ตรงกับข้อมูลชุดเดียวกับ
   ที่ใช้สร้าง HTML (ดูตาราง mapping ตัวแปรใน SKILL.md ของ `daily-ma-dashboard-xlsx` ถ้าไม่แน่ใจชื่อ
   ตัวแปรใดตัวแปรหนึ่ง) — **ห้ามแก้โครงสร้างสูตร/กราฟ/conditional formatting ส่วนอื่นของสคริปต์**
4. รันสคริปต์และตรวจสอบว่าไม่มี error ก่อนส่งเสมอ:
   ```bash
   python3 build_dashboard.py
   python3 /mnt/skills/public/xlsx/scripts/recalc.py /home/claude/daily_ma_dashboard.xlsx 60
   ```
   ต้องได้ `total_errors: 0` ถ้าไม่ใช่ ให้แก้สูตรที่ผิดตาม `error_summary` แล้วรันซ้ำ
5. คัดลอกไฟล์ผลลัพธ์ไปที่ `/mnt/user-data/outputs/daily_ma_dashboard.xlsx`
6. เรียก `present_files` ครั้งเดียวโดยส่งทั้ง 2 path พร้อมกัน:
   ```
   present_files(["/mnt/user-data/outputs/daily_ma_dashboard.html",
                  "/mnt/user-data/outputs/daily_ma_dashboard.xlsx"])
   ```

---

## กฎสำคัญที่ต้องรู้

### กฎสีตัวเลข Actual
```
actual > target → class="act-better"  สีเขียว #16a34a
actual < target → class="act-worse"   สีแดง   #dc2626
actual = target → class="act-equal"   สีดำ    #1f2937
```
> ⚠️ Gross MA มักติดลบ: -20,000 > -25,000 = ดีกว่า → act-better (สีเขียว)

### กฎไฟกระพริบ (Blinking Light)
ทุกจุดที่มีการตั้ง Target ให้แสดงไฟกระพริบหน้าค่า Actual เสมอ:
```
actual > target → <span class="light light-green"></span>  ไฟเขียวกระพริบ
actual < target → <span class="light light-red"></span>    ไฟแดงกระพริบ
actual = target หรือไม่มี Target → ไม่แสดงไฟ
```
ใช้ฟังก์ชัน `blink(actual, target)` ที่ให้มาในส่วน JavaScript เพื่อ generate HTML สัญลักษณ์ไฟ
ครอบคลุมทุกจุด: KPI cards, Period boxes, Breakdown cards, ตารางรายสัปดาห์, ตารางรายวัน

### กฎ buildAssignmentPanel (Available Time + Priority Ranking — ดูรายละเอียดเต็มในขั้นตอนที่ 3)
ใช้ **วันปัจจุบัน - 1 (yesterday)** เสมอสำหรับเช็ค %MA/Revenue เพื่อตัดสินใจว่าต้องมอบหมายงานหรือไม่
โดย `dayNum === TODAY_DAYNUM - 1` — ไม่ข้ามวันหยุด ใช้วันเมื่อวานตรงๆ (Target มีทุกวันแม้วันหยุด)
**นี่คือเกทเช็คตัวแรกสุดที่ต้องทำ (ดูหัวข้อ "🚦 เช็คเกท" ในขั้นตอนที่ 2) — ก่อนอ่าน/download Google
Sheet ที่สองเสมอ เพื่อประหยัดเวลา**
ถ้า `actual.pct >= target.pct` → แสดง ✅ ดีกว่า Target (ไม่ต้องอ่าน Google Sheet ที่สองเลย)

ถ้า `actual.pct < target.pct` → คำนวณ ยอดที่ขาด (Gap Gross MA บาท) **รวมทั้งวันยอดเดียว** (ไม่แยกราย
channel) แปลงผ่าน Gross MA Ratio คงที่ 66% → ยอดที่ต้องทำให้ได้ (Net Revenue เพื่อปิด Gap) —
**ไม่มีการแปลงเป็นนาที/ไม่หารด้วย Deptrate อีกต่อไป** จากนั้น **เช็ค `holiday` ของวันที่มอบหมายงาน คือ
entry ใน DAYS ที่ `dayNum === TODAY_DAYNUM` (วันนี้ — วันที่รัน skill นี้ ไม่ใช่เมื่อวาน)** เพื่อแยก
2 กรณี:
- **วันนี้เป็นวันหยุด** → มอบหมายยอดที่ขาด/ยอดที่ต้องทำให้ได้ทั้งหมดให้ "ปริยณัฐ รุจิชัยวิสิฐ" คนเดียว
  ใช้เฉพาะ BOM-PCRC-0001-D1 เท่านั้น (ไม่เรียก ELIGIBLE_RANKING เลย)
- **วันนี้ไม่ใช่วันหยุด** → มอบหมายยอดที่ขาด/ยอดที่ต้องทำให้ได้ทั้งหมดให้พนักงาน**อันดับ 1**ใน
  `ELIGIBLE_RANKING` (คำนวณไว้แล้วในขั้นตอนที่ 3 — เฉพาะคนที่มี Planning วันนี้, JobType="1-Office",
  ไม่ใช่ปริยณัฐ รุจิชัยวิสิฐ, **และต้องอยู่ใน `ASSIGNMENT_WHITELIST`**) **คนเดียวเสมอ ไม่แบ่ง/ไม่ไล่
  มอบหมายหลายคน** — งานที่มอบหมายใช้เฉพาะ BOM-PCRC-0001-A1 เท่านั้น

⚠️ "ปริยณัฐ รุจิชัยวิสิฐ" รับมอบหมายงาน **เฉพาะวันหยุดเท่านั้น** — ห้ามปรากฏใน ELIGIBLE_RANKING ของ
วันที่ไม่ใช่วันหยุด และในวันหยุดห้ามมอบหมายงานเพิ่มให้คนอื่นร่วมด้วย (วันหยุดมอบหมายให้เธอคนเดียว)
```javascript
const TODAY_DAYNUM = <เลขวันปัจจุบันจริง ณ ตอนสร้างไฟล์>;  // ค่าคงที่ตายตัว ห้ามใช้ new Date()
// d.dayNum < TODAY_DAYNUM → ผ่านมาแล้ว → แสดง Target/Actual/Gap ครบ 3 แถว
// d.dayNum >= TODAY_DAYNUM → วันนี้/อนาคต → Target + "ยังไม่มีข้อมูล"
// ใช้ DAYS.find(d=>d.dayNum===TODAY_DAYNUM) เพื่อเช็ค holiday ของ "วันที่มอบหมายงาน" (วันนี้) — ไม่ใช่
// holiday ของ latestDay (เมื่อวาน) ซึ่งใช้เช็คแค่เงื่อนไข actual<target เท่านั้น
```

### กฎ DAYS array
ทุก entry **ต้องมี** `dayNum` (ตัวเลขวันที่ เช่น 1–28)
```javascript
{week:23, date:'1 มิ.ย.', dayNum:1, holiday:true/false,
 target:{jrts, ap, ar, rc, so1, so2, pr, sv, total, pct},
 actual:{jrts, ap, ar, rc, so1, so2, pr, sv, total, pct}}
```

---

## HTML Template ที่ผ่านการทดสอบแล้ว (แทนค่าข้อมูลลงไปตรงๆ)

> **คำสั่ง**: ใช้ template นี้ทุกครั้ง — เปลี่ยนเฉพาะข้อมูลใน DAYS array, header meta,
> period boxes, KPI cards, compTable data และ analysis text เท่านั้น
> ห้ามเปลี่ยน CSS, function names, หรือโครงสร้าง HTML

```html
<!DOCTYPE html>
<html lang="th">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Daily MA Dashboard — PCRC</title>
<link href="https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;500;600;700&family=Prompt:wght@400;500;600;700&display=swap" rel="stylesheet">
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Sarabun',sans-serif;background:#f0f4f8;color:#1f2937;font-size:13px}
.tab-bar{display:flex;background:#fff;border-bottom:2px solid #dbeafe;position:sticky;top:0;z-index:100;box-shadow:0 2px 8px rgba(0,0,0,.06)}
.tab-btn{padding:12px 28px;font-family:'Sarabun',sans-serif;font-size:14px;font-weight:600;border:none;background:transparent;cursor:pointer;color:#6b7280;border-bottom:3px solid transparent;transition:all .2s}
.tab-btn.active{color:#1d4ed8;border-bottom-color:#1d4ed8;background:#f0f6ff}
.tab-content{display:none;padding:16px}
.tab-content.active{display:block}
.card{background:#fff;border-radius:10px;padding:16px;box-shadow:0 1px 4px rgba(0,0,0,.07);margin-bottom:14px}
.header-card{background:linear-gradient(135deg,#dbeafe,#ede9fe);border-left:5px solid #1d4ed8}
.header-card h2{font-family:'Prompt',sans-serif;font-size:18px;font-weight:700;color:#1e40af;margin-bottom:10px}
.header-meta{display:flex;flex-wrap:wrap;gap:8px 20px}
.meta-item{font-size:12px;color:#4b5563}.meta-item span{font-weight:600;color:#1e3a8a}
.period-row{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-bottom:14px}
.period-box{background:#fff;border-radius:10px;padding:14px;box-shadow:0 1px 4px rgba(0,0,0,.07)}
.period-box h4{font-size:13px;font-weight:700;color:#1e40af;margin-bottom:10px;border-bottom:1px solid #dbeafe;padding-bottom:6px}
.period-row-item{display:flex;justify-content:space-between;align-items:center;padding:4px 0;border-bottom:1px solid #f1f5f9}
.p-label{color:#6b7280;font-size:12px}.p-val{font-weight:600;font-size:13px}
.kpi-grid{display:grid;grid-template-columns:repeat(5,1fr);gap:12px;margin-bottom:14px}
.kpi-card{background:#fff;border-radius:10px;padding:14px;box-shadow:0 1px 4px rgba(0,0,0,.07);text-align:center}
.kpi-label{font-size:11px;color:#6b7280;margin-bottom:6px}
.kpi-actual{font-family:'Prompt',sans-serif;font-size:26px;font-weight:700;margin-bottom:4px}
.kpi-target{font-size:11px;color:#9ca3af;margin-bottom:6px}
.kpi-gap{display:inline-block;padding:3px 10px;border-radius:12px;font-size:11px;font-weight:700}
.gap-good{background:#dcfce7;color:#166534}.gap-bad{background:#fee2e2;color:#991b1b}.gap-neutral{background:#f1f5f9;color:#94a3b8}
.chart-row{display:grid;grid-template-columns:3fr 2fr;gap:14px;margin-bottom:14px}
.chart-card{background:#fff;border-radius:10px;padding:14px;box-shadow:0 1px 4px rgba(0,0,0,.07)}
.chart-card h4{font-size:12px;font-weight:700;color:#374151;margin-bottom:8px}
.breakdown-grid{display:grid;grid-template-columns:repeat(8,1fr);gap:10px;margin-bottom:14px}
.bd-card{background:#fff;border-radius:8px;padding:10px;box-shadow:0 1px 3px rgba(0,0,0,.06);text-align:center}
.bd-name{font-size:10px;font-weight:700;color:#374151;margin-bottom:4px}
.bd-target{font-size:10px;color:#9ca3af}
.bd-actual{font-size:14px;font-weight:700;margin:3px 0}
.tbl-wrap{background:#fff;border-radius:10px;box-shadow:0 1px 4px rgba(0,0,0,.07);margin-bottom:14px;overflow-x:auto}
table{width:100%;border-collapse:collapse}
th{background:#dbeafe;color:#1e40af;font-size:11px;font-weight:700;padding:8px 6px;text-align:center;white-space:nowrap}
td{padding:6px 6px;font-size:12px;text-align:right;border-bottom:1px solid #f1f5f9}
td:first-child,td:nth-child(2),td:nth-child(3){text-align:center}
tr.row-target td{background:#f8fafc;color:#374151}
tr.row-actual td{background:#fff;color:#1f2937}
tr.row-gap td{font-size:11px;font-weight:700}
tr.row-holiday td{background:#fef9ec}
.gap-good-cell{background:#dcfce7;color:#166534}
.gap-bad-cell{background:#fee2e2;color:#991b1b}
.gap-neu-cell{background:#f1f5f9;color:#94a3b8}
.nodata-row td{background:#fffbeb;color:#92400e;font-size:11px}
.holiday-badge{display:inline-block;font-size:10px;background:#fef3c7;color:#92400e;border-radius:4px;padding:1px 5px;margin-left:4px;font-weight:600}
.section-title{font-family:'Prompt',sans-serif;font-size:15px;font-weight:600;color:#1e40af;margin-bottom:10px;padding-left:6px;border-left:4px solid #1d4ed8}
.act-better{color:#16a34a}.act-worse{color:#dc2626}.act-equal{color:#1f2937}
.positive{color:#16a34a}.negative{color:#dc2626}
@keyframes blink-green{0%,100%{opacity:1;box-shadow:0 0 6px #16a34a}50%{opacity:.3;box-shadow:none}}
@keyframes blink-red{0%,100%{opacity:1;box-shadow:0 0 6px #dc2626}50%{opacity:.3;box-shadow:none}}
.light{display:inline-block;width:10px;height:10px;border-radius:50%;margin-right:4px;vertical-align:middle}
.light-green{background:#16a34a;animation:blink-green 1.2s ease-in-out infinite}
.light-red{background:#dc2626;animation:blink-red 1.2s ease-in-out infinite}
.analysis-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:14px;margin-bottom:14px}
.analysis-card{background:#fff;border-radius:10px;padding:14px;box-shadow:0 1px 4px rgba(0,0,0,.07)}
.analysis-card h4{font-size:13px;font-weight:700;color:#1e40af;margin-bottom:10px;padding-bottom:6px;border-bottom:1px solid #dbeafe}
.analysis-item{display:flex;gap:8px;margin-bottom:8px;align-items:flex-start}
.analysis-num{min-width:22px;height:22px;background:#dbeafe;color:#1e40af;border-radius:50%;font-size:11px;font-weight:700;display:flex;align-items:center;justify-content:center}
.analysis-text{font-size:12px;color:#374151;line-height:1.5}
.assignment-panel{background:#fff;border-radius:10px;padding:16px;box-shadow:0 1px 4px rgba(0,0,0,.07);margin-bottom:14px}
.assignment-note{font-size:12px;color:#6b7280;margin-bottom:12px;line-height:1.6}
.assignment-panel.all-good{background:#dcfce7;border:1px solid #86efac}
.assignment-panel h3{font-size:14px;font-weight:700;color:#1e40af;margin-bottom:12px}
.assignment-grid-inner{display:grid;grid-template-columns:repeat(auto-fill,minmax(280px,1fr));gap:14px}
.task-card{background:#fff5f5;border:1px solid #fecaca;border-radius:8px;padding:14px}
.task-header{display:flex;align-items:center;gap:8px;margin-bottom:8px;flex-wrap:wrap}
.bom-id{font-size:11px;background:#dbeafe;color:#1e40af;padding:2px 8px;border-radius:4px;font-family:monospace}
.gap-badge-red{font-size:12px;background:#fee2e2;color:#dc2626;padding:2px 8px;border-radius:4px;font-weight:600}
.task-desc{color:#1f2937;font-weight:500;margin-bottom:8px;font-size:13px}
.assigned-emp{color:#1d4ed8;font-weight:600;margin-top:6px;font-size:13px}
.net-rev{color:#dc2626;font-weight:700;margin-top:4px;font-size:13px}
.hc-info{color:#6b7280;font-size:12px;margin-top:4px}
.daily-tbl th{font-size:11px;padding:6px 8px}
.daily-tbl td{padding:5px 8px;font-size:12px}
.daily-tbl tr.row-target td{background:#f8fafc;color:#374151}
.daily-tbl tr.row-holiday td{background:#fef9ec;color:#374151}
.daily-tbl tr.row-actual td{background:#fff;color:#1f2937}
.daily-tbl tr.row-gap td{font-size:11px;font-weight:700}
.daily-tbl .nodata-row td{background:#fffbeb;color:#92400e;font-size:11px}
@media(max-width:900px){
  .kpi-grid{grid-template-columns:repeat(2,1fr)}
  .chart-row{grid-template-columns:1fr}
  .breakdown-grid{grid-template-columns:repeat(4,1fr)}
  .period-row{grid-template-columns:1fr}
  .analysis-grid{grid-template-columns:1fr}
}
</style>
</head>
<body>
<div class="tab-bar">
  <button class="tab-btn active" onclick="showTab('dashboard',event)">📊 Dashboard</button>
  <button class="tab-btn" onclick="showTab('dailyma',event)">📅 DailyMA</button>
</div>

<!-- TAB 1: DASHBOARD -->
<div id="tab-dashboard" class="tab-content active">
  <div class="card header-card">
    <h2>Daily MA Dashboard — PCRC</h2>
    <div class="header-meta">
      <div class="meta-item">หน่วยงาน: <span>PCRC</span></div>
      <div class="meta-item">สัปดาห์ที่: <span>23–26</span></div>
      <div class="meta-item">ช่วงเวลา: <span>1–28 มิถุนายน 2569</span></div>
      <div class="meta-item">ข้อมูล ณ วันที่: <span>10 มิถุนายน 2569</span></div>
      <div class="meta-item">Revenue สายธุรกิจ: <span>฿5,927,657.11</span></div>
      <div class="meta-item">เปรียบเทียบกับ: <span>สัปดาห์ 10–13 (มี.ค. 2569)</span></div>
    </div>
  </div>

  <div class="period-row">
    <div class="period-box">
      <h4>📆 สะสมเดือน (1–28 มิถุนายน 2569) สัปดาห์ 23–26</h4>
      <div class="period-row-item"><span class="p-label">Revenue ของสายธุรกิจ</span><span class="p-val">฿5,927,657.11</span></div>
      <div class="period-row-item"><span class="p-label">Gross MA Actual</span><span class="p-val act-better"><span class="light light-green"></span>฿-140,795.61</span></div>
      <div class="period-row-item"><span class="p-label">Gross MA Target</span><span class="p-val">฿-481,752.78</span></div>
      <div class="period-row-item"><span class="p-label">Gap (Actual − Target)</span><span class="kpi-gap gap-good">▲ +340,957.17</span></div>
      <div class="period-row-item"><span class="p-label">%MA/Revenue Actual</span><span class="p-val act-better">-2.38%</span></div>
      <div class="period-row-item"><span class="p-label">%MA/Revenue Target</span><span class="p-val">-8.13%</span></div>
    </div>
    <div class="period-box">
      <h4>📅 ล่าสุด วันที่ 10 มิถุนายน 2569 (สัปดาห์ 24)</h4>
      <div class="period-row-item"><span class="p-label">Revenue ของสายธุรกิจ</span><span class="p-val">฿211,702.04</span></div>
      <div class="period-row-item"><span class="p-label">Daily Gross MA Actual</span><span class="p-val act-better"><span class="light light-green"></span>฿-19,773.86</span></div>
      <div class="period-row-item"><span class="p-label">Daily Gross MA Target</span><span class="p-val">฿-25,418.42</span></div>
      <div class="period-row-item"><span class="p-label">Gap (Actual − Target)</span><span class="kpi-gap gap-good">▲ +5,644.56</span></div>
      <div class="period-row-item"><span class="p-label">%MA/Revenue Actual</span><span class="p-val act-better">-9.34%</span></div>
      <div class="period-row-item"><span class="p-label">%MA/Revenue Target</span><span class="p-val">-12.01%</span></div>
    </div>
  </div>

  <div class="kpi-grid">
    <div class="kpi-card">
      <div class="kpi-label">Incentive ประจำวัน (บาท)</div>
      <div class="kpi-actual act-better"><span class="light light-green"></span>941</div>
      <div class="kpi-target">Target: 26</div>
      <span class="kpi-gap gap-good">▲ + ดีกว่า Target 915</span>
    </div>
    <div class="kpi-card">
      <div class="kpi-label">Gap (Actual − Target) วันล่าสุด (บาท)</div>
      <div class="kpi-actual act-better"><span class="light light-green"></span>+5,645</div>
      <div class="kpi-target">Actual: -19,774 | Target: -25,418</div>
      <span class="kpi-gap gap-good">▲ + ดีกว่า Target 5,645</span>
    </div>
    <div class="kpi-card">
      <div class="kpi-label">Daily Gross MA วันล่าสุด (บาท)</div>
      <div class="kpi-actual act-better"><span class="light light-green"></span>-19,774</div>
      <div class="kpi-target">Target: -25,418</div>
      <span class="kpi-gap gap-good">▲ + ดีกว่า Target 5,645</span>
    </div>
    <div class="kpi-card">
      <div class="kpi-label">%MA/Revenue วันล่าสุด</div>
      <div class="kpi-actual act-better"><span class="light light-green"></span>-9.34%</div>
      <div class="kpi-target">Target: -12.01%</div>
      <span class="kpi-gap gap-good">▲ + ดีกว่า Target 2.67pp</span>
    </div>
    <div class="kpi-card">
      <div class="kpi-label">%MA/Revenue สะสมเดือน</div>
      <div class="kpi-actual act-better"><span class="light light-green"></span>-2.38%</div>
      <div class="kpi-target">Target: -8.13%</div>
      <span class="kpi-gap gap-good">▲ + ดีกว่า Target 5.75pp</span>
    </div>
  </div>

  <div class="chart-row">
    <div class="chart-card">
      <div style="height:300px;position:relative"><canvas id="maLineChart"></canvas></div>
    </div>
    <div class="chart-card">
      <h4>Components สะสม (Target vs Actual)</h4>
      <div style="height:268px;position:relative"><canvas id="barChart"></canvas></div>
    </div>
  </div>

  <p class="section-title">Breakdown รายสายธุรกิจ (สะสม 1–10 มิถุนายน)</p>
  <div class="breakdown-grid" id="breakdownGrid"></div>

  <p class="section-title">ตารางสรุปรายสัปดาห์</p>
  <div class="tbl-wrap card" style="padding:0;overflow:hidden">
    <table id="weeklyTable"></table>
  </div>

  <p class="section-title">เปรียบเทียบกับสัปดาห์อ้างอิง (สัปดาห์ 10–13 มี.ค. 2569)</p>
  <div class="tbl-wrap card" style="padding:0;overflow:hidden">
    <table id="compTable"></table>
  </div>

  <p class="section-title">การวิเคราะห์ผล</p>
  <div class="analysis-grid">
    <div class="analysis-card">
      <h4>🔍 สาเหตุ</h4>
      <div class="analysis-item"><div class="analysis-num">1</div><div class="analysis-text">JRTS สะสม -142,875.70 บาท เป็นต้นทุนแรงงานหลักที่ดึง MA ลง คิดเป็น 101.5% ของ MA ติดลบทั้งหมด</div></div>
      <div class="analysis-item"><div class="analysis-num">2</div><div class="analysis-text">Association Product, Refer, Reciprocation ยังไม่มี Actual ตลอด 10 วันที่ผ่านมา (0 บาท)</div></div>
      <div class="analysis-item"><div class="analysis-num">3</div><div class="analysis-text">Shop Online สะสม 2,080.09 บาท ดีกว่า Target 306.64 บาท อย่างมีนัยสำคัญ (+579%)</div></div>
    </div>
    <div class="analysis-card">
      <h4>📌 ที่มาของสาเหตุ</h4>
      <div class="analysis-item"><div class="analysis-num">1</div><div class="analysis-text">แผนก PCRC มีพนักงาน 23 คน ทำให้ JRTS สูงตามสัดส่วน ยากที่จะลดต้นทุนแรงงานโดยตรง</div></div>
      <div class="analysis-item"><div class="analysis-num">2</div><div class="analysis-text">ข้อมูล Association ยังไม่ถูกบันทึกลงระบบ หรือยังไม่มีการโทรติดตามลูกค้าในช่วง 10 วันแรก</div></div>
      <div class="analysis-item"><div class="analysis-num">3</div><div class="analysis-text">%MA/Revenue สะสม -2.38% ดีกว่า Target -8.13% เนื่องจาก JRTS Actual ต่ำกว่า Target มาก</div></div>
    </div>
    <div class="analysis-card">
      <h4>✅ แนวทางแก้ปัญหา</h4>
      <div class="analysis-item"><div class="analysis-num">1</div><div class="analysis-text">เร่งบันทึก Association Product/Refer ที่ตกค้าง และกระตุ้นทีมโทรติดตามลูกค้าให้ครบ Target 76.80 บาท/วัน</div></div>
      <div class="analysis-item"><div class="analysis-num">2</div><div class="analysis-text">เพิ่มยอด Reciprocation ติดตามลูกค้าซื้อซ้ำ เป้า 38.40 บาท/วัน ซึ่งยังเป็นศูนย์ทุกวัน</div></div>
      <div class="analysis-item"><div class="analysis-num">3</div><div class="analysis-text">รักษา Shop Online ต่อเนื่อง ปัจจุบัน 2,080 บาทสะสม ให้คงอัตรา 208 บาท/วัน ต่อไป</div></div>
    </div>
  </div>
  <div id="assignmentPanel"></div>
</div>

<!-- TAB 2: DAILYMA -->
<div id="tab-dailyma" class="tab-content">
  <div class="card header-card">
    <h2>DailyMA — รายวัน มิถุนายน 2569</h2>
    <div class="header-meta">
      <div class="meta-item">ช่วง: <span>1–28 มิถุนายน 2569</span></div>
      <div class="meta-item">สัปดาห์: <span>23–26</span></div>
      <div class="meta-item">Revenue: <span>฿5,927,657.11</span></div>
    </div>
  </div>
  <div class="kpi-grid" style="grid-template-columns:repeat(4,1fr)">
    <div class="kpi-card">
      <div class="kpi-label">JRTS สะสม (Actual)</div>
      <div class="kpi-actual act-better" style="font-size:20px"><span class="light light-green"></span>-142,876</div>
      <div class="kpi-target">Target: -485,477</div>
    </div>
    <div class="kpi-card">
      <div class="kpi-label">Shop Online สะสม (Actual)</div>
      <div class="kpi-actual act-better" style="font-size:20px"><span class="light light-green"></span>2,080</div>
      <div class="kpi-target">Target: 307</div>
    </div>
    <div class="kpi-card">
      <div class="kpi-label">Total MA สะสม (Actual)</div>
      <div class="kpi-actual act-better" style="font-size:20px"><span class="light light-green"></span>-140,796</div>
      <div class="kpi-target">Target: -481,753</div>
    </div>
    <div class="kpi-card">
      <div class="kpi-label">%MA/Rev สะสม</div>
      <div class="kpi-actual act-better" style="font-size:20px"><span class="light light-green"></span>-2.38%</div>
      <div class="kpi-target">Target: -8.13%</div>
    </div>
  </div>
  <div class="tbl-wrap card" style="padding:0;overflow:hidden">
    <table class="daily-tbl" id="dailyTable"></table>
  </div>
  <div class="card" style="margin-top:16px">
    <div style="height:300px;position:relative"><canvas id="maLineChart2"></canvas></div>
  </div>
  <div class="card" style="margin-top:16px">
    <h4 style="font-size:13px;font-weight:700;color:#374151;margin-bottom:12px">Components รายวัน Actual (บาท)</h4>
    <div style="height:260px;position:relative"><canvas id="stackedBar"></canvas></div>
  </div>
</div>

<script>
const TODAY_DAYNUM = 24; // ⚠️ แทนค่าด้วยเลขวันปัจจุบันจริง ณ ตอนสร้างไฟล์ทุกครั้ง (ห้ามใช้ new Date())
const DAYS = [
  {week:23,date:'1 มิ.ย.',dayNum:1,target:{jrts:-25551.44,ap:76.80,ar:6.87,rc:38.40,so1:10.95,so2:0,pr:0,sv:0,total:-25418.42,pct:-0.1201},actual:{jrts:-27417.13,ap:0,ar:0,rc:0,so1:413.86,so2:0,pr:0,sv:0,total:-27003.27,pct:-0.1276}},
  {week:23,date:'2 มิ.ย.',dayNum:2,target:{jrts:-25551.44,ap:76.80,ar:6.87,rc:38.40,so1:10.95,so2:0,pr:0,sv:0,total:-25418.42,pct:-0.1201},actual:{jrts:-14747.74,ap:0,ar:0,rc:0,so1:481.50,so2:0,pr:0,sv:0,total:-14266.24,pct:-0.0674}},
  {week:23,date:'3 มิ.ย.',dayNum:3,holiday:true,target:{jrts:0,ap:76.80,ar:6.87,rc:38.40,so1:10.95,so2:0,pr:0,sv:0,total:133.02,pct:0.0006},actual:{jrts:0,ap:0,ar:0,rc:0,so1:159.81,so2:0,pr:0,sv:0,total:159.81,pct:0.0008}},
  {week:23,date:'4 มิ.ย.',dayNum:4,target:{jrts:-25551.44,ap:76.80,ar:6.87,rc:38.40,so1:10.95,so2:0,pr:0,sv:0,total:-25418.42,pct:-0.1201},actual:{jrts:-20868.48,ap:0,ar:0,rc:0,so1:139.15,so2:0,pr:0,sv:0,total:-20729.33,pct:-0.0979}},
  {week:23,date:'5 มิ.ย.',dayNum:5,target:{jrts:-25551.44,ap:76.80,ar:6.87,rc:38.40,so1:10.95,so2:0,pr:0,sv:0,total:-25418.42,pct:-0.1201},actual:{jrts:-19590.12,ap:0,ar:0,rc:0,so1:145.11,so2:0,pr:0,sv:0,total:-19445.01,pct:-0.0919}},
  {week:23,date:'6 มิ.ย.',dayNum:6,holiday:true,target:{jrts:0,ap:76.80,ar:6.87,rc:38.40,so1:10.95,so2:0,pr:0,sv:0,total:133.02,pct:0.0006},actual:{jrts:0,ap:0,ar:0,rc:0,so1:0,so2:0,pr:0,sv:0,total:0,pct:0}},
  {week:23,date:'7 มิ.ย.',dayNum:7,holiday:true,target:{jrts:0,ap:76.80,ar:6.87,rc:38.40,so1:10.95,so2:0,pr:0,sv:0,total:133.02,pct:0.0006},actual:{jrts:0,ap:0,ar:0,rc:0,so1:0,so2:0,pr:0,sv:0,total:0,pct:0}},
  {week:24,date:'8 มิ.ย.',dayNum:8,target:{jrts:-25551.44,ap:76.80,ar:6.87,rc:38.40,so1:10.95,so2:0,pr:0,sv:0,total:-25418.42,pct:-0.1201},actual:{jrts:-18909.30,ap:0,ar:0,rc:0,so1:51.17,so2:0,pr:0,sv:0,total:-18858.13,pct:-0.0891}},
  {week:24,date:'9 มิ.ย.',dayNum:9,target:{jrts:-25551.44,ap:76.80,ar:6.87,rc:38.40,so1:10.95,so2:0,pr:0,sv:0,total:-25418.42,pct:-0.1201},actual:{jrts:-21181.48,ap:0,ar:0,rc:0,so1:301.90,so2:0,pr:0,sv:0,total:-20879.58,pct:-0.0986}},
  {week:24,date:'10 มิ.ย.',dayNum:10,target:{jrts:-25551.44,ap:76.80,ar:6.87,rc:38.40,so1:10.95,so2:0,pr:0,sv:0,total:-25418.42,pct:-0.1201},actual:{jrts:-20161.45,ap:0,ar:0,rc:0,so1:387.59,so2:0,pr:0,sv:0,total:-19773.86,pct:-0.0934}},
  {week:24,date:'11 มิ.ย.',dayNum:11,target:{jrts:-25551.44,ap:76.80,ar:6.87,rc:38.40,so1:10.95,so2:0,pr:0,sv:0,total:-25418.42,pct:-0.1201},actual:{jrts:0,ap:0,ar:0,rc:0,so1:0,so2:0,pr:0,sv:0,total:0,pct:0}},
  {week:24,date:'12 มิ.ย.',dayNum:12,target:{jrts:-25551.44,ap:76.80,ar:6.87,rc:38.40,so1:10.95,so2:0,pr:0,sv:0,total:-25418.42,pct:-0.1201},actual:{jrts:0,ap:0,ar:0,rc:0,so1:0,so2:0,pr:0,sv:0,total:0,pct:0}},
  {week:24,date:'13 มิ.ย.',dayNum:13,holiday:true,target:{jrts:0,ap:76.80,ar:6.87,rc:38.40,so1:10.95,so2:0,pr:0,sv:0,total:133.02,pct:0.0006},actual:{jrts:0,ap:0,ar:0,rc:0,so1:0,so2:0,pr:0,sv:0,total:0,pct:0}},
  {week:24,date:'14 มิ.ย.',dayNum:14,holiday:true,target:{jrts:0,ap:76.80,ar:6.87,rc:38.40,so1:10.95,so2:0,pr:0,sv:0,total:133.02,pct:0.0006},actual:{jrts:0,ap:0,ar:0,rc:0,so1:0,so2:0,pr:0,sv:0,total:0,pct:0}},
  {week:25,date:'15 มิ.ย.',dayNum:15,target:{jrts:-25551.44,ap:76.80,ar:6.87,rc:38.40,so1:10.95,so2:0,pr:0,sv:0,total:-25418.42,pct:-0.1201},actual:{jrts:0,ap:0,ar:0,rc:0,so1:0,so2:0,pr:0,sv:0,total:0,pct:0}},
  {week:25,date:'16 มิ.ย.',dayNum:16,target:{jrts:-25551.44,ap:76.80,ar:6.87,rc:38.40,so1:10.95,so2:0,pr:0,sv:0,total:-25418.42,pct:-0.1201},actual:{jrts:0,ap:0,ar:0,rc:0,so1:0,so2:0,pr:0,sv:0,total:0,pct:0}},
  {week:25,date:'17 มิ.ย.',dayNum:17,target:{jrts:-25551.44,ap:76.80,ar:6.87,rc:38.40,so1:10.95,so2:0,pr:0,sv:0,total:-25418.42,pct:-0.1201},actual:{jrts:0,ap:0,ar:0,rc:0,so1:0,so2:0,pr:0,sv:0,total:0,pct:0}},
  {week:25,date:'18 มิ.ย.',dayNum:18,target:{jrts:-25551.44,ap:76.80,ar:6.87,rc:38.40,so1:10.95,so2:0,pr:0,sv:0,total:-25418.42,pct:-0.1201},actual:{jrts:0,ap:0,ar:0,rc:0,so1:0,so2:0,pr:0,sv:0,total:0,pct:0}},
  {week:25,date:'19 มิ.ย.',dayNum:19,target:{jrts:-25551.44,ap:76.80,ar:6.87,rc:38.40,so1:10.95,so2:0,pr:0,sv:0,total:-25418.42,pct:-0.1201},actual:{jrts:0,ap:0,ar:0,rc:0,so1:0,so2:0,pr:0,sv:0,total:0,pct:0}},
  {week:25,date:'20 มิ.ย.',dayNum:20,holiday:true,target:{jrts:0,ap:76.80,ar:6.87,rc:38.40,so1:10.95,so2:0,pr:0,sv:0,total:133.02,pct:0.0006},actual:{jrts:0,ap:0,ar:0,rc:0,so1:0,so2:0,pr:0,sv:0,total:0,pct:0}},
  {week:25,date:'21 มิ.ย.',dayNum:21,holiday:true,target:{jrts:0,ap:76.80,ar:6.87,rc:38.40,so1:10.95,so2:0,pr:0,sv:0,total:133.02,pct:0.0004},actual:{jrts:0,ap:0,ar:0,rc:0,so1:0,so2:0,pr:0,sv:0,total:0,pct:0}},
  {week:26,date:'22 มิ.ย.',dayNum:22,target:{jrts:-25551.44,ap:76.80,ar:6.87,rc:38.40,so1:10.95,so2:0,pr:0,sv:0,total:-25418.42,pct:-0.1201},actual:{jrts:0,ap:0,ar:0,rc:0,so1:0,so2:0,pr:0,sv:0,total:0,pct:0}},
  {week:26,date:'23 มิ.ย.',dayNum:23,target:{jrts:-25551.44,ap:76.80,ar:6.87,rc:38.40,so1:10.95,so2:0,pr:0,sv:0,total:-25418.42,pct:-0.1201},actual:{jrts:0,ap:0,ar:0,rc:0,so1:0,so2:0,pr:0,sv:0,total:0,pct:0}},
  {week:26,date:'24 มิ.ย.',dayNum:24,target:{jrts:-25551.44,ap:76.80,ar:6.87,rc:38.40,so1:10.95,so2:0,pr:0,sv:0,total:-25418.42,pct:-0.1201},actual:{jrts:0,ap:0,ar:0,rc:0,so1:0,so2:0,pr:0,sv:0,total:0,pct:0}},
  {week:26,date:'25 มิ.ย.',dayNum:25,target:{jrts:-25551.44,ap:76.80,ar:6.87,rc:38.40,so1:10.95,so2:0,pr:0,sv:0,total:-25418.42,pct:-0.1201},actual:{jrts:0,ap:0,ar:0,rc:0,so1:0,so2:0,pr:0,sv:0,total:0,pct:0}},
  {week:26,date:'26 มิ.ย.',dayNum:26,target:{jrts:-25551.44,ap:76.80,ar:6.87,rc:38.40,so1:10.95,so2:0,pr:0,sv:0,total:-25418.42,pct:-0.1201},actual:{jrts:0,ap:0,ar:0,rc:0,so1:0,so2:0,pr:0,sv:0,total:0,pct:0}},
  {week:26,date:'27 มิ.ย.',dayNum:27,holiday:true,target:{jrts:0,ap:76.80,ar:6.87,rc:38.40,so1:10.95,so2:0,pr:0,sv:0,total:133.02,pct:0.0006},actual:{jrts:0,ap:0,ar:0,rc:0,so1:0,so2:0,pr:0,sv:0,total:0,pct:0}},
  {week:26,date:'28 มิ.ย.',dayNum:28,holiday:true,target:{jrts:0,ap:76.80,ar:6.87,rc:38.40,so1:10.95,so2:0,pr:0,sv:0,total:133.02,pct:0.0006},actual:{jrts:0,ap:0,ar:0,rc:0,so1:0,so2:0,pr:0,sv:0,total:0,pct:0}},
];

const FIELDS=['jrts','ap','ar','rc','so1','so2','pr','sv','total'];
const FIELD_LABELS={jrts:'JRTS',ap:'Asso Product',ar:'Asso Refer',rc:'Reciprocation',so1:'Shop Online 1',so2:'Shop Online 2',pr:'Project',sv:'Service',total:'Total'};
function isHoliday(d){return d.holiday===true}
function fmt(v,dec=2){return v===0?'0':v.toLocaleString('th-TH',{minimumFractionDigits:dec,maximumFractionDigits:dec})}
function pctFmt(v){return(v*100).toFixed(2)+'%'}
function gapColor(g){if(g===0)return{cls:'gap-neu-cell',prefix:'—'};return g>0?{cls:'gap-good-cell',prefix:'▲ +'}:{cls:'gap-bad-cell',prefix:'▼ '}}
function actColor(a,t){if(a===t)return'act-equal';return a>t?'act-better':'act-worse'}
function blink(a,t){if(a===t||t===0&&a===0)return'';return a>t?'<span class="light light-green"></span>':'<span class="light light-red"></span>'}

function showTab(name,e){
  document.querySelectorAll('.tab-content').forEach(el=>el.classList.remove('active'));
  document.querySelectorAll('.tab-btn').forEach(el=>el.classList.remove('active'));
  document.getElementById('tab-'+name).classList.add('active');
  if(e&&e.target)e.target.classList.add('active');
}

function buildBreakdown(){
  const el=document.getElementById('breakdownGrid');
  const today=TODAY_DAYNUM;
  const pastDays=DAYS.filter(d=>d.dayNum<today);
  ['jrts','ap','ar','rc','so1','so2','pr','sv'].forEach(f=>{
    let tgt=0,act=0;
    pastDays.forEach(d=>{tgt+=d.target[f];act+=d.actual[f]});
    const gap=act-tgt;const gc=gapColor(gap);const ac=actColor(act,tgt);
    el.innerHTML+=`<div class="bd-card"><div class="bd-name">${FIELD_LABELS[f]}</div><div class="bd-target">Target: ${fmt(tgt)}</div><div class="bd-actual ${ac}">${blink(act,tgt)}${fmt(act)}</div><span class="kpi-gap ${gc.cls.replace('-cell','')}" style="font-size:10px">${gc.prefix}${fmt(Math.abs(gap))}</span></div>`;
  });
}

function buildWeeklyTable(){
  const weeks=[23,24,25,26];
  const wr={23:'1–7 มิ.ย.',24:'8–14 มิ.ย.',25:'15–21 มิ.ย.',26:'22–28 มิ.ย.'};
  const el=document.getElementById('weeklyTable');
  let h=`<thead><tr><th>สัปดาห์</th><th>ช่วงวันที่</th><th>ประเภท</th>`;
  FIELDS.forEach(f=>h+=`<th>${FIELD_LABELS[f]}</th>`);
  h+=`<th>%MA/Rev</th></tr></thead><tbody>`;
  weeks.forEach((w,wi)=>{
    const wDays=DAYS.filter(d=>d.week===w);
    const tgt={},act={};FIELDS.forEach(f=>{tgt[f]=0;act[f]=0});
    let tp=0,ap2=0,cnt=0;
    wDays.forEach(d=>{FIELDS.forEach(f=>{tgt[f]+=d.target[f];act[f]+=d.actual[f]});tp+=d.target.pct;ap2+=d.actual.pct;cnt++;});
    tp/=Math.max(cnt,1);ap2/=Math.max(cnt,1);
    const bdr=wi>0?'border-top:2px solid #dbeafe':'';
    h+=`<tr class="row-target" style="${bdr}"><td>${w}</td><td>${wr[w]}</td><td>Target</td>`;
    FIELDS.forEach(f=>h+=`<td>${fmt(tgt[f])}</td>`);h+=`<td>${pctFmt(tp)}</td></tr>`;
    h+=`<tr class="row-actual"><td></td><td></td><td>Actual</td>`;
    FIELDS.forEach(f=>h+=`<td class="${actColor(act[f],tgt[f])}">${blink(act[f],tgt[f])}${fmt(act[f])}</td>`);
    h+=`<td class="${actColor(ap2,tp)}">${blink(ap2,tp)}${pctFmt(ap2)}</td></tr>`;
    h+=`<tr class="row-gap"><td></td><td></td><td>Gap</td>`;
    FIELDS.forEach(f=>{const g=act[f]-tgt[f];const gc=gapColor(g);h+=`<td class="${gc.cls}">${gc.prefix}${fmt(Math.abs(g))}</td>`;});
    const gp=ap2-tp;const gc2=gapColor(gp);h+=`<td class="${gc2.cls}">${gc2.prefix}${Math.abs(gp*100).toFixed(2)}%</td></tr>`;
  });
  h+=`</tbody>`;el.innerHTML=h;
}

function buildCompTable(){
  const el=document.getElementById('compTable');
  const refActual={jrts:-535443.11,ap:3231.94,ar:0,rc:0,so1:0,so2:0,pr:0,sv:0,total:-532211.17,pct:-0.0898};
  const curTarget={jrts:-485477.30,ap:2150.38,ar:76.92,rc:134.08,so1:132.96,so2:173.56,pr:0,sv:0,total:-482809.39,pct:-0.0815};
  const curActual={jrts:-45714.18,ap:0,ar:0,rc:0,so1:-184.72,so2:1218.11,pr:0,sv:0,total:-44680.79,pct:-0.0075};
  const rows=[
    {label:'Actual สัปดาห์ 10–13 (มี.ค. 2569)',period:'2–29 มี.ค. 2569',data:refActual,cls:'row-target'}, // ⚠️ ปี พ.ศ. ต้องคำนวณจาก ค.ศ.จริง+543 เสมอ ห้ามลอกเลขปีนี้ตรงๆ
    {label:'Target สัปดาห์ 23–26 (มิ.ย. 2569)',period:'1–28 มิ.ย. 2569',data:curTarget,cls:'row-actual'},
    {label:'Actual สะสม (ถึง 10 มิ.ย. 2569)',period:'1–10 มิ.ย. 2569',data:curActual,cls:'row-gap'},
  ];
  let h=`<thead><tr><th>รายการ</th><th>ช่วงเวลา</th>`;
  FIELDS.forEach(f=>h+=`<th>${FIELD_LABELS[f]}</th>`);
  h+=`<th>%MA/Rev</th></tr></thead><tbody>`;
  rows.forEach(r=>{
    h+=`<tr class="${r.cls}"><td style="text-align:left;font-size:11px">${r.label}</td><td>${r.period}</td>`;
    FIELDS.forEach(f=>h+=`<td>${fmt(r.data[f])}</td>`);h+=`<td>${pctFmt(r.data.pct)}</td></tr>`;
  });
  h+=`</tbody>`;el.innerHTML=h;
}

function buildLineChart(canvasId){
  const labels=DAYS.map(d=>d.date);
  const tData=DAYS.map(d=>parseFloat((d.target.pct*100).toFixed(2)));
  const aData=DAYS.map(d=>parseFloat((d.actual.pct*100).toFixed(2)));
  new Chart(document.getElementById(canvasId),{type:'line',data:{labels,datasets:[
    {label:'Target',data:tData,borderColor:'#5b9bd5',backgroundColor:'transparent',borderWidth:1.5,pointRadius:0,tension:0,spanGaps:true},
    {label:'Actual',data:aData,borderColor:'#ed7d31',backgroundColor:'transparent',borderWidth:1.5,pointRadius:0,tension:0,spanGaps:true}
  ]},options:{responsive:true,maintainAspectRatio:false,plugins:{
    title:{display:true,text:'กราฟแสดงการเปรียบเทียบ Target และ Actual Total MA Back Office',font:{size:13,family:"'Sarabun',sans-serif"},color:'#374151',padding:{bottom:12}},
    legend:{display:true,position:'top',labels:{usePointStyle:true,pointStyle:'line',boxWidth:30,font:{family:"'Sarabun',sans-serif"}}},
    tooltip:{callbacks:{label:ctx=>`${ctx.dataset.label}: ${ctx.parsed.y.toFixed(2)}%`}}
  },scales:{
    x:{grid:{display:true,color:'#e5e7eb'},ticks:{font:{size:9,family:"'Sarabun',sans-serif"},maxRotation:45,autoSkip:false}},
    y:{grid:{display:true,color:'#e5e7eb'},ticks:{font:{size:10,family:"'Sarabun',sans-serif"},callback:v=>v.toFixed(2)+'%'}}
  }}});
}

function buildBarChart(){
  new Chart(document.getElementById('barChart'),{type:'bar',data:{
    labels:['JRTS','Asso Prod','Asso Ref','Recipro','Shop Online 1','Shop Online 2','Project','Service'],
    datasets:[
      {label:'Target',data:[-485477.30,2150.38,76.92,134.08,132.96,173.56,0,0],backgroundColor:'rgba(91,155,213,0.7)',borderColor:'#5b9bd5',borderWidth:1},
      {label:'Actual',data:[-45714.18,0,0,0,-184.72,1218.11,0,0],backgroundColor:'rgba(237,125,49,0.7)',borderColor:'#ed7d31',borderWidth:1}
    ]
  },options:{responsive:true,maintainAspectRatio:false,
    plugins:{legend:{position:'top',labels:{font:{family:"'Sarabun',sans-serif"}}},tooltip:{callbacks:{label:ctx=>`${ctx.dataset.label}: ${ctx.parsed.y.toLocaleString('th-TH',{minimumFractionDigits:2})}`}}},
    scales:{x:{ticks:{font:{size:9,family:"'Sarabun',sans-serif"},maxRotation:30}},y:{ticks:{font:{size:9,family:"'Sarabun',sans-serif"},callback:v=>v.toLocaleString('th-TH',{notation:'compact'})}}}}});
}

function buildDailyTable(){
  const el=document.getElementById('dailyTable');
  const todayDay=TODAY_DAYNUM;
  let h=`<thead><tr><th>สัปดาห์</th><th>วันที่</th><th>ประเภท</th>`;
  FIELDS.forEach(f=>h+=`<th>${FIELD_LABELS[f]}</th>`);
  h+=`<th>%MA/Rev</th></tr></thead><tbody>`;
  let prevWeek=null;
  DAYS.forEach(d=>{
    const wc=d.week!==prevWeek?d.week:'';prevWeek=d.week;
    const dl=isHoliday(d)?`${d.date}<span class="holiday-badge">วันหยุด</span>`:d.date;
    const rc=isHoliday(d)?'row-holiday':'row-target';
    const isPast=d.dayNum<todayDay;
    if(isPast){
      h+=`<tr class="${rc}"><td>${wc}</td><td>${dl}</td><td>Target</td>`;
      FIELDS.forEach(f=>h+=`<td>${fmt(d.target[f])}</td>`);h+=`<td>${pctFmt(d.target.pct)}</td></tr>`;
      h+=`<tr class="row-actual"><td></td><td></td><td>Actual</td>`;
      FIELDS.forEach(f=>h+=`<td class="${actColor(d.actual[f],d.target[f])}">${blink(d.actual[f],d.target[f])}${fmt(d.actual[f])}</td>`);
      h+=`<td class="${actColor(d.actual.pct,d.target.pct)}">${blink(d.actual.pct,d.target.pct)}${pctFmt(d.actual.pct)}</td></tr>`;
      h+=`<tr class="row-gap"><td></td><td></td><td>Gap</td>`;
      FIELDS.forEach(f=>{const g=d.actual[f]-d.target[f];const gc=gapColor(g);h+=`<td class="${gc.cls}">${gc.prefix}${fmt(Math.abs(g))}</td>`;});
      const gp=d.actual.pct-d.target.pct;const gc2=gapColor(gp);
      h+=`<td class="${gc2.cls}">${gc2.prefix}${Math.abs(gp*100).toFixed(2)}%</td></tr>`;
    } else {
      h+=`<tr class="${rc}"><td>${wc}</td><td>${dl}</td><td>Target</td>`;
      FIELDS.forEach(f=>h+=`<td>${fmt(d.target[f])}</td>`);h+=`<td>${pctFmt(d.target.pct)}</td></tr>`;
      h+=`<tr class="nodata-row"><td></td><td></td><td colspan="${FIELDS.length+1}" style="text-align:center">ยังไม่มีข้อมูล</td></tr>`;
    }
  });
  h+=`</tbody>`;el.innerHTML=h;
}

function buildStackedBar(){
  const todayDay=TODAY_DAYNUM;
  const pastDays=DAYS.filter(d=>d.dayNum<todayDay);
  const labels=pastDays.map(d=>isHoliday(d)?d.date+' 🏖':d.date);
  const flds=['jrts','ap','ar','rc','so1','so2','pr','sv'];
  const colors=['#60a5fa','#34d399','#a78bfa','#f87171','#fbbf24','#f59e0b','#94a3b8','#f9a8d4'];
  new Chart(document.getElementById('stackedBar'),{type:'bar',data:{labels,datasets:flds.map((f,i)=>({
    label:FIELD_LABELS[f],data:pastDays.map(d=>d.actual[f]),backgroundColor:colors[i],stack:'actual'
  }))},options:{responsive:true,maintainAspectRatio:false,
    plugins:{legend:{position:'top',labels:{font:{family:"'Sarabun',sans-serif"},boxWidth:12}}},
    scales:{x:{stacked:true,ticks:{font:{size:10,family:"'Sarabun',sans-serif"}}},y:{stacked:true,ticks:{font:{size:10,family:"'Sarabun',sans-serif"},callback:v=>v.toLocaleString('th-TH',{notation:'compact'})}}}}});
}

// ⚠️ GROSS_MA_RATIO เป็นค่าคงที่ตายตัว ห้ามเปลี่ยน — ใช้แปลง Gap บาท → Net Revenue บาท เท่านั้น
// ห้ามแปลงเป็นนาที/ห้ามหารด้วย Department rate อีกต่อไป
const GROSS_MA_RATIO = 0.66;

// ⚠️ BOM_TASK: งานที่มอบหมายในกรณีวันนี้ไม่ใช่วันหยุด ใช้ได้ "เฉพาะ BOM-PCRC-0001-A1 เท่านั้น"
// (ดึงคำอธิบายจริงจากชีต 3.BOM มาแทนค่า desc — ห้ามใช้ BOM No. อื่นในกรณีนี้)
const BOM_TASK = {id:'BOM-PCRC-0001-A1', desc:'โทรติดตามลูกค้า Association เพื่อซื้อสินค้า'};

// ⚠️ PRIYANAT_NAME / PRIYANAT_TASK: ใช้เฉพาะกรณีวันนี้ (วันที่มอบหมายงาน) ตรงกับวันหยุด — มอบหมาย
// ให้คนนี้คนเดียวเท่านั้น ด้วยงานนี้เท่านั้น ห้ามใช้ BOM No. อื่น ห้ามให้คนอื่นรับงานร่วมในวันหยุด
const PRIYANAT_NAME = 'ปริยณัฐ รุจิชัยวิสิฐ';
const PRIYANAT_TASK = {id:'BOM-PCRC-0001-D1', desc:'ติดตามลูกค้าใน Facebook (Lead, no show, repurchase)'};

// ⚠️ ASSIGNMENT_WHITELIST: รายชื่อพนักงานเดียวที่มีสิทธิ์รับมอบหมายงานในกรณี B (วันทำงาน) — ค่าคงที่
// ตายตัว ห้ามเปลี่ยน/เพิ่ม/ลด ห้ามมอบหมายให้ใครนอกรายชื่อนี้ไม่ว่า Available Time จะสูงแค่ไหนก็ตาม
// (ไม่เกี่ยวกับกรณี A ซึ่งมอบหมายให้ปริยณัฐ รุจิชัยวิสิฐ คนเดียวตามเดิมเสมอ)
const ASSIGNMENT_WHITELIST = [
  'วิลาวรรณ์ ก้อนกลิ่น','หนึ่งนภา ตามระเบียบ','กรกนก อรรถวิลัย','ศุภลักษณ์ จั่นเพ็ชร์',
  'เรวัตร เจริญวงษ์','วิจิตรา พันธ์แสง','วัททลิกา สิริมิ่งขวัญ','ชุติมา ทิพยกุล',
  'กัญญรส น้อยดัด','ธาวิน เภาทอง','อนุพงศ์ แสนกล้า','เข็มพร พรมราช','ศิริลักษณ์ ภู่สวรรค์ชัย'
];

// ⚠️ ELIGIBLE_RANKING: แทนค่าด้วยผลคำนวณ Available Time จริงของวันนี้ (ขั้นตอนที่ 3) เรียงมากไปน้อย
// เฉพาะคนที่มีแถว Planning วันนี้และ JobType="1-Office" เท่านั้น — คำนวณใหม่ทุกครั้ง ห้ามใช้ list ตัวอย่างนี้ซ้ำ
// ⚠️ ต้องไม่มี "ปริยณัฐ รุจิชัยวิสิฐ" อยู่ใน list นี้เด็ดขาด — เธอรับงานเฉพาะกรณีวันหยุดผ่าน PRIYANAT_NAME/
// PRIYANAT_TASK ด้านบนเท่านั้น ไม่ใช่ผ่าน ranking นี้
// ⚠️ ทุกคนใน list นี้ต้องอยู่ใน ASSIGNMENT_WHITELIST เท่านั้น — คนที่ไม่อยู่ใน whitelist ต้องถูกตัดออก
// ก่อนจัดอันดับ แม้ Available Time จะสูงที่สุดในแผนกก็ตาม (join ด้วยรหัสพนักงาน ไม่ใช่ชื่อดิบ)
// ⚠️ ranking นี้ใช้ "จัดอันดับ" เท่านั้น — มอบหมายงานจริงให้แค่ ELIGIBLE_RANKING[0] (อันดับ 1) คนเดียวเสมอ
const ELIGIBLE_RANKING = [
  {id:'51320116',name:'สรวิชญ์ เสมอเชื้อ',avail:300.0},
  {id:'51320013',name:'วิจิตรา พันธ์แสง',avail:276.2},
  {id:'51320084',name:'เข็มพร พรมราช',avail:274.4},
];

// ⚠️ EXCLUDED_NOTE: ระบุชื่อพนักงานที่ถูกกันออกเพราะไม่มีแถว Planning วันนี้ (ตรวจ JobType ไม่ได้)
// ไม่ต้องระบุคนที่ถูกกันออกเพราะไม่อยู่ใน ASSIGNMENT_WHITELIST (เป็นกฎปกติ ไม่ใช่ข้อยกเว้นที่ต้องแจ้ง)
// ถ้าไม่มีใครถูกกันออก ให้ใส่ค่าว่าง ''
const EXCLUDED_NOTE = '';

function buildAssignmentPanel(){
  const el=document.getElementById('assignmentPanel');
  const yesterdayDay=TODAY_DAYNUM-1;
  const latestDay=DAYS.find(d=>d.dayNum===yesterdayDay);
  if(!latestDay){el.innerHTML='';return}

  const actualPct=latestDay.actual.pct;
  const targetPct=latestDay.target.pct;

  if(actualPct>=targetPct){
    el.innerHTML=`<div class="assignment-panel all-good"><h3>✅ %MA/Revenue เมื่อวาน (${latestDay.date}) = ${(actualPct*100).toFixed(2)}% ดีกว่าหรือเท่ากับ Target (${(targetPct*100).toFixed(2)}%) — ไม่มีการมอบหมายงานเพิ่มเติม</h3></div>`;
    return;
  }

  // ยอดที่ขาด (Gross MA Gap บาท) และยอดที่ต้องทำให้ได้ (Net Revenue เพื่อปิด Gap) — คำนวณเป็นเงินเท่านั้น
  const gapBaht=latestDay.target.total-latestDay.actual.total;
  const netRevenueNeeded=gapBaht/GROSS_MA_RATIO;

  // วันที่มอบหมายงาน = วันนี้ (TODAY_DAYNUM) ไม่ใช่เมื่อวาน — holiday flag ของ "วันนี้" เป็นตัวตัดสิน
  // ว่าจะมอบหมายแบบไหน (กรณี A: ปริยณัฐคนเดียว / กรณี B: อันดับ 1 ใน ELIGIBLE_RANKING คนเดียว)
  const todayEntry=DAYS.find(d=>d.dayNum===TODAY_DAYNUM);
  const assignDayIsHoliday=todayEntry?isHoliday(todayEntry):false;
  const assignDayLabel=todayEntry?todayEntry.date:'';
  const holidayBadge=assignDayIsHoliday?` <span style="font-size:12px;background:#fef3c7;color:#92400e;border-radius:4px;padding:2px 6px;font-weight:600">วันหยุด</span>`:'';

  // ⚠️ มอบหมายให้ "คนเดียว" เสมอ ไม่ว่ากรณีไหน — รับยอดที่ขาด/ยอดที่ต้องทำให้ได้เต็มจำนวนทั้งหมด
  let assigneeName, taskId, taskDesc, noEligible=false;

  if(assignDayIsHoliday){
    // กรณี A — วันนี้ตรงกับวันหยุด: มอบหมายให้ปริยณัฐ รุจิชัยวิสิฐ คนเดียว ใช้ BOM-D1 เท่านั้น
    assigneeName=PRIYANAT_NAME;
    taskId=PRIYANAT_TASK.id;
    taskDesc=PRIYANAT_TASK.desc;
  } else {
    // กรณี B — วันนี้ไม่ใช่วันหยุด: มอบหมายให้อันดับ 1 ใน ELIGIBLE_RANKING คนเดียว ใช้ BOM-A1 เท่านั้น
    const top=ELIGIBLE_RANKING[0];
    if(!top){noEligible=true;}
    assigneeName=top?top.name:'';
    taskId=BOM_TASK.id;
    taskDesc=BOM_TASK.desc;
  }

  let h=`<div class="assignment-panel"><h3>📋 การมอบหมายงานวันนี้ (${assignDayLabel})${holidayBadge} — %MA เมื่อวาน (${latestDay.date}) = ${(actualPct*100).toFixed(2)}% ต่ำกว่า Target ${(targetPct*100).toFixed(2)}%</h3>
  <div class="assignment-note">
    ยอดที่ขาด (Gross MA Gap รวมทั้งวัน) <b>${fmt(gapBaht)} บาท</b> → ยอดที่ต้องทำให้ได้เพื่อปิด Gap (Net Revenue) <b>${fmt(netRevenueNeeded)} บาท</b> (÷ Gross MA Ratio 66%)<br>
    ${assignDayIsHoliday?('วันนี้เป็นวันหยุด — มอบหมายให้ '+PRIYANAT_NAME+' คนเดียวตามกฎพิเศษวันหยุด'):'มอบหมายให้พนักงานอันดับ 1 ใน Available Time Ranking คนเดียวรับผิดชอบยอดทั้งหมด'}
  </div>
  <div class="assignment-grid-inner">`;

  if(noEligible){
    h+=`<p style="font-size:12px;color:#dc2626">⚠️ ไม่มีพนักงานที่มีแถว Planning วันนี้และ JobType="1-Office" ให้มอบหมายงาน</p>`;
  } else {
    h+=`<div class="task-card"><div class="task-header"><span class="bom-id">${taskId}</span></div>
    <div class="task-desc">${taskDesc}</div>
    <div class="assigned-emp">ผู้ได้รับมอบหมาย: ${assigneeName}</div>
    <div class="hc-info">ยอดที่ขาด: ${fmt(gapBaht)} บาท</div>
    <div class="hc-info">ยอดที่ต้องทำให้ได้: ${fmt(netRevenueNeeded)} บาท</div></div>`;
  }

  h+=`</div>`;
  if(!assignDayIsHoliday && EXCLUDED_NOTE){
    h+=`<p style="font-size:11px;color:#9ca3af;margin-top:10px">${EXCLUDED_NOTE}</p>`;
  }
  h+=`</div>`;
  el.innerHTML=h;
}

buildBreakdown();
buildWeeklyTable();
buildCompTable();
buildDailyTable();
buildLineChart('maLineChart');
buildLineChart('maLineChart2');
buildBarChart();
buildStackedBar();
buildAssignmentPanel();
</script>
</body>
</html>

```
