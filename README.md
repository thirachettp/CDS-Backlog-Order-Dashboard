# CDS Backlog Dashboard v2 (Streamlit)

## โครงสร้างไฟล์
```
cds_dashboard/
├── app.py              ← หน้าแอปหลัก (Dashboard + Settings)
├── storage.py          ← จัดการบันทึก/โหลดข้อมูลถาวรจาก disk
├── cutoff_parser.py    ← อ่านไฟล์ Stock Cut-off เฉพาะชีท "CDS"
├── requirements.txt
└── data/                ← โฟลเดอร์เก็บไฟล์ล่าสุด + การตั้งค่า (สร้างอัตโนมัติ)
```

## วิธีรัน
```bash
pip install -r requirements.txt
streamlit run app.py
```
เปิดเบราว์เซอร์ที่ http://localhost:8501

**สำคัญ:** ทั้งโฟลเดอร์ `cds_dashboard/` ต้องอยู่ด้วยกัน (app.py เรียกใช้ storage.py และ
cutoff_parser.py จากโฟลเดอร์เดียวกัน)

## การเก็บข้อมูลถาวร (Persistence)
ทุกครั้งที่อัปโหลดไฟล์ Order Status ใหม่ในหน้า Dashboard ระบบจะบันทึกสำเนาไว้ที่
`data/latest_order.xlsx` โดยอัตโนมัติ — เปิดแอปครั้งถัดไปโดยไม่อัปโหลดไฟล์ใหม่
ระบบจะโหลดไฟล์ล่าสุดที่เคยอัปโหลดไว้ให้เองพร้อมบอกวันเวลาที่อัปโหลด

ข้อมูลอื่นที่บันทึกถาวรเช่นกัน:
- `data/owner_mapping.json` — คนดูแลแต่ละ Sub Dept (ตั้งค่าในหน้า Settings)
- `data/latest_stock_cutoff.csv` — ตาราง Stock Cut-off ที่แปลงแล้ว

**ข้อควรระวัง:** ถ้า deploy บน Streamlit Community Cloud หรือ host ที่ filesystem
เป็นแบบชั่วคราว (ephemeral) ไฟล์ในโฟลเดอร์ `data/` จะหายไปเมื่อแอป restart/redeploy
แนะนำให้รันบนเครื่อง/เซิร์ฟเวอร์ของบริษัทที่มี disk ถาวร (เช่น ผ่าน Docker + volume mount)

## หน้า Settings (⚙️)
1. **จัดการคนดูแลแต่ละ Sub Dept** — ตารางแก้ไขได้ (data editor) ดึงรายการ Sub Dept
   จากไฟล์ Order Status ที่โหลดล่าสุดโดยอัตโนมัติ ใส่ชื่อคนดูแลแล้วกด "บันทึกคนดูแล"
2. **อัปโหลดไฟล์ Stock Cut-off** — อัปโหลดไฟล์ตารางแจ้งช่วงงดรับสินค้าระหว่างนับสต๊อก
   ระบบจะ**อ่านเฉพาะชีทชื่อ "CDS"** เท่านั้น แปลงช่วงวันที่ (เช่น "4-10/08/2026")
   เป็นวันที่เริ่ม-สิ้นสุดจริง แล้วบันทึกไว้ใช้ทั่วทั้งแดชบอร์ด
3. ดูข้อมูลไฟล์ที่บันทึกไว้ในระบบ และปุ่มล้างข้อมูลทั้งหมด

## หน้า Dashboard (📊) — 4 ส่วน
- **Part 1 · Filter & KPI Cards** — ตัวกรอง (Status, Brand, To Store, Order Type,
  คนดูแล, ช่วงวันที่) + KPI การ์ด 8 ใบ (เพิ่ม "ติด Stock Cutoff" จากไฟล์ที่อัปโหลดใน Settings)
- **Part 2 · Backlog ตามคนดูแล** — ตาราง pivot สไตล์เดียวกับตัวอย่างที่ส่งมา
  (CDS Over All + แยกตามคนดูแลแต่ละคน) แถว = สถานะ, คอลัมน์ = วันที่, มีคอลัมน์/แถว Total
  สามารถสลับค่าที่แสดงได้ระหว่าง "Required QTY (ผลรวม)" กับ "จำนวน Order (ไม่ซ้ำ)"
- **Part 3 · Backlog Items Details** — ตารางรายละเอียด ค้นหาได้ และปุ่ม Export เป็น **Excel**
- **Part 4 · Dashboard** — กราฟ Donut (by Status), Top 10 Store, และ Order Backlog รายวัน

## สมมติฐานสำคัญที่ใช้ (ปรับได้ในโค้ด)
- Total Backlog Orders / Required / Allocated / Picked QTY คำนวณจากแถวที่ Status
  เป็น "รอเบิก" หรือ "รอยืนยัน" เท่านั้น
- Shipped QTY = ผลรวม Pick QTY ของแถว backlog ที่ Is Shipped = "Y"
  (ไฟล์ต้นฉบับไม่มีคอลัมน์ Shipped QTY ตรงๆ)
- Part 2 ใช้สถานะจริงจากไฟล์ (รอเบิก / รอยืนยัน / Allocated / Allocated ShortAll /
  ปิดเอกสาร / ยกเลิก) แทนสถานะ Printed/In Packing/Pack Complete/Loaded on Truck
  ในภาพตัวอย่าง เพราะไฟล์ Order Status ที่มีไม่มีข้อมูลสถานะแพ็ก/โหลดขึ้นรถ
  — ถ้ามีไฟล์ที่มีสถานะเหล่านี้ ส่งมาได้เลย จะปรับให้ตรงตามจริง
- "ติด Stock Cutoff" เทียบ Start Ship Date ของแต่ละแถว กับช่วง CutFrom–CutTo
  ของ Store+Sub Dept เดียวกันจากไฟล์ Stock Cut-off ที่อัปโหลดไว้ใน Settings

## ทดสอบแล้ว
รันผ่าน `streamlit.testing.v1.AppTest` ทั้งหน้า Dashboard (พร้อมไฟล์ตัวอย่างจริง)
และหน้า Settings แล้วไม่มี exception
