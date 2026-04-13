# 3 TIME POS — GitHub Deploy Ready

ชุดนี้ถูกจัดไฟล์ใหม่ให้พร้อมใช้กับ GitHub + Google Apps Script (`clasp`) แล้ว

## โครงสร้างไฟล์
- `Code.gs`
- `Index.html`
- `Customer.html`
- `KitchenDisplay.html`
- `appsscript.json`
- `.clasp.json.example`
- `.claspignore`
- `package.json`
- `.github/workflows/deploy-appsscript.yml`
- `PRINTER_BRIDGE_GUIDE.md`

## วิธีใช้งานแบบเร็ว
1. สร้าง GitHub repository แล้วอัปโหลดไฟล์ทั้งหมดชุดนี้
2. ติดตั้ง Node.js 20+ บนเครื่องคุณ
3. ติดตั้ง clasp
   ```bash
   npm install
   npm run login
   ```
4. สร้าง Apps Script project ใหม่ หรือ clone ของเดิม
   ```bash
   npm run create
   ```
   หรือ
   ```bash
   npx clasp clone YOUR_SCRIPT_ID
   ```
5. แก้ไฟล์ `.clasp.json` จาก `.clasp.json.example`
6. push ขึ้น Apps Script
   ```bash
   npm run push
   ```
7. เปิด Apps Script แล้ว Deploy เป็น Web App

## GitHub Actions
ถ้าต้องการ deploy จาก GitHub อัตโนมัติ ให้ตั้งค่า GitHub Secrets ดังนี้
- `CLASPRC_JSON` = เนื้อหาจากไฟล์ `~/.clasprc.json`
- `CLASP_JSON` = เนื้อหาจากไฟล์ `.clasp.json`

จากนั้น push branch `main` ระบบจะรัน workflow ให้

## คำสั่งใช้งาน
```bash
npm run login
npm run create
npm run push
npm run deploy
```

## หมายเหตุ
- ต้องเปิด Apps Script API ในบัญชี Google ก่อนใช้ `clasp`
- ถ้าเป็น Web App ต้องตั้ง deployment ครั้งแรกใน Apps Script editor อย่างน้อยหนึ่งครั้ง
- `KitchenDisplay.html` ถูกแยกไฟล์ไว้แล้ว แต่ใน `doGet(e)` ยังอ้างชื่อไฟล์เดิมถูกต้อง สามารถ push ได้ทันที
