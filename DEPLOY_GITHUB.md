# Deploy ผ่าน GitHub แบบย่อ

## 1) เตรียม local ครั้งแรก
```bash
npm install
npx clasp login
npx clasp create --type standalone --title "3 TIME POS" --rootDir .
```

หรือถ้ามี Script ID อยู่แล้ว
```bash
npx clasp clone YOUR_SCRIPT_ID
```

## 2) เปิด Apps Script API
ไปที่ Apps Script settings แล้วเปิด Apps Script API

## 3) ตั้ง GitHub Secrets
- `CLASPRC_JSON` = ไฟล์ `~/.clasprc.json`
- `CLASP_JSON` = ไฟล์ `.clasp.json`

## 4) Push ขึ้น GitHub
เมื่อ push เข้า branch `main` workflow จะ `clasp push -f` ไปยัง Apps Script project

## 5) Deploy Web App
หลัง push สำเร็จ ให้ไปที่ Apps Script > Deploy > New deployment > Web app

ถ้าต้องการให้ผมทำเวอร์ชันที่ผูก Script ID ให้เลยด้วย คุณต้องส่ง Script ID มาด้วย
