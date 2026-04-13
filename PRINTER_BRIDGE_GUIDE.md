# Printer Bridge Guide

ไฟล์ `BRIDGE_AGENT_SAMPLE.py` เป็นตัวอย่าง agent ภายนอกสำหรับต่อ printer จริง

แนวทางใช้งาน
1. Deploy Apps Script เป็น Web App
2. เปิดหน้า Settings ในระบบ แล้วเปิด `Printer Bridge`
3. ตั้ง `Bridge Token` และ `Printer เริ่มต้น`
4. นำ Web App URL ไปใส่ใน `BRIDGE_AGENT_SAMPLE.py`
5. ติดตั้ง Python และไลบรารี `requests`
6. ปรับฟังก์ชัน `print_job()` ให้ส่งงานเข้า printer จริง เช่น ESC/POS, Windows spooler หรือ CUPS
7. รัน agent บนเครื่องที่ต่อ printer

หมายเหตุ
- เวอร์ชันนี้ใช้คิว `print_queue` เป็นแกน
- Agent จะ claim งานที่ status = `queued`
- เมื่อพิมพ์เสร็จจะ ack กลับไปที่ระบบ
