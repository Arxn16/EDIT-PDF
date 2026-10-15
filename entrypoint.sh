#!/bin/sh
# ถ้าผู้ใช้วางฟอนต์ของตัวเองไว้ในโฟลเดอร์ my-fonts (เช่น THSarabunNew.ttf) ให้ระบบรู้จัก
if [ -d /usr/share/fonts/truetype/user ] && [ "$(ls -A /usr/share/fonts/truetype/user 2>/dev/null)" ]; then
  fc-cache -f >/dev/null 2>&1
fi
# timeout ยาว เพราะ OCR / แปลงไฟล์หลายหน้าใช้เวลานาน
exec gunicorn -w 2 --threads 4 -t 1800 -b 0.0.0.0:8080 app:app
