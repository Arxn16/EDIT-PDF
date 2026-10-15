FROM python:3.12-slim

# Tesseract (OCR ไทย/อังกฤษ) + LibreOffice (แปลง Word/Excel/PowerPoint เป็น PDF) + ฟอนต์ไทย
RUN apt-get update \
 && apt-get install -y --no-install-recommends \
      tesseract-ocr tesseract-ocr-tha tesseract-ocr-eng libglib2.0-0 \
      libreoffice-writer-nogui libreoffice-calc-nogui libreoffice-impress-nogui \
      fonts-thai-tlwg fontconfig \
 && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt \
 && python -c "from pythainlp.corpus import tnc, thai_words; tnc.word_freqs(); thai_words()"

# ฟอนต์ "Sarabun TH Compat" (ขนาดเท่า TH Sarabun New / TH SarabunIT๙) ใช้แทนเมื่อไม่มีฟอนต์จริง
COPY static/fonts/SarabunTHCompat-*.ttf /usr/share/fonts/truetype/thai-compat/
COPY fonts/thai-aliases.conf /etc/fonts/conf.d/50-thai-aliases.conf
RUN fc-cache -f

COPY converter.py docx_fix.py thai_ocr_fix.py thai_pdf_text.py editor.py app.py entrypoint.sh ./
COPY templates ./templates
COPY static ./static
COPY tests ./tests
COPY tools ./tools
RUN chmod +x entrypoint.sh

ENV PYTHONUNBUFFERED=1 MAX_UPLOAD_MB=200 HOME=/tmp
EXPOSE 8080
ENTRYPOINT ["./entrypoint.sh"]
