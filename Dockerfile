# Weekly Report Formatter — Docker image for Render
# Required for Tesseract OCR (Opinionn Review Summary PDF reading)
# since Render's free tier ignores Aptfile.

FROM python:3.11-slim

# Install system packages: tesseract for OCR, poppler-utils for pdf2image
RUN apt-get update && apt-get install -y --no-install-recommends \
    tesseract-ocr \
    poppler-utils \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Install Python dependencies first (cached layer if requirements don't change)
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy app code
COPY app.py .

# Render sets PORT env var; app.py reads it. Default to 10000 for local Docker runs.
ENV PORT=10000
EXPOSE 10000

# Use gunicorn for production (more reliable than Flask dev server on Render)
RUN pip install --no-cache-dir gunicorn
CMD gunicorn --bind 0.0.0.0:$PORT --timeout 120 --workers 1 app:app
