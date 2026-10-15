FROM python:3.12-slim
ENV PYTHONDONTWRITEBYTECODE=1 PYTHONUNBUFFERED=1 PORT=10000
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-writer libreoffice-calc libreoffice-impress \
    tesseract-ocr tesseract-ocr-eng tesseract-ocr-urd tesseract-ocr-spa tesseract-ocr-fra tesseract-ocr-deu \
    ghostscript fonts-dejavu-core fonts-liberation \
    && rm -rf /var/lib/apt/lists/*
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY app.py .
COPY static ./static
EXPOSE 10000
CMD ["sh", "-c", "waitress-serve --listen=0.0.0.0:${PORT:-10000} --threads=4 app:app"]
