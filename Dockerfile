FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    SOFFICE_PATH=soffice

# LibreOffice is required by app.py for PPT/PPTX -> PDF conversion.
RUN apt-get update \
    && apt-get install -y --no-install-recommends libreoffice-impress fonts-noto-cjk \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt ./
RUN pip install -r requirements.txt

COPY . .

# Runtime-generated directories should exist and be writable.
RUN mkdir -p uploads tmp_convert static/slides data

EXPOSE 5000
CMD ["sh", "-c", "gunicorn --bind 0.0.0.0:${PORT:-5000} --workers 1 --timeout 180 app:app"]
