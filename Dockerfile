FROM python:3.12-slim

ENV PYTHONUNBUFFERED=1 \
    PY_NO_PAUSE=1 \
    MAX_UPLOAD_MB=900 \
    TESSDATA_PREFIX=/usr/share/tesseract-ocr/5/tessdata

RUN apt-get update \
    && apt-get install -y --no-install-recommends tesseract-ocr tesseract-ocr-por \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements-web.txt ./
RUN pip install --no-cache-dir -r requirements-web.txt

COPY Consolidado_Emendas.py Consolidado_Emendas_V2.py web_app.py ./
COPY "Planilha_emendas_pl4_2025 atualizada até 09.04.2026.xlsx" ./

RUN mkdir -p /app/_web_outputs

EXPOSE 10000
CMD ["sh", "-c", "gunicorn --bind 0.0.0.0:${PORT:-10000} --timeout 1800 --workers 1 web_app:app"]
