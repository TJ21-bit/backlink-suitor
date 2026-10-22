# Backlink Auditor — container for Render (or any Docker host)
FROM python:3.12-slim

ENV PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    PLAYWRIGHT_BROWSERS_PATH=/ms-playwright

WORKDIR /app

COPY requirements.txt .
RUN pip install -r requirements.txt \
 && python -m playwright install --with-deps chromium

COPY . .

# One worker process so every request sees the same in-memory job list;
# threads handle concurrent users. Render provides $PORT.
CMD gunicorn backend_server:app --workers 1 --threads 8 --timeout 180 --bind 0.0.0.0:${PORT:-8080}
