FROM python:3.11-slim

# Zależności systemowe dla Playwright Chromium
RUN apt-get update && apt-get install -y \
    wget gnupg \
    libnss3 libnspr4 libatk1.0-0 libatk-bridge2.0-0 libcups2 \
    libdrm2 libxkbcommon0 libxcomposite1 libxdamage1 libxfixes3 \
    libxrandr2 libgbm1 libpango-1.0-0 libcairo2 libasound2 \
    fonts-liberation fonts-noto-color-emoji \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
RUN playwright install chromium
RUN playwright install-deps chromium

COPY server.py .
COPY templates/ ./templates/

ENV PORT=8000
EXPOSE 8000

# Gunicorn z 1 workerem (Playwright trzyma chromium subprocess, więcej workerów = więcej RAM)
# Timeout 60s żeby render miał czas
CMD ["gunicorn", "--bind", "0.0.0.0:8000", "--workers", "1", "--timeout", "60", "server:app"]
