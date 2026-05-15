FROM python:3.11.18-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    DEBIAN_FRONTEND=noninteractive

WORKDIR /app

RUN apt-get update \
    && apt-get install -y --no-install-recommends \
        fontconfig \
        fonts-dejavu-core \
        libgdiplus \
        libglib2.0-0 \
        libgomp1 \
        libicu-dev \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN python -m pip install --no-cache-dir -r requirements.txt

COPY . .

EXPOSE 8000

CMD ["uvicorn", "pdf_table_extractor.app:app", "--host", "0.0.0.0", "--port", "8000"]
