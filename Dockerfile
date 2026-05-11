FROM python:3.11-slim

WORKDIR /app

# Dependencias del sistema para python-docx y openpyxl
RUN apt-get update && apt-get install -y --no-install-recommends \
    gcc \
    libxml2 \
    libxslt1.1 \
 && rm -rf /var/lib/apt/lists/*

# Instalar dependencias Python primero (mejor cache de Docker)
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copiar TODOS los módulos Python del repo
COPY *.py ./

# Directorio de outputs (DOCX/XLSX generados temporalmente)
ENV ZEBRA_OUTPUT_DIR=/tmp/zebra
RUN mkdir -p /tmp/zebra

EXPOSE 8080

HEALTHCHECK --interval=30s --timeout=5s --start-period=10s --retries=3 \
  CMD python -c "import urllib.request; urllib.request.urlopen('http://localhost:8080/health').read()" || exit 1

CMD ["uvicorn", "zebra_api:app", "--host", "0.0.0.0", "--port", "8080"]
