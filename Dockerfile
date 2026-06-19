FROM python:3.12-slim

WORKDIR /app

# Dependencias del sistema
RUN apt-get update && apt-get install -y --no-install-recommends \
    gcc \
    libxml2 \
    libxslt1.1 \
    curl \
    libpango-1.0-0 \
    libpangocairo-1.0-0 \
    libcairo2 \
    libgdk-pixbuf-2.0-0 \
    libffi8 \
    shared-mime-info \
    fonts-dejavu-core \
 && rm -rf /var/lib/apt/lists/*

# Instalar deps Python explícitamente, sin depender del cache
RUN pip install --no-cache-dir --upgrade pip && \
    pip install --no-cache-dir \
      "fastapi>=0.110.0" \
      "uvicorn[standard]>=0.27.0" \
      "python-docx>=1.1.0" \
      "openpyxl>=3.1.0" \
      "weasyprint>=60.0"

# Verificación: el build falla aquí si algo no quedó instalado
RUN python -c "import fastapi, uvicorn, docx, openpyxl, weasyprint; \
    print('fastapi', fastapi.__version__); \
    print('openpyxl', openpyxl.__version__); \
    print('python-docx', docx.__version__); \
    print('weasyprint', weasyprint.__version__)"

# Copiar módulos Python, assets de imagen y archivos de ejemplo
COPY *.py *.png *.json ./

ENV ZEBRA_OUTPUT_DIR=/tmp/zebra
RUN mkdir -p /tmp/zebra

EXPOSE 8080

CMD ["uvicorn", "zebra_api:app", "--host", "0.0.0.0", "--port", "8080"]
