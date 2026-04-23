FROM python:3.11-slim
 
WORKDIR /app
 
# Dependencias del sistema. libxml2 es requerido por lxml (dep de python-docx).
# curl lo usamos para el healthcheck del container.
RUN apt-get update && apt-get install -y --no-install-recommends \
    libxml2 \
    curl \
    && rm -rf /var/lib/apt/lists/*
 
# Dependencias de Python (capa separada para cache eficiente en rebuilds)
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
 
# Código del servicio
COPY zebra_proposal_builder.py .
COPY zebra_api.py .
 
# Directorio para archivos generados (efímero por default; montar volumen
# si se quiere persistencia entre reinicios)
RUN mkdir -p /tmp/zebra
ENV ZEBRA_OUTPUT_DIR=/tmp/zebra
 
EXPOSE 8080
 
# Healthcheck: EasyPanel lo usa para saber si el container está sano
HEALTHCHECK --interval=30s --timeout=5s --start-period=10s --retries=3 \
    CMD curl -fsS http://localhost:8080/health || exit 1
 
# --workers 2 es razonable para un VPS de 2GB RAM. Subir si hay más CPU disponible.
CMD ["uvicorn", "zebra_api:app", "--host", "0.0.0.0", "--port", "8080", "--workers", "2"]
