"""
zebra_api.py

Microservicio HTTP que expone el zebra_proposal_builder como API REST.
Pensado para ser consumido desde n8n (HTTP Request node) o desde un
dashboard custom.

Endpoints:
    POST /generate        → recibe proposal_data JSON, devuelve DOCX binario
    POST /generate/base64 → igual que /generate pero devuelve base64 (útil
                            cuando el cliente prefiere JSON-in-JSON-out)
    GET  /health          → healthcheck

Run (dev):
    uvicorn zebra_api:app --host 0.0.0.0 --port 8080 --reload

Run (prod):
    uvicorn zebra_api:app --host 0.0.0.0 --port 8080 --workers 2
"""

from __future__ import annotations

import base64
import os
import tempfile
import uuid
from typing import Any

from fastapi import FastAPI, HTTPException, Request, Response
from fastapi.responses import FileResponse, JSONResponse
from pydantic import BaseModel

from zebra_proposal_builder import generate_proposal


# ---------------------------------------------------------------------------
# Configuración
# ---------------------------------------------------------------------------

# Carpeta de archivos generados. Por defecto /tmp/zebra; se puede sobreescribir
# con la variable de entorno ZEBRA_OUTPUT_DIR.
OUTPUT_DIR = os.environ.get("ZEBRA_OUTPUT_DIR", "/tmp/zebra")
os.makedirs(OUTPUT_DIR, exist_ok=True)

# API key opcional. Si ZEBRA_API_KEY está seteada, todas las requests
# deben traer el header X-API-Key con ese valor.
API_KEY = os.environ.get("ZEBRA_API_KEY")

# Campos mínimos que deben estar en proposal_data
REQUIRED_FIELDS = [
    "cliente", "modelo", "objetivo",
    "contexto", "diagnostico", "cambio_paradigma",
    "sistema_propuesto", "arquitectura",
    "plan_implementacion", "metas", "inversion", "por_que_zebra",
]


app = FastAPI(
    title="Zebra Proposal API",
    description="Genera propuestas estratégicas en DOCX desde JSON estructurado.",
    version="1.0.0",
)


# ---------------------------------------------------------------------------
# Utilidades
# ---------------------------------------------------------------------------

def _check_auth(request: Request) -> None:
    """Valida API key si está configurada."""
    if not API_KEY:
        return
    key = request.headers.get("x-api-key")
    if key != API_KEY:
        raise HTTPException(status_code=401, detail="API key inválida o faltante.")


def _validate_schema(data: Any) -> list[str]:
    """Retorna una lista de errores de validación (vacía si es válido)."""
    errors: list[str] = []
    if not isinstance(data, dict):
        return ["proposal_data debe ser un objeto JSON (dict)."]

    for field in REQUIRED_FIELDS:
        if field not in data:
            errors.append(f"Campo obligatorio faltante: '{field}'.")

    diag = data.get("diagnostico")
    if isinstance(diag, list) and not (3 <= len(diag) <= 4):
        errors.append(
            f"'diagnostico' debe tener entre 3 y 4 items (tiene {len(diag)})."
        )

    modelos = data.get("sistema_propuesto", {}).get("modelos") if isinstance(data.get("sistema_propuesto"), dict) else None
    if isinstance(modelos, list) and not (2 <= len(modelos) <= 4):
        errors.append(
            f"'sistema_propuesto.modelos' debe tener entre 2 y 4 items "
            f"(tiene {len(modelos)})."
        )

    fases = data.get("plan_implementacion")
    if isinstance(fases, list) and len(fases) != 4:
        errors.append(
            f"'plan_implementacion' debe tener exactamente 4 fases "
            f"(tiene {len(fases)})."
        )

    kpis = data.get("metas", {}).get("kpis") if isinstance(data.get("metas"), dict) else None
    if isinstance(kpis, list) and len(kpis) != 3:
        errors.append(
            f"'metas.kpis' debe tener exactamente 3 items (tiene {len(kpis)})."
        )

    return errors


def _build_docx(proposal_data: dict, slug: str | None = None) -> str:
    """Genera el DOCX y retorna el path absoluto."""
    slug = slug or "propuesta"
    # Limpiar el slug para que sea seguro como nombre de archivo
    safe_slug = "".join(c if c.isalnum() or c in "-_" else "_" for c in slug).strip("_")
    if not safe_slug:
        safe_slug = "propuesta"

    filename = f"{safe_slug}_{uuid.uuid4().hex[:8]}.docx"
    out_path = os.path.join(OUTPUT_DIR, filename)
    generate_proposal(proposal_data, out_path)
    return out_path


# ---------------------------------------------------------------------------
# Modelos Pydantic (solo para documentación — el payload es pasthrough)
# ---------------------------------------------------------------------------

class GenerateRequest(BaseModel):
    """
    El body debe ser el objeto proposal_data completo.
    Opcionalmente puede incluir el campo especial __slug__ para controlar
    el nombre del archivo de salida.
    """
    model_config = {"extra": "allow"}


class GenerateResponse(BaseModel):
    status: str
    filename: str
    size_bytes: int
    docx_base64: str | None = None


class ErrorResponse(BaseModel):
    status: str = "error"
    error_type: str
    message: str
    errors: list[str] | None = None


# ---------------------------------------------------------------------------
# Endpoints
# ---------------------------------------------------------------------------

@app.get("/health")
def health():
    return {"status": "ok", "service": "zebra-proposal-api", "version": "1.0.0"}


@app.post(
    "/generate",
    responses={
        200: {
            "content": {"application/vnd.openxmlformats-officedocument.wordprocessingml.document": {}},
            "description": "DOCX binario con la propuesta generada.",
        },
        400: {"model": ErrorResponse},
        401: {"model": ErrorResponse},
        500: {"model": ErrorResponse},
    },
)
async def generate(request: Request):
    """
    Recibe proposal_data como JSON en el body y devuelve el DOCX binario.

    Headers:
        Content-Type: application/json
        X-API-Key: <key>  (si ZEBRA_API_KEY está seteada)

    Opcional en el body: campo __slug__ para nombrar el archivo (ej: "grupo_indes").
    """
    _check_auth(request)

    try:
        data = await request.json()
    except Exception as exc:
        raise HTTPException(
            status_code=400,
            detail={"error_type": "JSONDecodeError", "message": str(exc)},
        )

    # Soporte para diagnóstico preliminar (Claude puede devolverlo en casos
    # de información insuficiente). En ese caso no generamos DOCX; devolvemos
    # el diagnóstico tal cual con status 200 y content-type JSON.
    if isinstance(data, dict) and data.get("status") == "diagnostico_preliminar":
        return JSONResponse(content=data, status_code=200)

    errors = _validate_schema(data)
    if errors:
        raise HTTPException(
            status_code=400,
            detail={
                "error_type": "ValidationError",
                "message": "El JSON no cumple el esquema mínimo.",
                "errors": errors,
            },
        )

    slug = data.pop("__slug__", data.get("subtitulo_propuesta", "propuesta"))

    try:
        out_path = _build_docx(data, slug=slug)
    except Exception as exc:
        raise HTTPException(
            status_code=500,
            detail={"error_type": type(exc).__name__, "message": str(exc)},
        )

    filename = os.path.basename(out_path)
    return FileResponse(
        path=out_path,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=f"Propuesta_Zebra_{slug}.docx",
        headers={
            "X-Generated-Filename": filename,
            "X-Size-Bytes": str(os.path.getsize(out_path)),
        },
    )


@app.post(
    "/generate/base64",
    response_model=GenerateResponse,
    responses={
        400: {"model": ErrorResponse},
        401: {"model": ErrorResponse},
        500: {"model": ErrorResponse},
    },
)
async def generate_base64(request: Request):
    """
    Igual que /generate pero devuelve el DOCX como base64 dentro de un JSON.
    Útil cuando el cliente prefiere manejar todo como JSON (ej: n8n donde
    el post-procesamiento es más simple).
    """
    _check_auth(request)

    try:
        data = await request.json()
    except Exception as exc:
        raise HTTPException(
            status_code=400,
            detail={"error_type": "JSONDecodeError", "message": str(exc)},
        )

    if isinstance(data, dict) and data.get("status") == "diagnostico_preliminar":
        return JSONResponse(content=data, status_code=200)

    errors = _validate_schema(data)
    if errors:
        raise HTTPException(
            status_code=400,
            detail={
                "error_type": "ValidationError",
                "message": "El JSON no cumple el esquema mínimo.",
                "errors": errors,
            },
        )

    slug = data.pop("__slug__", data.get("subtitulo_propuesta", "propuesta"))

    try:
        out_path = _build_docx(data, slug=slug)
    except Exception as exc:
        raise HTTPException(
            status_code=500,
            detail={"error_type": type(exc).__name__, "message": str(exc)},
        )

    with open(out_path, "rb") as fh:
        docx_b64 = base64.b64encode(fh.read()).decode("ascii")

    return GenerateResponse(
        status="ok",
        filename=os.path.basename(out_path),
        size_bytes=os.path.getsize(out_path),
        docx_base64=docx_b64,
    )


# Manejador de HTTPException para dejar el formato de error consistente
@app.exception_handler(HTTPException)
async def http_exception_handler(request: Request, exc: HTTPException):
    detail = exc.detail if isinstance(exc.detail, dict) else {"message": str(exc.detail)}
    payload = {"status": "error", **detail}
    return JSONResponse(status_code=exc.status_code, content=payload)
