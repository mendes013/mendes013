"""
api_server.py
Servidor FastAPI que expõe a integração Claude-Excel como uma API REST.
Destinado a rodar na VPS Hostinger (Ubuntu 24.04, IP: 72.60.12.66).

Endpoints:
  POST /analyze          — analisa uma planilha enviada via upload
  POST /analyze-structured — análise com retorno JSON estruturado
  POST /generate-report  — gera um relatório Excel a partir de um prompt
  GET  /health           — verificação de saúde do serviço
"""

import os
import tempfile
from pathlib import Path
from typing import Any

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse
from pydantic import BaseModel

from claude_excel_integration import ClaudeExcelIntegration

app = FastAPI(
    title="Claude Excel API",
    description="API REST para integração entre Claude AI e planilhas Excel.",
    version="1.0.0",
)

# ---------------------------------------------------------------------------
# Modelos de resposta
# ---------------------------------------------------------------------------


class AnalyzeResponse(BaseModel):
    response: str


class ReportResponse(BaseModel):
    message: str
    filename: str


# ---------------------------------------------------------------------------
# Helper
# ---------------------------------------------------------------------------


def _get_integration() -> ClaudeExcelIntegration:
    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        raise HTTPException(status_code=500, detail="ANTHROPIC_API_KEY não configurada no servidor.")
    return ClaudeExcelIntegration(api_key=api_key)


# ---------------------------------------------------------------------------
# Endpoints
# ---------------------------------------------------------------------------


@app.get("/health")
def health() -> dict[str, str]:
    """Retorna o status de saúde do serviço."""
    return {"status": "ok", "service": "claude-excel-api"}


@app.post("/analyze", response_model=AnalyzeResponse)
async def analyze(
    file: UploadFile = File(..., description="Arquivo Excel (.xlsx)"),
    prompt: str = Form(..., description="Pergunta ou instrução para o Claude"),
    sheet_name: str = Form(None, description="Nome da aba (opcional, usa a ativa se omitido)"),
    system: str = Form(None, description="Mensagem de sistema para o Claude (opcional)"),
) -> AnalyzeResponse:
    """
    Recebe um arquivo Excel e um prompt, envia os dados ao Claude
    e retorna a resposta em texto.
    """
    if not file.filename or not file.filename.endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="Somente arquivos .xlsx são aceitos.")

    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        tmp.write(await file.read())
        tmp_path = tmp.name

    try:
        integration = _get_integration()
        response = integration.analyze_sheet(
            filepath=tmp_path,
            prompt=prompt,
            sheet_name=sheet_name or None,
            system=system or None,
        )
    finally:
        Path(tmp_path).unlink(missing_ok=True)

    return AnalyzeResponse(response=response)


@app.post("/analyze-structured")
async def analyze_structured(
    file: UploadFile = File(..., description="Arquivo Excel (.xlsx)"),
    prompt: str = Form(..., description="Pergunta ou instrução para o Claude"),
    json_schema: str = Form(..., description="Esquema JSON esperado na resposta (string JSON)"),
    sheet_name: str = Form(None, description="Nome da aba (opcional)"),
    system: str = Form(None, description="Mensagem de sistema (opcional)"),
) -> Any:
    """
    Analisa a planilha e retorna a resposta do Claude como JSON
    conforme o esquema fornecido.
    """
    import json

    if not file.filename or not file.filename.endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="Somente arquivos .xlsx são aceitos.")

    try:
        schema: dict[str, Any] = json.loads(json_schema)
    except json.JSONDecodeError as exc:
        raise HTTPException(status_code=400, detail=f"json_schema inválido: {exc}") from exc

    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        tmp.write(await file.read())
        tmp_path = tmp.name

    try:
        integration = _get_integration()
        result = integration.analyze_sheet_structured(
            filepath=tmp_path,
            prompt=prompt,
            json_schema=schema,
            sheet_name=sheet_name or None,
            system=system or None,
        )
    finally:
        Path(tmp_path).unlink(missing_ok=True)

    return result


@app.post("/generate-report")
async def generate_report(
    prompt: str = Form(..., description="Instrução para o Claude gerar o relatório"),
    system: str = Form(None, description="Mensagem de sistema (opcional)"),
    context_file: UploadFile = File(None, description="Planilha de contexto opcional (.xlsx)"),
) -> FileResponse:
    """
    Pede ao Claude para gerar um relatório tabular e retorna
    o arquivo Excel gerado para download.
    """
    context_path: str | None = None

    if context_file is not None:
        if not context_file.filename or not context_file.filename.endswith(".xlsx"):
            raise HTTPException(status_code=400, detail="Somente arquivos .xlsx são aceitos.")
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            tmp.write(await context_file.read())
            context_path = tmp.name

    output_path = tempfile.mktemp(suffix=".xlsx")

    try:
        integration = _get_integration()
        integration.generate_report(
            prompt=prompt,
            output_path=output_path,
            system=system or None,
            context_file=context_path,
        )
    finally:
        if context_path:
            Path(context_path).unlink(missing_ok=True)

    return FileResponse(
        path=output_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename="relatorio_claude.xlsx",
        background=None,
    )


# ---------------------------------------------------------------------------
# Ponto de entrada para execução direta
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    import uvicorn

    host = os.environ.get("API_HOST", "0.0.0.0")
    port = int(os.environ.get("API_PORT", "8000"))
    uvicorn.run("api_server:app", host=host, port=port, reload=False)
