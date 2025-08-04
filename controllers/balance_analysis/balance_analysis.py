"""Controller for Balance Analysis operations."""

import os
from fastapi import File, UploadFile, HTTPException
from fastapi.responses import FileResponse
from services.balance_analysis import balance_analysis as balance_analysis_service
from .config import BALANCE_ANALYSIS, router


@router.post(
    path='/balance-analysis',
    summary='Update Balance Analysis Automation',
    tags=[BALANCE_ANALYSIS['name']],
    response_class=FileResponse,
)
async def update_balance_analysis(
    files: list[UploadFile] = File(...),
) -> str:
    """Update the balance analysis automation."""
    excel_path = os.path.join("static", "files", "analise_balanco_modelo.xlsx")
    # Definições de colunas por ano
    col_map = {
        '2021': ('B', 'B'),
        '2022': ('C', 'D'),
        '2023': ('D', 'F'),
        '2024': ('E', 'H'),
    }
    for file in files:
        if file.content_type != "application/pdf":
            raise HTTPException(status_code=400, detail=f"{file.filename} não é um PDF.")
        name_no_ext, _ = os.path.splitext(file.filename)
        year = name_no_ext[-4:]
        if year not in col_map:
            raise HTTPException(
                status_code=400,
                detail=f"Ano '{year}' não suportado em {file.filename}."
            )
        bal_col, dre_col = col_map[year]
        pdf_bytes = await file.read()
        try:
            balance_analysis_service.process_balance_analysis_pdf(
                pdf_bytes=pdf_bytes,
                balanco_column_prefix=bal_col,
                dre_column_prefix=dre_col
            )
        except Exception as e:
            raise HTTPException(
                status_code=500,
                detail=f"Erro ao processar {file.filename}: {e}"
            ) from e

    return FileResponse(
        path=excel_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=os.path.basename(excel_path)
    )
