from typing import Optional

from fastapi import APIRouter, File, HTTPException, Query, UploadFile

from pdf_table_extractor.core.config import DEFAULT_PATTERN
from pdf_table_extractor.services.pdf_conversion_service import process_uploaded_pdf


router = APIRouter()


@router.post("/convert")
async def convert_pdf(
    file: UploadFile = File(...),
    pattern_name: Optional[str] = None,
    customer_name: Optional[str] = Query(default=None, alias="customerName"),
):
    selected_pattern = pattern_name or customer_name or DEFAULT_PATTERN

    try:
        return await process_uploaded_pdf(file, selected_pattern)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc

