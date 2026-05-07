from typing import Optional

from fastapi import APIRouter, File, HTTPException, Query, UploadFile


from pdf_table_extractor.services.pdf_conversion_service import process_uploaded_pdf


router = APIRouter()


@router.post("/convert")
async def convert_pdf(
    customer_name: str,
    file: UploadFile = File(...),
):


    try:
        return await process_uploaded_pdf(file, customer_name)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc

