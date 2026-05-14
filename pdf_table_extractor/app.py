from fastapi import FastAPI

from pdf_table_extractor.api.routes import router


app = FastAPI(title="PDF Table Extractor")
app.include_router(router)

