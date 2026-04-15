"""
web_app.py — Web interface using FastAPI (plug in when ready).

To activate:
  1. pip install fastapi uvicorn python-multipart
  2. Uncomment WEB_HOST / WEB_PORT in app/config.py and .env
  3. Run: uvicorn interfaces.web_app:app --host 0.0.0.0 --port 8000

This file contains ZERO business logic.
All processing is delegated to app.services.invoice_service.process_invoice().
"""

# ── Uncomment everything below when ready ─────────────────────────────────────

# import os
# import shutil
# from fastapi import FastAPI, UploadFile, File, HTTPException
# from fastapi.responses import FileResponse
#
# from app.config import ASSETS_DIR, OUTPUT_DIR, INPUT_DIR
# from app.services.invoice_service import process_invoice
#
# app = FastAPI(title="Commercial Invoice Generator")
#
#
# @app.get("/")
# def root():
#     return {"status": "ok", "message": "Invoice generator is running."}
#
#
# @app.post("/process")
# async def process(file: UploadFile = File(...)):
#     if not file.filename.lower().endswith(('.xls', '.xlsx')):
#         raise HTTPException(status_code=400, detail="Only .xls or .xlsx files accepted.")
#
#     input_path = os.path.join(INPUT_DIR, file.filename)
#     with open(input_path, "wb") as f:
#         shutil.copyfileobj(file.file, f)
#
#     try:
#         result = process_invoice(
#             excel_file_path=input_path,
#             assets_dir=ASSETS_DIR,
#             output_dir=OUTPUT_DIR,
#         )
#         return {
#             "status":       "success",
#             "invoice":      result["invoice_docx_path"],
#             "packing_list": result["packing_list_docx_path"],
#             "validation":   result["validation"],
#         }
#     except Exception as e:
#         raise HTTPException(status_code=500, detail=str(e))
#
#
# @app.get("/download/{filename}")
# def download(filename: str):
#     path = os.path.join(OUTPUT_DIR, filename)
#     if not os.path.exists(path):
#         raise HTTPException(status_code=404, detail="File not found.")
#     return FileResponse(path, filename=filename)
