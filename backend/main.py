import os
import shutil
import tempfile
from typing import Optional
from fastapi import FastAPI, File, UploadFile, Form, BackgroundTasks
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel

from services.document_rebuilder import rebuild_pdf_to_docx, rebuild_docx, convert_docx_to_pdf_headless, clean_pdf_directly

app = FastAPI(title="Document Cleaner API")

# Configure CORS for frontend
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # Adjust in production
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

def remove_temp_file(path: str):
    try:
        if os.path.exists(path):
            os.remove(path)
    except Exception as e:
        print(f"Error removing temp file {path}: {e}")

@app.post("/api/process")
async def process_document(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...),
    output_format: str = Form("docx") # "docx" or "pdf"
):
    # Create temp directory for processing
    temp_dir = tempfile.mkdtemp()
    
    input_path = os.path.join(temp_dir, file.filename)
    
    # Save uploaded file
    with open(input_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)
        
    ext = os.path.splitext(file.filename)[1].lower()
    base_name, _ = os.path.splitext(file.filename)
    
    # Xây dựng lại file sạch từ đầu vào (DOCX hoặc PDF)
    if ext == ".pdf":
        output_path = os.path.join(temp_dir, f"rebuilt_{base_name}.docx")
        rebuild_pdf_to_docx(input_path, output_path)
    elif ext == ".docx":
        output_path = os.path.join(temp_dir, f"rebuilt_{base_name}.docx")
        rebuild_docx(input_path, output_path, temp_dir)
    else:
        return {"error": "Unsupported file format. Please upload .docx or .pdf"}

    # BƯỚC 3: Chuyển đổi định dạng nếu cần
    if output_format == "pdf":
        pdf_path = convert_docx_to_pdf_headless(output_path, temp_dir)
        if pdf_path and os.path.exists(pdf_path):
            output_path = pdf_path

    # Schedule cleanup after response
    # background_tasks.add_task(shutil.rmtree, temp_dir, ignore_errors=True)
    
    filename_to_return = os.path.basename(output_path)
    return FileResponse(
        output_path, 
        media_type='application/octet-stream', 
        filename=filename_to_return
    )

@app.get("/api/health")
def health_check():
    return {"status": "ok"}
