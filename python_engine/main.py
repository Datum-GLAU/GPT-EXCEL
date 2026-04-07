import os
import shutil
from pathlib import Path

import pandas as pd
from fastapi import FastAPI, UploadFile, File, HTTPException, Query
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse

from analysis import analyze_data
from excel_generator import read_excel, generate_advanced_excel, create_excel_template
from charts import create_chart
from data_cleaning import clean_data
from report_generator import generate_report
from word_generator import generate_word_report
from ppt_generator import generate_ppt
from prompt_engine import process_prompt
from file_segmentation import (
    segment_by_column, segment_by_row_count,
    segment_by_date_column, merge_excel_files, get_file_info
)
from SQLite_storage import (
    save_file_record, get_all_files, get_file_by_id,
    save_analysis, get_analysis_by_file,
    save_generated_file, get_generated_files,
    get_automation_logs, get_db_stats
)
from automation import run_in_background
from utils import ensure_directory, unique_output_path, is_within_directory

# ── App setup ─────────────────────────────────────────────────────────────────
app = FastAPI(
    title="GPT Excel — Python Engine",
    version="2.1.0",
    description="Offline-first data processing, analysis, chart, and document engine"
)

ELECTRON_ORIGIN = os.environ.get("ELECTRON_ORIGIN", "http://localhost:3000")

app.add_middleware(
    CORSMiddleware,
    allow_origins=[ELECTRON_ORIGIN, "http://localhost:3000", "http://localhost:5173"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

UPLOAD_DIR = ensure_directory(os.environ.get("UPLOAD_DIR", "uploads"))
OUTPUT_DIR = ensure_directory(os.environ.get("OUTPUT_DIR", "outputs"))
SEGMENT_DIR = ensure_directory(Path(OUTPUT_DIR) / "segments")
TEMPLATE_DIR = ensure_directory(Path(OUTPUT_DIR) / "templates")
ALLOWED_DOWNLOAD_ROOTS = [Path(UPLOAD_DIR).resolve(), Path(OUTPUT_DIR).resolve()]


# ── Startup ───────────────────────────────────────────────────────────────────
@app.on_event("startup")
def on_startup():
    run_in_background()


# ── Helpers ───────────────────────────────────────────────────────────────────
def save_upload(file: UploadFile) -> str:
    file_path = unique_output_path(UPLOAD_DIR, file.filename or "upload.xlsx")
    with open(file_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)
    return file_path


def out(filename: str) -> str:
    return unique_output_path(OUTPUT_DIR, filename)


def _record_file(file_path: str) -> int:
    try:
        df = pd.read_excel(file_path)
        return save_file_record(
            filename=os.path.basename(file_path),
            file_path=file_path,
            row_count=len(df),
            col_count=len(df.columns),
            columns=list(df.columns),
        )
    except Exception:
        return save_file_record(os.path.basename(file_path), file_path)


def _validate_excel_filename(filename: str | None) -> None:
    if not filename or not filename.lower().endswith((".xlsx", ".xls")):
        raise HTTPException(400, "Only Excel files (.xlsx / .xls) are allowed")


# ── Health ────────────────────────────────────────────────────────────────────
@app.get("/")
def root():
    return {
        "status": "running",
        "engine": "GPT Excel Python Engine",
        "version": "2.1.0",
        "mode": "offline-first",
    }


# ── File Upload ───────────────────────────────────────────────────────────────
@app.post("/upload")
async def upload_file(file: UploadFile = File(...)):
    _validate_excel_filename(file.filename)
    file_path = save_upload(file)
    file_id = _record_file(file_path)
    return {"message": "File uploaded successfully", "file_path": file_path, "file_id": file_id}


@app.post("/read")
async def read_file(file: UploadFile = File(...), limit: int = Query(25, ge=1, le=500)):
    _validate_excel_filename(file.filename)
    file_path = save_upload(file)
    data = read_excel(file_path, limit=limit)
    return {"data": data}


# ── Core Processing ───────────────────────────────────────────────────────────
@app.post("/analyze")
async def analyze_file(file: UploadFile = File(...)):
    _validate_excel_filename(file.filename)
    file_path = save_upload(file)
    result = analyze_data(file_path)
    file_id = _record_file(file_path)
    save_analysis(file_id, "analyze", result)
    return {"analysis": result}


@app.post("/chart")
async def generate_chart(
    file: UploadFile = File(...),
    chart_type: str = Query("auto", description="bar | line | pie | scatter | auto")
):
    _validate_excel_filename(file.filename)
    file_path = save_upload(file)
    output_path = out("chart.png")
    result = create_chart(file_path, chart_type, output_path=output_path)
    file_id = _record_file(file_path)
    save_generated_file(file_id, "chart", output_path)
    return {"message": result, "chart_path": output_path}


@app.post("/clean")
async def clean_file(file: UploadFile = File(...)):
    _validate_excel_filename(file.filename)
    file_path = save_upload(file)
    output_path = out("cleaned_data.xlsx")
    result = clean_data(file_path, output_path=output_path)
    file_id = _record_file(file_path)
    save_generated_file(file_id, "cleaned_excel", output_path)
    return {"message": result, "output_path": output_path}


@app.post("/report")
async def create_report(file: UploadFile = File(...)):
    _validate_excel_filename(file.filename)
    file_path = save_upload(file)
    output_path = out("report.txt")
    result = generate_report(file_path, output_path=output_path)
    file_id = _record_file(file_path)
    save_generated_file(file_id, "report_txt", output_path)
    return {"message": result, "output_path": output_path}


# ── Document Generation ───────────────────────────────────────────────────────
@app.post("/word")
async def create_word(file: UploadFile = File(...)):
    _validate_excel_filename(file.filename)
    file_path = save_upload(file)
    output_path = out("output_report.docx")
    result = generate_word_report(file_path, output_path=output_path)
    file_id = _record_file(file_path)
    save_generated_file(file_id, "word_doc", output_path)
    return {"message": result, "output_path": output_path}


@app.post("/ppt")
async def create_ppt(file: UploadFile = File(...)):
    _validate_excel_filename(file.filename)
    file_path = save_upload(file)
    output_path = out("output_report.pptx")
    result = generate_ppt(file_path, output_path=output_path)
    file_id = _record_file(file_path)
    save_generated_file(file_id, "pptx", output_path)
    return {"message": result, "output_path": output_path}


@app.post("/excel-advanced")
async def create_advanced_excel(file: UploadFile = File(...)):
    _validate_excel_filename(file.filename)
    file_path = save_upload(file)
    output_path = out("final_output.xlsx")
    result = generate_advanced_excel(file_path, output_path=output_path)
    file_id = _record_file(file_path)
    save_generated_file(file_id, "advanced_excel", output_path)
    return {"message": result, "output_path": output_path}


@app.post("/template")
async def make_template(
    rows: int = Query(10, ge=1, le=500),
    include_sample_data: bool = Query(True),
):
    output_path = unique_output_path(TEMPLATE_DIR, "excel_template.xlsx")
    result = create_excel_template(output_path=output_path, rows=rows, include_sample_data=include_sample_data)
    return {"message": result, "output_path": output_path}


# ── File Segmentation ─────────────────────────────────────────────────────────
@app.post("/segment/by-column")
async def segment_column(
    file: UploadFile = File(...),
    column: str = Query(..., description="Column name to segment by")
):
    _validate_excel_filename(file.filename)
    file_path = save_upload(file)
    result = segment_by_column(file_path, column, output_dir=str(SEGMENT_DIR))
    return result


@app.post("/segment/by-rows")
async def segment_rows(
    file: UploadFile = File(...),
    chunk_size: int = Query(1000, description="Rows per chunk")
):
    _validate_excel_filename(file.filename)
    file_path = save_upload(file)
    result = segment_by_row_count(file_path, chunk_size, output_dir=str(SEGMENT_DIR))
    return result


@app.post("/segment/by-date")
async def segment_date(
    file: UploadFile = File(...),
    date_column: str = Query(..., description="Date column name"),
    freq: str = Query("M", description="D=daily W=weekly M=monthly Q=quarterly Y=yearly")
):
    _validate_excel_filename(file.filename)
    file_path = save_upload(file)
    result = segment_by_date_column(file_path, date_column, freq, output_dir=str(SEGMENT_DIR))
    return result


@app.post("/file-info")
async def file_info(file: UploadFile = File(...)):
    _validate_excel_filename(file.filename)
    file_path = save_upload(file)
    return get_file_info(file_path)


@app.post("/merge")
async def merge_files(files: list[UploadFile] = File(...)):
    if len(files) < 2:
        raise HTTPException(400, "Upload at least two Excel files to merge")
    for file in files:
        _validate_excel_filename(file.filename)
    file_paths = [save_upload(file) for file in files]
    output_path = out("merged_output.xlsx")
    result = merge_excel_files(file_paths, output_path=output_path)
    if "error" in result:
        raise HTTPException(400, result["error"])
    return result


# ── Smart Prompt ──────────────────────────────────────────────────────────────
@app.post("/process")
async def process_with_prompt(
    file: UploadFile = File(...),
    prompt: str = Query(..., description="Natural language instruction")
):
    if not prompt.strip():
        raise HTTPException(400, "Prompt cannot be empty")
    _validate_excel_filename(file.filename)
    file_path = save_upload(file)
    result = process_prompt(file_path, prompt)
    return {"result": result}


# ── File Download ─────────────────────────────────────────────────────────────
@app.get("/download")
def download_file(path: str = Query(..., description="File path to download")):
    resolved = Path(path).resolve()
    if not resolved.exists():
        raise HTTPException(404, f"File not found: {path}")
    if not any(is_within_directory(resolved, root) for root in ALLOWED_DOWNLOAD_ROOTS):
        raise HTTPException(403, "Downloads are limited to uploads and generated outputs")
    return FileResponse(str(resolved), filename=resolved.name)


# ── SQLite / History ──────────────────────────────────────────────────────────
@app.get("/history/files")
def history_files():
    return {"files": get_all_files()}


@app.get("/history/files/{file_id}")
def history_file(file_id: int):
    record = get_file_by_id(file_id)
    if not record:
        raise HTTPException(404, "File record not found")
    return record


@app.get("/history/analysis/{file_id}")
def history_analysis(file_id: int):
    return {"results": get_analysis_by_file(file_id)}


@app.get("/history/outputs/{file_id}")
def history_outputs(file_id: int):
    return {"outputs": get_generated_files(file_id)}


@app.get("/history/automation")
def history_automation(limit: int = Query(50, ge=1, le=500)):
    return {"logs": get_automation_logs(limit)}


@app.get("/db/stats")
def db_stats():
    return get_db_stats()
