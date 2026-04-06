from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
import shutil
import os
 
from analysis import analyze_data
from excel_generator import read_excel, generate_advanced_excel
from charts import create_chart
from data_cleaning import clean_data
from report_generator import generate_report
from prompt_engine import process_prompt
 
app = FastAPI(title="GPT Excel - Python Engine", version="1.0.0")
 
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)
 
UPLOAD_DIR = "uploads"
os.makedirs(UPLOAD_DIR, exist_ok=True)
 
 
def save_upload(file: UploadFile) -> str:
    file_path = os.path.join(UPLOAD_DIR, file.filename)
    with open(file_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)
    return file_path
 
 
@app.get("/")
def root():
    return {"message": "GPT Excel Python Engine is running!"}
 
 
@app.post("/upload")
async def upload_file(file: UploadFile = File(...)):
    if not file.filename.endswith((".xlsx", ".xls")):
        raise HTTPException(status_code=400, detail="Only Excel files are allowed")
    file_path = save_upload(file)
    return {"message": "File uploaded successfully", "file_path": file_path}
 
 
@app.post("/read")
async def read_file(file: UploadFile = File(...)):
    file_path = save_upload(file)
    data = read_excel(file_path)
    return {"data": data}
 
 
@app.post("/analyze")
async def analyze_file(file: UploadFile = File(...)):
    file_path = save_upload(file)
    result = analyze_data(file_path)
    return {"analysis": result}
 
 
@app.post("/chart")
async def generate_chart(file: UploadFile = File(...), chart_type: str = "auto"):
    file_path = save_upload(file)
    result = create_chart(file_path, chart_type)
    return {"message": result, "chart_path": "chart.png"}
 
 
@app.post("/clean")
async def clean_file(file: UploadFile = File(...)):
    file_path = save_upload(file)
    result = clean_data(file_path)
    return {"message": result}
 
 
@app.post("/report")
async def create_report(file: UploadFile = File(...)):
    file_path = save_upload(file)
    result = generate_report(file_path)
    return {"message": result}
 
 
@app.post("/process")
async def process_with_prompt(file: UploadFile = File(...), prompt: str = ""):
    if not prompt:
        raise HTTPException(status_code=400, detail="Prompt cannot be empty")
    file_path = save_upload(file)
    result = process_prompt(file_path, prompt)
    return {"result": result}
