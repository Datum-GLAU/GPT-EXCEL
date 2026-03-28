import os
os.environ["OPENAI_API_KEY"] = "your_key_here"

from fastapi import FastAPI, UploadFile, File
import pandas as pd
import shutil
import threading

from automation import start_scheduler
from excel_generator import read_excel
from charts import create_chart
from analysis import analyze_data
from report_generator import generate_report
from data_cleaning import clean_data
from prompt_engine import process_prompt

app = FastAPI()

FILE_PATH = "uploaded_file.xlsx"

@app.get("/")
def home():
    return {"message": "Excel GPT Python Engine Running"}

# -------- Upload --------
@app.post("/upload")
async def upload_excel(file: UploadFile = File(...)):
    try:
        with open(FILE_PATH, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        df = pd.read_excel(FILE_PATH)

        return {
            "message": "File uploaded successfully",
            "rows": df.to_dict(orient="records")
        }

    except Exception as e:
        return {"error": str(e)}

# -------- Read --------
@app.get("/read")
def read():
    return {"data": read_excel(FILE_PATH)}

# -------- Analyze --------
@app.get("/analyze")
def analyze():
    return {"analysis": analyze_data(FILE_PATH)}

# -------- Chart --------
@app.get("/chart")
def chart():
    return {"result": create_chart(FILE_PATH)}

# -------- Clean --------
@app.get("/clean")
def clean():
    return {"result": clean_data(FILE_PATH)}

# -------- Report --------
@app.get("/report")
def report():
    return {"result": generate_report(FILE_PATH)}

# -------- Smart Prompt --------
@app.get("/process")
def process(prompt: str):
    return {"result": process_prompt(prompt, FILE_PATH)}

# -------- Automation (FINAL) --------
@app.on_event("startup")
def start_auto():
    thread = threading.Thread(target=start_scheduler)
    thread.daemon = True
    thread.start()
