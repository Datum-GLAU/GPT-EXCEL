# XtronExcel
 
Desktop application built with Electron, React, and Node.js (Express) to analyze Excel files, generate documents, charts, and automation workflows using natural language AI.
 
---
 
## 🚀 Quick Start
 
### 1. Frontend + Electron
 
```bash
npm install
npm run dev
```
 
The Electron app will launch automatically.
 
---
 
### 2. Node.js Server (API + AI Layer)
 
Open a **second terminal**:
 
```bash
cd server
npm install
node index.js
```
 
Runs on: http://localhost:3001  
Handles file uploads, Excel parsing, and all AI (Gemini / HuggingFace) calls.
 
---
### 3. Python Engine (Optional — Excel Generation)

Open a **third terminal**:

```bash
cd python_engine
pip install -r requirements.txt
uvicorn main:app --reload --port 8001
```

Runs on: http://127.0.0.1:8001   
API Docs: http://127.0.0.1:8001/docs

Handles Excel generation, data cleaning, chart creation, Word & PowerPoint
document generation, file segmentation, and offline automation.
---

---

### 4. Offline Python Commands

Run common Excel tasks locally without starting the Node server:

```bash
cd python_engine
python run_offline.py template
python run_offline.py info sample.xlsx
python run_offline.py analyze sample.xlsx
python run_offline.py clean sample.xlsx
python run_offline.py chart sample.xlsx
python run_offline.py pipeline sample.xlsx
python run_offline.py merge sample1.xlsx --files sample2.xlsx sample3.xlsx
```

Offline outputs are saved in `python_engine/offline_outputs` by default.
Background automation is disabled by default. Enable it only when needed:

```bash
set GPT_EXCEL_AUTOMATION=true
uvicorn main:app --reload --port 8001
```

---
 
## ⚙️ Tech Stack
 
| Layer | Technology |
|---|---|
| Frontend | React, TypeScript, Redux |
| Desktop | Electron |
| Backend | Node.js, Express |
| AI Providers | Google Gemini 2.0 Flash, HuggingFace Mistral-7B |
| Excel | SheetJS (xlsx), OpenPyXL |
| Python Engine | FastAPI, Pandas, Matplotlib |
 
---
