# GPT-EXCEL
 
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
pip install fastapi uvicorn pandas matplotlib openpyxl schedule xlsxwriter
uvicorn main:app --reload
```
 
Runs on: http://127.0.0.1:8000  
API Docs: http://127.0.0.1:8000/docs  
 
Handles advanced Excel generation, data processing, and chart creation.
 
---

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