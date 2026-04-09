
# GPT-Excel Agent

A production-grade terminal-based AI Excel agent powered by Claude (Anthropic) or Gemini (Google) and ExcelJS.

Describe what you want in plain English — the agent reasons, plans, calls tools, and produces a professional `.xlsx` file.

---

## Setup

```bash
cd Ai_Agent
npm install
```

Set your API key for whichever model you prefer to use:

**For Gemini (Google):**
```bash
# Linux / macOS / Windows
export GEMINI_API_KEY=AIzaSy...
```

**For Claude (Anthropic):**
```bash
# Linux / macOS / Windows
export ANTHROPIC_API_KEY=sk-ant-...
```

---

## Usage

Run the agent script corresponding to the LLM you want to use:

```bash
# Gemini
node agent_gemini.mjs "<prompt>" <file.xlsx>

# Claude
node agent_claude.mjs "<prompt>" <file.xlsx>
```

- If the file **exists** → loads it and edits in-place.
- If the file **does not exist** → creates it from scratch.

---

## Example Commands

### 1. Create a budget tracker (new file)
```bash
node agent_gemini.mjs "create a monthly budget tracker with income, expenses, and savings for Jan–Dec" budget.xlsx
```

### 2. Edit an existing file
```bash
node agent_claude.mjs "add a Totals row at the bottom with SUM formulas for all numeric columns" sales.xlsx
```

### 3. Apply formulas
```bash
node agent_gemini.mjs "apply SUM formulas to column D and AVERAGE to column E, then bold the header row" data.xlsx
```

### 4. Add a chart
```bash
node agent_gemini.mjs "add a pie chart showing expenses by category" expenses.xlsx
```

### 5. Format and style
```bash
node agent_claude.mjs "make the header row blue with white text, format column C as currency, and auto-fit columns" report.xlsx
```

---

## How It Works

```
┌─────────────────────┐
│  CLI: prompt + file  │
└────────┬────────────┘
         │
         ▼
┌────────────────────┐     ┌──────────────────┐
│  Load / Create     │────▶│  Snapshot for     │
│  ExcelJS Workbook  │     │  rollback         │
└────────┬───────────┘     └──────────────────┘
         │
         ▼
┌─────────────────────────────────────────────┐
│              Agentic Loop                   │
│                                             │
│  1. Send messages + tools to LLM            │
│  2. If stop_reason = "end_turn" → break     │
│  3. If stop_reason = "tool_use":            │
│     • Execute each tool via ExcelJS         │
│     • Collect results                       │
│     • Push results back to messages         │
│     • Continue loop                         │
│                                             │
│  Max 20 iterations · 120s global timeout    │
└────────┬────────────────────────────────────┘
         │
         ▼
┌─────────────────────┐
│  Save .xlsx file    │
│  Print summary      │
└─────────────────────┘
```

### Available Tools

| Tool | Description |
|------|-------------|
| `create_sheet` | Creates a new worksheet with styled headers |
| `edit_cells` | Writes a 2-D data array into a cell range |
| `apply_formula` | Writes Excel formulas (SUM, AVERAGE, IF, etc.) |
| `format_cells` | Applies styling: bold, colours, number formats |
| `add_chart` | Adds a chart (best-effort; may write a placeholder) |
| `read_sheet` | Returns sheet data as JSON for verification |

---

## Troubleshooting

### API key not set
```
[error]  GEMINI_API_KEY environment variable is not set.
```
→ Make sure you've configured your `.env` file or exported the key in your current terminal session.

### File permission errors (EBUSY / EACCES)
```
[error]  Rollback also failed: EBUSY: resource busy or locked
```
→ **Close the Excel file** if it is open in Microsoft Excel or another program. Node.js is denied permission to overwrite a file while Excel has it actively locked.

### Chart limitations
ExcelJS does not have a stable chart API. When `add_chart` is called, the agent writes a descriptive placeholder cell instead. Open the resulting `.xlsx` in Excel and insert the chart manually using the data range noted in the placeholder.

### Max iterations reached
The agent loop is capped at 20 iterations to prevent runaway API calls. If the task is very complex, break it into smaller prompts.

### Timeout (120 seconds)
The entire agent run has a 2-minute global timeout. For very large files, consider operating on specific sheets rather than the entire workbook.

---

## File Structure

```
Ai_Agent/
├── agent_gemini.mjs   ← main entry point for Gemini 2.5 Flash
├── agent_claude.mjs   ← main entry point for Claude Haiku
├── tools.mjs          ← tool definitions (JSON schema)
├── executor.mjs       ← tool implementations using ExcelJS
├── utils.mjs          ← shared helpers (snapshot, logging)
├── package.json       ← dependencies + "type":"module"
└── README.md          ← this file
```

---

## License

MIT