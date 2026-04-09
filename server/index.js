const express = require('express')
const cors = require('cors')
const multer = require('multer')
const path = require('path')
const fs = require('fs')
const XLSX = require('xlsx')
const { ensureDir, createWorkbookFromRows, readWorkbookPreview, runSpreadsheetAgent } = require('./ai-agent-runner')

const app = express()
const PORT = 3001

app.use(cors({
  origin: ['http://localhost:3000','http://localhost:5173','http://127.0.0.1:3000','http://127.0.0.1:5173'],
  credentials: true,
  methods: ['GET','POST','PUT','DELETE','PATCH','OPTIONS'],
  allowedHeaders: ['Content-Type','Authorization']
}))
app.use(express.json({ limit: '100mb' }))
app.use(express.urlencoded({ extended: true, limit: '100mb' }))

const UPLOAD_DIR = path.join(__dirname, 'uploads')
if (!fs.existsSync(UPLOAD_DIR)) fs.mkdirSync(UPLOAD_DIR, { recursive: true })
app.use('/uploads', express.static(UPLOAD_DIR))

app.use((req, res, next) => { console.log(`${req.method} ${req.url}`); next() })

const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    const dept = req.query.department || 'general'
    const dir = path.join(UPLOAD_DIR, dept)
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true })
    cb(null, dir)
  },
  filename: (req, file, cb) => cb(null, `${Date.now()}_${file.originalname}`)
})
const upload = multer({ storage, limits: { fileSize: 100 * 1024 * 1024 } })

// ── IN-MEMORY DB ─────────────────────────────────────────────────
let usersDB = [
  { id: 'u1', name: 'Admin User', email: 'admin@university.edu', role: 'admin', department: 'Administration', plan: 'pro' },
  { id: 'u2', name: 'Dr. Krishna Koushik', email: 'krishna@university.edu', role: 'faculty', department: 'CSE', plan: 'pro' },
  { id: 'u3', name: 'Prof. Demo', email: 'demo@university.edu', role: 'faculty', department: 'ECE', plan: 'free' },
]
let filesDB = []
let workflowsDB = [
  { id: 'w1', name: 'Monthly Report Generator', status: 'active', schedule: 'Monthly', lastRun: 'Mar 1, 2026', runs: 12 },
  { id: 'w2', name: 'Weekly Attendance Summary', status: 'active', schedule: 'Weekly', lastRun: 'Mar 24, 2026', runs: 8 },
]

function slugifyName(value = '', fallback = 'ai_result') {
  const cleaned = String(value)
    .replace(/\.[^.]+$/, '')
    .trim()
    .replace(/[^a-zA-Z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '')
  return cleaned || fallback
}

function buildAiWorkbookName(preferredName, fallback = 'ai_result') {
  const base = slugifyName(preferredName, fallback).slice(0, 60)
  return `${base}.xlsx`
}

function buildAiTextName(preferredName, fallback = 'ai_document', ext = 'txt') {
  const base = slugifyName(preferredName, fallback).slice(0, 60)
  return `${base}.${ext}`
}

function saveGeneratedAsset({ department = 'general', name, content, uploadedBy = 'AI' }) {
  const dir = path.join(UPLOAD_DIR, department)
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true })
  const filename = `${Date.now()}_${name}`
  const fullPath = path.join(dir, filename)
  fs.writeFileSync(fullPath, content, 'utf8')

  const record = {
    id: `f${Date.now()}_${Math.random().toString(36).slice(2, 6)}`,
    name,
    path: fullPath,
    size: fs.statSync(fullPath).size,
    type: path.extname(name).toLowerCase().slice(1),
    department,
    uploadedBy,
    uploadedAt: new Date().toISOString(),
    url: `/uploads/${department}/${filename}`
  }
  filesDB.push(record)
  syncFilesDBFromDisk()
  return record
}

function saveGeneratedWorkbook({ department = 'general', name, rows, uploadedBy = 'AI', sheetName = 'Result' }) {
  const dir = path.join(UPLOAD_DIR, department)
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true })
  const filename = `${Date.now()}_${name}`
  const fullPath = path.join(dir, filename)
  createWorkbookFromRows(fullPath, rows || [], sheetName)

  const record = {
    id: `f${Date.now()}_${Math.random().toString(36).slice(2, 6)}`,
    name,
    path: fullPath,
    size: fs.statSync(fullPath).size,
    type: 'xlsx',
    department,
    uploadedBy,
    uploadedAt: new Date().toISOString(),
    url: `/uploads/${department}/${filename}`
  }
  filesDB.push(record)
  syncFilesDBFromDisk()
  return record
}

function isQuotaExceededError(error) {
  const message = String(error?.message || error || '')
  return /quota|resource_exhausted|429|rate.?limit/i.test(message)
}

function recordFromPath(fullPath) {
  const rel = path.relative(UPLOAD_DIR, fullPath)
  const parts = rel.split(path.sep)
  const department = parts.length > 1 ? parts[0] : 'general'
  const filename = path.basename(fullPath)
  const stat = fs.statSync(fullPath)
  return {
    id: `disk_${rel.replace(/[\\/. ]+/g, '_')}`,
    name: filename.replace(/^\d+_/, ''),
    path: fullPath,
    size: stat.size,
    type: path.extname(filename).toLowerCase().slice(1),
    department,
    uploadedBy: 'local',
    uploadedAt: stat.mtime.toISOString(),
    url: `/uploads/${rel.split(path.sep).join('/')}`
  }
}

function syncFilesDBFromDisk() {
  const seen = new Map()
  const allowedExts = new Set([
    '.xlsx', '.xls', '.csv',
    '.txt', '.md', '.json',
    '.png', '.jpg', '.jpeg', '.gif', '.webp', '.svg',
    '.pdf', '.doc', '.docx', '.ppt', '.pptx'
  ])

  function walk(dir) {
    for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
      const full = path.join(dir, entry.name)
      if (entry.isDirectory()) {
        walk(full)
        continue
      }
      const ext = path.extname(entry.name).toLowerCase()
      if (!allowedExts.has(ext)) continue
      try {
        const record = recordFromPath(full)
        seen.set(record.path, record)
      } catch (e) {}
    }
  }

  walk(UPLOAD_DIR)

  const memoryRecords = filesDB.filter(file => fs.existsSync(file.path))
  memoryRecords.forEach(file => {
    seen.set(file.path, { ...seen.get(file.path), ...file })
  })

  filesDB = Array.from(seen.values()).sort((a, b) => new Date(b.uploadedAt) - new Date(a.uploadedAt))
}

syncFilesDBFromDisk()

function resolveUploadPath(target = '') {
  const resolved = path.resolve(UPLOAD_DIR, String(target || ''))
  const relative = path.relative(UPLOAD_DIR, resolved)
  if (relative.startsWith('..') || path.isAbsolute(relative)) {
    throw new Error('Path must stay within uploads directory')
  }
  return resolved
}

// ── KEY STORE ─────────────────────────────────────────────────────
const ROOT_DIR = path.resolve(__dirname, '..')
const AGENT_ENV_FILE = path.join(ROOT_DIR, 'Ai_Agent', '.env')
const ROOT_ENV_FILE = path.join(ROOT_DIR, '.env')

function readEnvFile(filePath) {
  if (!fs.existsSync(filePath)) return {}
  const values = {}
  const lines = fs.readFileSync(filePath, 'utf8').split(/\r?\n/)
  for (const line of lines) {
    const trimmed = line.trim()
    if (!trimmed || trimmed.startsWith('#')) continue
    const match = trimmed.match(/^([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.*)$/)
    if (!match) continue
    const [, key, rawValue] = match
    values[key] = rawValue.replace(/^['"]|['"]$/g, '').trim()
  }
  return values
}

function writeEnvFile(filePath, nextValues = {}) {
  ensureDir(path.dirname(filePath))
  const existing = readEnvFile(filePath)
  const merged = { ...existing, ...nextValues }
  const lines = Object.entries(merged)
    .filter(([, value]) => value !== undefined && value !== null && String(value).trim() !== '')
    .map(([key, value]) => `${key}=${String(value).trim()}`)
  fs.writeFileSync(filePath, `${lines.join('\n')}${lines.length ? '\n' : ''}`, 'utf8')
}

function getAgentConfig() {
  const agentEnv = readEnvFile(AGENT_ENV_FILE)
  const rootEnv = readEnvFile(ROOT_ENV_FILE)
  const geminiKey = process.env.GEMINI_API_KEY || agentEnv.GEMINI_API_KEY || rootEnv.GEMINI_API_KEY || ''
  const anthropicKey = process.env.ANTHROPIC_API_KEY || agentEnv.ANTHROPIC_API_KEY || rootEnv.ANTHROPIC_API_KEY || ''
  const requestedProvider = String(process.env.AI_PROVIDER || agentEnv.AI_PROVIDER || rootEnv.AI_PROVIDER || '').toLowerCase()

  let provider = 'gemini'
  if (requestedProvider === 'claude' && anthropicKey) provider = 'claude'
  else if (requestedProvider === 'gemini' && geminiKey) provider = 'gemini'
  else if (geminiKey) provider = 'gemini'
  else if (anthropicKey) provider = 'claude'
  else if (requestedProvider === 'claude') provider = 'claude'

  return {
    mode: 'ai-agent',
    provider,
    gemini: !!geminiKey,
    anthropic: !!anthropicKey,
    ready: !!(geminiKey || anthropicKey),
    envFilePresent: fs.existsSync(AGENT_ENV_FILE),
    rootEnvPresent: fs.existsSync(ROOT_ENV_FILE)
  }
}

function pickColumnIndex(headers = [], patterns = []) {
  return headers.findIndex(header => patterns.some(pattern => pattern.test(String(header || ''))))
}

function getDataRows(rawRows = []) {
  return Array.isArray(rawRows) && rawRows.length > 1 ? rawRows.slice(1).filter(row => Array.isArray(row) && row.some(cell => String(cell ?? '').trim() !== '')) : []
}

function computeAverageForRow(row = [], numericIndexes = []) {
  const values = numericIndexes.map(index => parseFloat(row[index])).filter(value => !Number.isNaN(value))
  if (!values.length) return 0
  return +(values.reduce((sum, value) => sum + value, 0) / values.length).toFixed(2)
}

function buildRowsFromObjects(headers, objects) {
  return [headers, ...objects.map(object => headers.map(header => object[header] ?? ''))]
}

function generateFormulaSuggestion(description = '', headers = []) {
  const lower = String(description).toLowerCase()
  const cleanedHeaders = headers.map(header => String(header || ''))
  const amountIdx = pickColumnIndex(cleanedHeaders, [/total/i, /amount/i, /marks?/i, /score/i, /attendance/i])
  const amountCol = amountIdx >= 0 ? XLSX.utils.encode_col(amountIdx) : 'B'

  if (/average|mean/.test(lower)) {
    return {
      formula: `=AVERAGE(${amountCol}2:${amountCol}100)`,
      explanation: `Calculates the average of the values in column ${amountCol}.`,
      alternatives: [`=ROUND(AVERAGE(${amountCol}2:${amountCol}100),2)`]
    }
  }

  if (/percentage|pass rate|rate/.test(lower)) {
    return {
      formula: '=COUNTIF(B2:B100,">=40")/COUNT(B2:B100)',
      explanation: 'Counts values meeting the condition and divides by the total numeric entries.',
      alternatives: ['=ROUND(COUNTIF(B2:B100,">=40")/COUNT(B2:B100)*100,2)']
    }
  }

  if (/lookup|xlookup|vlookup|find/.test(lower)) {
    return {
      formula: '=XLOOKUP(E2,A:A,B:B,"Not found")',
      explanation: 'Looks up the value in E2 within column A and returns the matching value from column B.',
      alternatives: ['=VLOOKUP(E2,A:B,2,FALSE)']
    }
  }

  if (/grade|pass|fail/.test(lower)) {
    return {
      formula: '=IF(B2>=40,"Pass","Fail")',
      explanation: 'Checks whether the score is at least 40 and labels the result.',
      alternatives: ['=IFS(B2>=90,"O",B2>=80,"A+",B2>=70,"A",B2>=60,"B+",B2>=50,"B",B2>=40,"C",TRUE,"F")']
    }
  }

  return {
    formula: `=SUM(${amountCol}2:${amountCol}100)`,
    explanation: `Adds all numeric values in column ${amountCol}.`,
    alternatives: [`=SUBTOTAL(9,${amountCol}2:${amountCol}100)`]
  }
}

function buildCommandResult(command = '', rawRows = [], stats = null) {
  const headers = (rawRows[0] || []).map(cell => String(cell ?? ''))
  const rows = getDataRows(rawRows)
  if (!headers.length || !rows.length) {
    return { action: 'summary', description: 'No spreadsheet data is loaded right now.', answer: 'No spreadsheet data is loaded right now.' }
  }

  const lower = String(command).toLowerCase()
  const attendanceIdx = pickColumnIndex(headers, [/attend/i])
  const nameIdx = pickColumnIndex(headers, [/^name$/i, /student/i])
  const sectionIdx = pickColumnIndex(headers, [/section/i, /\bsec\b/i])
  const requestedMetricIdx = headers.findIndex(header => {
    const normalizedHeader = String(header || '').toLowerCase()
    return normalizedHeader && lower.includes(normalizedHeader)
  })
  const numericIndexes = headers
    .map((header, index) => ({ header, index }))
    .filter(({ header }) => !/name|roll|id|section|department|email|phone/i.test(header))
    .filter(({ index }) => rows.filter(row => !Number.isNaN(parseFloat(row[index]))).length >= Math.max(1, Math.floor(rows.length * 0.4)))
    .map(({ index }) => index)

  const averageIndex = pickColumnIndex(headers, [/average/i, /\bavg\b/i])
  const rankingIndex = requestedMetricIdx >= 0 && numericIndexes.includes(requestedMetricIdx)
    ? requestedMetricIdx
    : averageIndex >= 0
      ? averageIndex
      : -1

  const enriched = rows.map(row => {
    const avg = rankingIndex >= 0 ? parseFloat(row[rankingIndex]) || 0 : computeAverageForRow(row, numericIndexes)
    return { row, avg }
  })

  if (/top\s*\d+|highest|rank/.test(lower)) {
    const countMatch = lower.match(/top\s*(\d+)/)
    const limit = countMatch ? Math.max(parseInt(countMatch[1], 10), 1) : 10
    const topRows = enriched.sort((a, b) => b.avg - a.avg).slice(0, limit).map(item => item.row)
    const metricName = rankingIndex >= 0 ? headers[rankingIndex] : 'average score'
    return {
      action: 'sort',
      description: `Top ${topRows.length} rows ranked by ${metricName}.`,
      rows: [headers, ...topRows]
    }
  }

  if (/failed|fail\b/.test(lower)) {
    const failedRows = enriched.filter(item => item.avg < 40).map(item => item.row)
    return {
      action: 'filter',
      description: failedRows.length ? `Found ${failedRows.length} failed students.` : 'No failed students were found.',
      rows: failedRows.length ? [headers, ...failedRows] : [headers]
    }
  }

  if (/\bpassed|pass\b/.test(lower)) {
    const passedRows = enriched.filter(item => item.avg >= 40).map(item => item.row)
    return {
      action: 'filter',
      description: passedRows.length ? `Found ${passedRows.length} passed students.` : 'No passed students were found.',
      rows: passedRows.length ? [headers, ...passedRows] : [headers]
    }
  }

  if (/below\s*75|attendance/.test(lower) && attendanceIdx >= 0) {
    const filtered = rows.filter(row => (parseFloat(row[attendanceIdx]) || 0) < 75)
    return {
      action: 'filter',
      description: filtered.length ? `Found ${filtered.length} rows with attendance below 75%.` : 'No rows are below 75% attendance.',
      rows: filtered.length ? [headers, ...filtered] : [headers]
    }
  }

  if (/section/.test(lower) && sectionIdx >= 0) {
    const sectionMap = new Map()
    for (const item of enriched) {
      const key = String(item.row[sectionIdx] || 'Unknown')
      if (!sectionMap.has(key)) sectionMap.set(key, [])
      sectionMap.get(key).push(item.avg)
    }
    const summaryHeaders = ['Section', 'Students', 'Average Score']
    const summaryRows = Array.from(sectionMap.entries()).map(([section, avgs]) => [
      section,
      avgs.length,
      +(avgs.reduce((sum, value) => sum + value, 0) / avgs.length).toFixed(2)
    ])
    return {
      action: 'summary_table',
      description: 'Built a section-wise summary.',
      rows: [summaryHeaders, ...summaryRows]
    }
  }

  if (/subject average|subject stats|subject/.test(lower) && stats?.subjectStats?.length) {
    const summaryHeaders = ['Subject', 'Average', 'Max', 'Min', 'Pass Rate']
    const summaryRows = stats.subjectStats.map(subject => [
      subject.subject,
      subject.avg,
      subject.max,
      subject.min,
      `${subject.passRate}%`
    ])
    return {
      action: 'summary_table',
      description: 'Built a subject-wise summary.',
      rows: [summaryHeaders, ...summaryRows]
    }
  }

  if (/grade/.test(lower) && !headers.some(header => /^grade$/i.test(header))) {
    const nextHeaders = [...headers, 'Grade']
    const nextRows = rows.map(row => {
      const avg = computeAverageForRow(row, numericIndexes)
      const grade = avg >= 90 ? 'O' : avg >= 80 ? 'A+' : avg >= 70 ? 'A' : avg >= 60 ? 'B+' : avg >= 50 ? 'B' : avg >= 40 ? 'C' : 'F'
      return [...row, grade]
    })
    return {
      action: 'add_column',
      column_name: 'Grade',
      description: 'Added a Grade column based on the average score.',
      rows: [nextHeaders, ...nextRows]
    }
  }

  if (/who|find|search|show|details/.test(lower) && nameIdx >= 0) {
    const query = lower
      .replace(/\b(who|find|search|show|details|for|student|about)\b/g, ' ')
      .trim()
    if (query) {
      const matches = rows.filter(row => row.some(cell => String(cell || '').toLowerCase().includes(query)))
      if (matches.length) {
        const answer = matches.slice(0, 3).map(row => headers.map((header, index) => `${header}: ${row[index] ?? ''}`).join(', ')).join('\n')
        return {
          action: 'summary',
          description: `Found ${matches.length} matching row(s).`,
          answer
        }
      }
    }
  }

  if (stats) {
    const answer = [
      `Total students: ${stats.total}`,
      `Passed: ${stats.passed} (${stats.passRate}%)`,
      `Failed: ${stats.failed}`,
      `Average score: ${stats.avgScore}`,
      stats.below75Count != null ? `Below 75% attendance: ${stats.below75Count}` : ''
    ].filter(Boolean).join('\n')
    return {
      action: 'summary',
      description: 'Summarized the current sheet.',
      answer
    }
  }

  return {
    action: 'summary',
    description: 'The sheet is loaded and ready.',
    answer: `Loaded ${rows.length} data rows across ${headers.length} columns. Ask for a filter, summary, grade column, or a new generated sheet.`
  }
}

function buildChatResponse(message = '', rawRows = [], stats = null) {
  const lower = String(message).toLowerCase().trim()
  if (!rawRows?.length && !stats) {
    return 'I can help create a workbook, analyze a loaded sheet, suggest formulas, or prepare a filtered result once data is available.'
  }

  if (/^(hi|hello|hey|hii|yo)\b/.test(lower)) {
    return 'Hi. Tell me what you want to create or change.'
  }

  if (/^(ok|okay|kk|cool|fine|alright|thanks|thank you)\b/.test(lower)) {
    return 'Tell me the next thing you want me to do.'
  }

  if (/formula|sumif|xlookup|vlookup|countif|averageif|iferror/.test(lower)) {
    const suggestion = generateFormulaSuggestion(message, rawRows[0] || [])
    return `${suggestion.explanation}\n\nFormula: ${suggestion.formula}`
  }

  const looksLikeDataIntent = /top|fail|pass|section|attend|below|above|sort|filter|search|add.*col|grade|subject|rank|list|show|find|who|which|compare|avg|average|highest|lowest|summary|analy[sz]e/.test(lower)
  if (!looksLikeDataIntent) {
    return 'Describe the exact result you want.'
  }

  const commandResult = buildCommandResult(message, rawRows, stats)
  if (commandResult.action === 'summary') {
    return commandResult.answer || commandResult.description || 'Done.'
  }
  if (commandResult.rows?.length > 1) {
    return `${commandResult.description}\n\nI prepared a tabular result with ${commandResult.rows.length - 1} row(s).`
  }
  return commandResult.description || 'Done.'
}

// ── EXCEL PARSER ──────────────────────────────────────────────────
function parseExcelFull(filePath) {
  try {
    const wb = XLSX.readFile(filePath, { cellStyles: true, cellDates: true })
    const result = {}
    wb.SheetNames.forEach(name => {
      const ws = wb.Sheets[name]
      const range = XLSX.utils.decode_range(ws['!ref'] || 'A1')
      const rows = []
      for (let R = range.s.r; R <= range.e.r; R++) {
        const row = []
        for (let C = range.s.c; C <= range.e.c; C++) {
          const cell = ws[XLSX.utils.encode_cell({ r: R, c: C })]
          row.push({ v: cell ? (cell.v !== undefined ? cell.v : '') : '', addr: XLSX.utils.encode_cell({ r: R, c: C }) })
        }
        rows.push(row)
      }
      result[name] = {
        rows,
        rowCount: range.e.r - range.s.r + 1,
        colCount: range.e.c - range.s.c + 1,
        rawRows: XLSX.utils.sheet_to_json(ws, { defval: '', header: 1 })
      }
    })
    return result
  } catch (e) { console.error('Parse error:', e); return null }
}

function computeStats(rawRows) {
  if (!rawRows || rawRows.length < 2) return null
  const headers = rawRows[0].map(h => String(h || ''))
  const dataRows = rawRows.slice(1).filter(r => r.some(c => c !== ''))
  if (!dataRows.length) return null

  const skip = new Set(['name','roll','rollno','roll_no','id','student','section','department','dept','branch','email','phone','sno','s.no','sl.no'])
  const markColIdxs = headers.map((h, i) => {
    const key = h.toLowerCase().replace(/[\s_.\-]/g, '')
    if (skip.has(key)) return -1
    const vals = dataRows.map(r => parseFloat(r[i])).filter(v => !isNaN(v))
    return vals.length > dataRows.length * 0.4 ? i : -1
  }).filter(i => i >= 0)

  const attendIdx = headers.findIndex(h => /attend/i.test(h))
  const nameIdx = headers.findIndex(h => /^name$/i.test(h.trim()))
  const rollIdx = headers.findIndex(h => /roll|sno|s\.no/i.test(h))
  const secIdx = headers.findIndex(h => /section|sec/i.test(h))
  const deptIdx = headers.findIndex(h => /dept|department/i.test(h))
  const passThreshold = 40

  const students = dataRows.map(row => {
    const marks = markColIdxs.map(i => parseFloat(row[i]) || 0)
    const total = marks.reduce((a, b) => a + b, 0)
    const avg = marks.length ? total / marks.length : 0
    const passed = marks.length ? marks.every(m => m >= passThreshold) : false
    const attendance = attendIdx >= 0 ? parseFloat(row[attendIdx]) || 0 : 0
    const avg2dp = +avg.toFixed(2)
    return {
      name: nameIdx >= 0 ? String(row[nameIdx] || 'Unknown') : 'Unknown',
      roll: rollIdx >= 0 ? String(row[rollIdx] || '') : '',
      section: secIdx >= 0 ? String(row[secIdx] || 'N/A') : 'N/A',
      department: deptIdx >= 0 ? String(row[deptIdx] || 'N/A') : 'N/A',
      marks, markCols: markColIdxs.map(i => headers[i]),
      total, avg: avg2dp,
      passed, attendance,
      below75: attendance > 0 && attendance < 75,
      grade: avg2dp >= 90 ? 'O' : avg2dp >= 80 ? 'A+' : avg2dp >= 70 ? 'A' : avg2dp >= 60 ? 'B+' : avg2dp >= 50 ? 'B' : avg2dp >= 40 ? 'C' : 'F'
    }
  }).filter(s => s.name !== 'Unknown' || s.marks.some(m => m > 0))

  if (!students.length) return null

  const passed = students.filter(s => s.passed)
  const failed = students.filter(s => !s.passed)
  const avgScore = students.reduce((a, b) => a + b.avg, 0) / students.length
  const below75 = students.filter(s => s.below75)

  const sections = {}
  students.forEach(s => { if (!sections[s.section]) sections[s.section] = []; sections[s.section].push(s) })
  const sectionStats = Object.entries(sections).map(([sec, stds]) => ({
    section: sec, count: stds.length,
    avg: +(stds.reduce((a, b) => a + b.avg, 0) / stds.length).toFixed(2),
    passed: stds.filter(s => s.passed).length,
    failed: stds.filter(s => !s.passed).length,
    passRate: +((stds.filter(s => s.passed).length / stds.length) * 100).toFixed(1)
  }))

  const subjectStats = markColIdxs.map((colIdx, mi) => {
    const vals = students.map(s => s.marks[mi])
    const avg = vals.reduce((a, b) => a + b, 0) / vals.length
    const p = vals.filter(v => v >= passThreshold).length
    return {
      subject: headers[colIdx], avg: +avg.toFixed(2),
      max: Math.max(...vals), min: Math.min(...vals),
      passed: p, failed: vals.length - p,
      passRate: +((p / vals.length) * 100).toFixed(1)
    }
  })

  const gradeMap = {}
  students.forEach(s => { gradeMap[s.grade] = (gradeMap[s.grade] || 0) + 1 })
  const gradeDistribution = ['O','A+','A','B+','B','C','F'].filter(g => gradeMap[g]).map(grade => ({ grade, count: gradeMap[grade] }))

  const scoreDistribution = []
  for (let i = 0; i <= 90; i += 10) {
    scoreDistribution.push({ range: `${i}-${i+9}`, count: students.filter(s => s.avg >= i && s.avg < i + 10).length })
  }

  return {
    total: students.length, passed: passed.length, failed: failed.length,
    passRate: +((passed.length / students.length) * 100).toFixed(1),
    failRate: +((failed.length / students.length) * 100).toFixed(1),
    avgScore: +avgScore.toFixed(2), topStudents: [...students].sort((a, b) => b.avg - a.avg).slice(0, 10),
    below75, below75Count: below75.length, sectionStats, subjectStats,
    gradeDistribution, scoreDistribution, students,
    markCols: markColIdxs.map(i => headers[i]), headers
  }
}

function buildDataContext(rawRows, stats) {
  let ctx = ''
  if (rawRows && rawRows.length > 0) {
    const headers = rawRows[0].map(h => String(h || ''))
    ctx += `SPREADSHEET HEADERS: ${headers.join(', ')}\n`
    ctx += `TOTAL ROWS: ${rawRows.length - 1}\n`
    ctx += `SAMPLE (first 8 rows):\n`
    rawRows.slice(0, 9).forEach(row => { ctx += row.map(c => String(c || '')).join('\t') + '\n' })
  }
  if (stats) {
    ctx += `\nSTATISTICS:\n`
    ctx += `Total: ${stats.total}, Passed: ${stats.passed} (${stats.passRate}%), Failed: ${stats.failed}\n`
    ctx += `Average: ${stats.avgScore}, Below 75% attendance: ${stats.below75Count}\n`
    if (stats.sectionStats && stats.sectionStats.length) {
      ctx += `Sections: ${stats.sectionStats.map(s => `${s.section}(avg=${s.avg},pass=${s.passRate}%)`).join(', ')}\n`
    }
    if (stats.subjectStats && stats.subjectStats.length) {
      ctx += `Subjects: ${stats.subjectStats.map(s => `${s.subject}(avg=${s.avg})`).join(', ')}\n`
    }
    if (stats.topStudents && stats.topStudents.length) {
      ctx += `Top 3: ${stats.topStudents.slice(0, 3).map(s => `${s.name}(${s.avg})`).join(', ')}\n`
    }
    ctx += `Mark columns: ${stats.markCols ? stats.markCols.join(', ') : 'N/A'}\n`
  }
  return ctx
}

function buildAgentRules({ preserveSourceSheet = false, allowSyntheticData = false } = {}) {
  return [
    'Interpret the user request directly and complete it inside the workbook.',
    'Operate on the real workbook data and sheet sizes, not only on the preview context.',
    'Handle long worksheets completely. Do not truncate filtered, sorted, grouped, or ranked outputs.',
    preserveSourceSheet
      ? 'Preserve existing sheets unless the user explicitly asks to replace them. Put derived output into a separate result worksheet.'
      : 'Create a polished workbook that satisfies the request end to end.',
    'Infer relevant columns dynamically from the workbook schema.',
    'When the request involves filtering, sorting, ranking, grouping, thresholds, sections, percentages, pass or fail logic, or attendance logic, produce the full matching result sheet.',
    'When formulas improve correctness or maintainability, use formulas that scale with the true row count.',
    'Use clear worksheet names, clean headers, readable formatting, and consistent structure.',
    allowSyntheticData
      ? 'Only generate new rows when the user is asking for a brand-new workbook or synthetic dataset.'
      : 'Do not invent sample or demo rows. Work only from the workbook data.',
    'Avoid canned examples and do not narrow the task to a few special cases.'
  ].join('\n- ')
}

function buildWorkbookAgentPrompt({ request, rawRows, stats, preserveSourceSheet = false, allowSyntheticData = false }) {
  return [
    `User request: "${request}"`,
    'Workbook operating rules:',
    `- ${buildAgentRules({ preserveSourceSheet, allowSyntheticData })}`,
    rawRows?.length ? `Workbook preview:\n${buildDataContext(rawRows, stats)}` : '',
    preserveSourceSheet
      ? 'Create or update a dedicated result worksheet with the complete answer to the request.'
      : 'Create the workbook from scratch with complete usable output.'
  ].filter(Boolean).join('\n\n')
}

async function runAgentCommand(command, rawRows, stats) {
  const agent = getAgentConfig()
  if (!agent.ready) return buildCommandResult(command, rawRows, stats)

  const tempDir = path.join(UPLOAD_DIR, '_agent_tmp')
  ensureDir(tempDir)
  const tempPath = path.join(tempDir, `cmd_${Date.now()}_${Math.random().toString(36).slice(2, 8)}.xlsx`)

  try {
    createWorkbookFromRows(tempPath, rawRows, 'SourceData')
    const agentPrompt = buildWorkbookAgentPrompt({
      request: command,
      rawRows,
      stats,
      preserveSourceSheet: true,
      allowSyntheticData: false
    })
    const agentResult = await runSpreadsheetAgent({
      provider: agent.provider,
      prompt: agentPrompt,
      filePath: tempPath
    })
    const preview = readWorkbookPreview(tempPath)
    if (preview.rows?.length > 1 && preview.sheetName !== 'SourceData') {
      return {
        action: 'sheet_result',
        description: agentResult.summary || `Created result sheet "${preview.sheetName}".`,
        answer: agentResult.summary,
        sheet_name: preview.sheetName,
        rows: preview.rows
      }
    }
    return buildCommandResult(command, rawRows, stats)
  } finally {
    try { if (fs.existsSync(tempPath)) fs.unlinkSync(tempPath) } catch (e) {}
  }
}

// ── AUTH ─────────────────────────────────────────────────────────
app.post('/api/auth/login', (req, res) => {
  const { email, password } = req.body
  const passwords = { 'admin@university.edu': 'admin123', 'krishna@university.edu': 'datum2026', 'demo@university.edu': 'demo123' }
  const user = usersDB.find(u => u.email === email)
  if (!user || passwords[email] !== password) return res.status(401).json({ error: 'Invalid credentials' })
  res.json({ user, token: `tok_${user.id}_${Date.now()}` })
})
app.post('/api/auth/register', (req, res) => {
  const { name, email, department, role = 'faculty' } = req.body
  if (usersDB.find(u => u.email === email)) return res.status(400).json({ error: 'Email already registered' })
  const user = { id: `u${Date.now()}`, name, email, role, department, plan: 'free' }
  usersDB.push(user); res.json({ user, token: `tok_${user.id}_${Date.now()}` })
})
app.get('/api/users', (req, res) => res.json(usersDB))

// ── FILES ─────────────────────────────────────────────────────────
app.post('/api/files/upload', (req, res, next) => {
  upload.array('files', 20)(req, res, err => {
    if (err) return res.status(400).json({ error: err.message })
    const dept = req.query.department || 'general'
    const uploaded = (req.files || []).map(f => {
      const record = {
        id: `f${Date.now()}_${Math.random().toString(36).slice(2, 6)}`,
        name: f.originalname, path: f.path, size: f.size,
        type: path.extname(f.originalname).toLowerCase().slice(1),
        department: dept, uploadedBy: req.query.uploadedBy || 'user',
        uploadedAt: new Date().toISOString(),
        url: `/uploads/${dept}/${f.filename}`
      }
      filesDB.push(record); return record
    })
    res.json({ files: uploaded })
  })
})
app.get('/api/files', (req, res) => {
  syncFilesDBFromDisk()
  const { department, type, search } = req.query
  let r = [...filesDB]
  if (department) r = r.filter(f => f.department === department)
  if (type) r = r.filter(f => f.type === type)
  if (search) r = r.filter(f => f.name.toLowerCase().includes(search.toLowerCase()))
  res.json(r)
})
app.delete('/api/files/:id', (req, res) => {
  const idx = filesDB.findIndex(f => f.id === req.params.id)
  if (idx === -1) return res.status(404).json({ error: 'Not found' })
  try { if (fs.existsSync(filesDB[idx].path)) fs.unlinkSync(filesDB[idx].path) } catch (e) {}
  filesDB.splice(idx, 1); res.json({ success: true })
})
app.get('/api/files/browse', (req, res) => {
  try {
    const dir = resolveUploadPath(req.query.path || '')
    const items = fs.readdirSync(dir, { withFileTypes: true }).map(e => {
      const full = path.join(dir, e.name)
      let size = 0, mtime = null
      try { const s = fs.statSync(full); size = s.size; mtime = s.mtime } catch (e) {}
      return { name: e.name, path: full, isDir: e.isDirectory(), size, type: e.isDirectory() ? 'folder' : path.extname(e.name).toLowerCase().slice(1), modified: mtime }
    })
    const parent = dir === UPLOAD_DIR ? UPLOAD_DIR : path.dirname(dir)
    res.json({ path: dir, parent, items })
  } catch (e) { res.status(400).json({ error: e.message }) }
})
app.post('/api/files/mkdir', (req, res) => {
  try {
    const target = resolveUploadPath(req.body.path || '')
    fs.mkdirSync(target, { recursive: true })
    res.json({ success: true, path: target })
  }
  catch (e) { res.status(400).json({ error: e.message }) }
})
app.get('/api/files/search-disk', (req, res) => {
  const query = req.query.query || ''
  const directory = req.query.directory || UPLOAD_DIR
  const results = []
  function walk(dir) {
    try {
      fs.readdirSync(dir, { withFileTypes: true }).forEach(e => {
        const full = path.join(dir, e.name)
        if (e.isDirectory()) walk(full)
        else if (e.name.toLowerCase().includes(query.toLowerCase())) {
          let size = 0, mtime = null
          try { const s = fs.statSync(full); size = s.size; mtime = s.mtime } catch (e) {}
          results.push({ name: e.name, path: full, size, type: path.extname(e.name).slice(1), modified: mtime })
        }
      })
    } catch (e) {}
  }
  walk(directory)
  res.json(results.slice(0, 500))
})

// ── EXCEL ─────────────────────────────────────────────────────────
app.get('/api/excel/analyze/:fileId', (req, res) => {
  const record = filesDB.find(f => f.id === req.params.fileId)
  if (!record) return res.status(404).json({ error: 'File not found' })
  const data = parseExcelFull(record.path)
  if (!data) return res.status(400).json({ error: 'Cannot parse' })
  const results = {}
  Object.entries(data).forEach(([name, sheet]) => { results[name] = computeStats(sheet.rawRows) })
  res.json({ sheets: results, sheetNames: Object.keys(data), rawData: data, file: record })
})
app.post('/api/excel/analyze', upload.single('file'), (req, res) => {
  const filePath = req.file?.path
  if (!filePath) return res.status(400).json({ error: 'No file' })
  const data = parseExcelFull(filePath)
  if (!data) return res.status(400).json({ error: 'Cannot parse' })
  const results = {}
  Object.entries(data).forEach(([name, sheet]) => { results[name] = computeStats(sheet.rawRows) })
  res.json({ sheets: results, sheetNames: Object.keys(data), rawData: data })
})
app.post('/api/excel/save/:fileId', (req, res) => {
  const record = filesDB.find(f => f.id === req.params.fileId)
  if (!record) return res.status(404).json({ error: 'Not found' })
  try {
    const wb = XLSX.readFile(record.path)
    const ws = XLSX.utils.aoa_to_sheet((req.body.rows || []).map(r => r.map(c => c.v !== undefined ? c.v : c)))
    wb.Sheets[req.body.sheetName || wb.SheetNames[0]] = ws
    XLSX.writeFile(wb, record.path)
    res.json({ success: true })
  } catch (e) { res.status(500).json({ error: e.message }) }
})
app.post('/api/excel/generate', (req, res) => {
  const { data, filename = 'output.xlsx', sheetName = 'Sheet1', department = 'general' } = req.body
  if (!data || !Array.isArray(data)) return res.status(400).json({ error: 'Invalid data' })
  const wb = XLSX.utils.book_new()
  const ws = Array.isArray(data[0]) ? XLSX.utils.aoa_to_sheet(data) : XLSX.utils.json_to_sheet(data)
  XLSX.utils.book_append_sheet(wb, ws, sheetName)
  const dir = path.join(UPLOAD_DIR, department)
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true })
  const outFile = `gen_${Date.now()}_${filename}`
  const outPath = path.join(dir, outFile)
  XLSX.writeFile(wb, outPath)
  const record = {
    id: `f${Date.now()}`, name: filename, path: outPath,
    size: fs.statSync(outPath).size, type: 'xlsx', department,
    uploadedBy: 'AI', uploadedAt: new Date().toISOString(),
    url: `/uploads/${department}/${outFile}`
  }
  filesDB.push(record)
  res.json({ file: record, downloadUrl: `http://localhost:${PORT}${record.url}` })
})
app.post('/api/excel/compare', (req, res) => {
  const { fileId1, fileId2 } = req.body
  const f1 = filesDB.find(f => f.id === fileId1)
  const f2 = filesDB.find(f => f.id === fileId2)
  if (!f1 || !f2) return res.status(404).json({ error: 'Files not found' })
  const d1 = parseExcelFull(f1.path), d2 = parseExcelFull(f2.path)
  const s1 = computeStats(Object.values(d1)[0].rawRows)
  const s2 = computeStats(Object.values(d2)[0].rawRows)
  res.json({ file1: { name: f1.name, stats: s1 }, file2: { name: f2.name, stats: s2 }, comparison: { avgDiff: +(s1.avgScore - s2.avgScore).toFixed(2), passRateDiff: +(s1.passRate - s2.passRate).toFixed(1) } })
})

// ── LLM ROUTES ────────────────────────────────────────────────────

app.post('/api/llm/save-keys', (req, res) => {
  const provider = String(req.body?.provider || 'gemini').toLowerCase() === 'claude' ? 'claude' : 'gemini'
  const currentEnv = readEnvFile(AGENT_ENV_FILE)
  const nextValues = {
    AI_PROVIDER: provider,
    GEMINI_API_KEY: req.body?.geminiKey !== undefined ? String(req.body.geminiKey || '').trim() : currentEnv.GEMINI_API_KEY,
    ANTHROPIC_API_KEY: req.body?.anthropicKey !== undefined ? String(req.body.anthropicKey || '').trim() : currentEnv.ANTHROPIC_API_KEY
  }

  writeEnvFile(AGENT_ENV_FILE, nextValues)
  res.json({ success: true, ...getAgentConfig() })
})

app.get('/api/llm/status', (req, res) => {
  res.json(getAgentConfig())
})

app.post('/api/llm/stream', async (req, res) => {
  const { message = '', rawRows, stats } = req.body || {}

  res.setHeader('Content-Type', 'text/event-stream')
  res.setHeader('Cache-Control', 'no-cache')
  res.setHeader('Connection', 'keep-alive')
  res.setHeader('Access-Control-Allow-Origin', '*')

  try {
    let text = buildChatResponse(message, rawRows, stats)
    const lower = String(message).toLowerCase().trim()

    if (rawRows?.length && lower && !/^(hi|hello|hey|hii|yo|ok|okay|kk|cool|fine|alright|thanks|thank you)\b/.test(lower)) {
      const result = await runAgentCommand(message, rawRows, stats)
      if (result?.action === 'summary') {
        text = result.answer || result.description || text
      } else if (result?.rows?.length > 1) {
        text = [result.description, result.answer].filter(Boolean).join('\n\n') || `I prepared ${result.rows.length - 1} row(s) from the current sheet.`
      } else if (result?.description) {
        text = result.description
      }
    }

    res.write(`data: ${JSON.stringify({ text })}\n\n`)
    res.write(`data: ${JSON.stringify({ done: true })}\n\n`)
    res.end()
  } catch (e) {
    res.write(`data: ${JSON.stringify({ error: e.message || 'Chat failed' })}\n\n`)
    res.end()
  }
})

app.post('/api/llm/command', async (req, res) => {
  const { command = '', rawRows, stats } = req.body || {}
  if (!rawRows?.length) return res.status(400).json({ error: 'No data loaded' })

  try {
    res.json(await runAgentCommand(command, rawRows, stats))
  } catch (e) {
    res.status(500).json({ error: e.message })
  }
})

app.post('/api/llm/formula', async (req, res) => {
  const { description = '', headers = [] } = req.body || {}
  try {
    res.json(generateFormulaSuggestion(description, headers))
  } catch (e) {
    res.status(500).json({ error: e.message })
  }
})

// Generate Excel from prompt
app.post('/api/llm/generate-excel', async (req, res) => {
  const { prompt: userPrompt, department = 'general', rawRows, stats } = req.body
  const agent = getAgentConfig()
  if (!agent.ready) return res.status(400).json({ error: 'AI agent is not configured. Add GEMINI_API_KEY or ANTHROPIC_API_KEY to Ai_Agent/.env.' })

  try {
    const dir = path.join(UPLOAD_DIR, department)
    ensureDir(dir)
    const workbookName = buildAiWorkbookName(userPrompt, 'ai_generated_sheet')
    const fname = `${Date.now()}_${workbookName}`
    const fpath = path.join(dir, fname)

    const agentResult = await runSpreadsheetAgent({
      provider: agent.provider,
      prompt: buildWorkbookAgentPrompt({
        request: userPrompt,
        rawRows,
        stats,
        preserveSourceSheet: false,
        allowSyntheticData: true
      }),
      filePath: fpath
    })

    const preview = readWorkbookPreview(fpath)

    const record = {
      id: `f${Date.now()}`, name: workbookName,
      path: fpath, size: fs.statSync(fpath).size, type: 'xlsx', department,
      uploadedBy: 'AI', uploadedAt: new Date().toISOString(),
      url: `/uploads/${department}/${fname}`
    }
    filesDB.push(record)
    syncFilesDBFromDisk()
    res.json({
      success: true,
      file: record,
      sheet_name: preview.sheetName,
      description: agentResult.summary,
      rowCount: Math.max((preview.rows || []).length - 1, 0),
      rows: preview.rows
    })
  } catch (e) {
    if (isQuotaExceededError(e) && rawRows?.length) {
      const fallback = buildCommandResult(userPrompt, rawRows, stats)
      if (fallback?.rows?.length > 1) {
        const workbookName = buildAiWorkbookName(userPrompt, 'sheet_result')
        const record = saveGeneratedWorkbook({
          department,
          name: workbookName,
          rows: fallback.rows,
          uploadedBy: 'AI Fallback',
          sheetName: fallback.sheet_name || 'Result'
        })
        const preview = readWorkbookPreview(record.path)
        return res.json({
          success: true,
          file: record,
          sheet_name: preview.sheetName,
          description: `${fallback.description || 'Created a result sheet from the current workbook.'} Gemini quota was exceeded, so this result was generated locally from your current sheet.`,
          rowCount: Math.max((preview.rows || []).length - 1, 0),
          rows: preview.rows,
          fallback: true
        })
      }
      return res.status(429).json({
        error: `Gemini quota exceeded. Try again later${e?.retryAfterSeconds ? ` in about ${e.retryAfterSeconds}s` : ''}, switch provider, or request a simpler local result.`
      })
    }
    res.status(e?.status === 429 ? 429 : 500).json({
      error: e?.status === 429
        ? `Gemini quota exceeded. Try again later${e?.retryAfterSeconds ? ` in about ${e.retryAfterSeconds}s` : ''}.`
        : e.message
    })
  }
})

// New sheet from existing data
app.post('/api/llm/new-sheet', async (req, res) => {
  const { instruction, rawRows, stats, department = 'general' } = req.body
  const agent = getAgentConfig()
  if (!agent.ready) return res.status(400).json({ error: 'AI agent is not configured. Add GEMINI_API_KEY or ANTHROPIC_API_KEY to Ai_Agent/.env.' })
  if (!rawRows || !rawRows.length) return res.status(400).json({ error: 'No data' })

  try {
    const dir = path.join(UPLOAD_DIR, department)
    ensureDir(dir)
    const workbookName = buildAiWorkbookName(instruction, 'ai_sheet')
    const fname = `${Date.now()}_${workbookName}`
    const fpath = path.join(dir, fname)
    createWorkbookFromRows(fpath, rawRows, 'SourceData')

    const agentResult = await runSpreadsheetAgent({
      provider: agent.provider,
      prompt: buildWorkbookAgentPrompt({
        request: instruction,
        rawRows,
        stats,
        preserveSourceSheet: true,
        allowSyntheticData: false
      }),
      filePath: fpath
    })

    const preview = readWorkbookPreview(fpath)

    const record = {
      id: `f${Date.now()}`, name: workbookName,
      path: fpath, size: fs.statSync(fpath).size, type: 'xlsx', department,
      uploadedBy: 'AI', uploadedAt: new Date().toISOString(),
      url: `/uploads/${department}/${fname}`
    }
    filesDB.push(record)
    syncFilesDBFromDisk()
    res.json({
      success: true,
      file: record,
      sheet_name: preview.sheetName,
      description: agentResult.summary,
      rowCount: Math.max((preview.rows || []).length - 1, 0),
      rows: preview.rows
    })
  } catch (e) {
    if (isQuotaExceededError(e)) {
      const fallback = buildCommandResult(instruction, rawRows, stats)
      if (fallback?.rows?.length > 1) {
        const workbookName = buildAiWorkbookName(instruction, 'sheet_result')
        const record = saveGeneratedWorkbook({
          department,
          name: workbookName,
          rows: fallback.rows,
          uploadedBy: 'AI Fallback',
          sheetName: fallback.sheet_name || 'Result'
        })
        const preview = readWorkbookPreview(record.path)
        return res.json({
          success: true,
          file: record,
          sheet_name: preview.sheetName,
          description: `${fallback.description || 'Created a result sheet from the current workbook.'} Gemini quota was exceeded, so this result was generated locally from your current sheet.`,
          rowCount: Math.max((preview.rows || []).length - 1, 0),
          rows: preview.rows,
          fallback: true
        })
      }
      return res.status(429).json({
        error: 'Gemini quota exceeded and no local fallback result matched this request. Try again later or switch provider.'
      })
    }
    res.status(500).json({ error: e.message })
  }
})

// Formula generator


// ── WORKFLOWS ─────────────────────────────────────────────────────
app.get('/api/workflows', (req, res) => res.json(workflowsDB))
app.post('/api/workflows', (req, res) => {
  const w = { id: `w${Date.now()}`, ...req.body, status: 'paused', runs: 0, lastRun: 'Never' }
  workflowsDB.push(w); res.json(w)
})
app.patch('/api/workflows/:id', (req, res) => {
  const idx = workflowsDB.findIndex(w => w.id === req.params.id)
  if (idx === -1) return res.status(404).json({ error: 'Not found' })
  workflowsDB[idx] = { ...workflowsDB[idx], ...req.body }; res.json(workflowsDB[idx])
})
app.delete('/api/workflows/:id', (req, res) => {
  workflowsDB = workflowsDB.filter(w => w.id !== req.params.id); res.json({ success: true })
})
app.post('/api/workflows/:id/run', (req, res) => {
  const idx = workflowsDB.findIndex(w => w.id === req.params.id)
  if (idx === -1) return res.status(404).json({ error: 'Not found' })
  workflowsDB[idx].lastRun = new Date().toLocaleDateString(); workflowsDB[idx].runs++
  res.json({ success: true, workflow: workflowsDB[idx] })
})

// ── DOCUMENTS ─────────────────────────────────────────────────────
app.post('/api/documents/generate', async (req, res) => {
  const { type = 'report', data = {}, title } = req.body
  const docTitle = title || slugifyName(data?.prompt || type, 'Report').replace(/_/g, ' ')

  let content = `${docTitle}\nGenerated: ${new Date().toLocaleString()}\n${'='.repeat(50)}\n\n`
  if (data?.prompt) content += `Request:\n${data.prompt}\n\n`
  if (data?.notes) content += `Context:\n${data.notes}\n\n`
  content += `Draft Body:\nThis ${type} was prepared for ${data?.department || 'the department'}`
  if (data?.user) content += ` by request of ${data.user}`
  content += `. Review the points below and refine as needed.\n\n`
  content += `1. Overview of the current spreadsheet context.\n`
  content += `2. Key findings, risks, and recommendations.\n`
  content += `3. Suggested follow-up actions and approvals.\n`

  const file = saveGeneratedAsset({
    department: data.department || 'general',
    name: buildAiTextName(docTitle, 'ai_document', 'txt'),
    content
  })

  res.json({ id: `doc${Date.now()}`, title: docTitle, type, content, createdAt: new Date().toISOString(), file })
})

app.post('/api/ppt/generate', async (req, res) => {
  const {
    prompt: userPrompt,
    data = {},
    audience = 'Management',
    theme = 'Executive Blue',
    goal = 'Explain performance and recommendations',
    department = 'general'
  } = req.body
  const defaultDeckTitle = userPrompt || 'AI Presentation'

  const slides = [
    { title: 'Overview', content: userPrompt || 'Presentation overview', chart: 'none', notes: 'Introduce the topic and goal.' },
    { title: 'Key Findings', content: data.notes || 'Summarize findings from the workbook.', chart: 'bar', notes: 'Explain the main metrics and trends.' },
    { title: 'Recommendations', content: `Actions for ${audience} using the ${theme} theme.`, chart: 'line', notes: 'Close with decisions and next steps.' }
  ]
  const file = saveGeneratedAsset({
    department,
    name: buildAiTextName(defaultDeckTitle, 'ai_presentation', 'txt'),
    content: JSON.stringify({ deck_title: defaultDeckTitle, subtitle: goal, slides }, null, 2)
  })

  res.json({ deck_title: defaultDeckTitle, subtitle: goal, slides, file })
})

// ── DASHBOARD ─────────────────────────────────────────────────────
app.get('/api/dashboard/stats', (req, res) => res.json({
  totalFiles: filesDB.length, totalUsers: usersDB.length,
  totalWorkflows: workflowsDB.length,
  activeWorkflows: workflowsDB.filter(w => w.status === 'active').length,
  recentFiles: filesDB.slice(-5).reverse(),
  departments: [...new Set(usersDB.map(u => u.department))]
}))

app.get('/api/health', (req, res) => res.json({ status: 'ok', timestamp: new Date() }))

app.listen(PORT, () => console.log(`\n✓ Server running → http://localhost:${PORT}\n`))
