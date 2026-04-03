const express = require('express')
const cors = require('cors')
const multer = require('multer')
const path = require('path')
const fs = require('fs')
const XLSX = require('xlsx')
const { callLLM, streamLLM } = require('./llm-providers')

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
let savedKeys = {
  gemini: process.env.GEMINI_API_KEY || '',
  hf: process.env.HF_API_KEY || ''
}

function getKeys(body = {}) {
  return {
    gemini: body.geminiKey || body.apiKey || savedKeys.gemini || process.env.GEMINI_API_KEY || '',
    hf: body.hfKey || savedKeys.hf || process.env.HF_API_KEY || ''
  }
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

// Save keys
app.post('/api/llm/save-keys', (req, res) => {
  const { geminiKey, hfKey } = req.body
  if (geminiKey) savedKeys.gemini = geminiKey.trim()
  if (hfKey) savedKeys.hf = hfKey.trim()
  res.json({ success: true, gemini: !!savedKeys.gemini, hf: !!savedKeys.hf })
})

app.get('/api/llm/key-status', (req, res) => {
  res.json({ gemini: !!savedKeys.gemini, hf: !!savedKeys.hf })
})

// Stream chat
app.post('/api/llm/stream', async (req, res) => {
  const { message, rawRows, stats } = req.body
  const keys = getKeys(req.body)

  res.setHeader('Content-Type', 'text/event-stream')
  res.setHeader('Cache-Control', 'no-cache')
  res.setHeader('Connection', 'keep-alive')
  res.setHeader('Access-Control-Allow-Origin', '*')

  if (!keys.gemini && !keys.hf) {
    res.write(`data: ${JSON.stringify({ error: 'No API keys saved. Add Gemini or HuggingFace key in Settings tab.' })}\n\n`)
    return res.end()
  }

  const ctx = buildDataContext(rawRows, stats)
  const prompt = `${ctx}\n\nUser: ${message}\n\nAnswer specifically using the data above. Be helpful and direct.`

  await streamLLM(
    prompt,
    'You are GPT-EXCEL, expert AI for student data analysis. Be concise and use actual data.',
    keys,
    (text) => res.write(`data: ${JSON.stringify({ text })}\n\n`),
    () => { res.write(`data: ${JSON.stringify({ done: true })}\n\n`); res.end() },
    (err) => { res.write(`data: ${JSON.stringify({ error: err })}\n\n`); res.end() }
  )
})

// AI command
app.post('/api/llm/command', async (req, res) => {
  const { command, rawRows, stats } = req.body
  const keys = getKeys(req.body)

  if (!rawRows?.length) return res.status(400).json({ error: 'No data loaded' })
  if (!keys.gemini && !keys.hf) return res.status(400).json({ error: 'No API keys configured. Add keys in Settings.' })

  const headers = rawRows[0].map(h => String(h || ''))
  const ctx = buildDataContext(rawRows, stats)

  const prompt = `${ctx}

User command: "${command}"

Return ONE JSON action. Choose the most appropriate:

For filtering students:
{"action":"filter","description":"what was done","rows":[[header1,header2,...],[val1,val2,...],...]}

For sorting:
{"action":"sort","description":"what was done","rows":[[header1,header2,...],[val1,val2,...],...]}

For adding a column:
{"action":"add_column","description":"what was done","column_name":"NewCol","rows":[[header1,header2,...,NewCol],[val1,val2,...,computed],...]}

For text answer / summary:
{"action":"summary","description":"answer","answer":"plain text answer"}

Rules:
- Return ONLY valid JSON, nothing else
- Include header row as first row in rows array
- Include ALL matching rows (not just samples)
- If the user is asking to search for a person, tell details about someone, answer "who", or wants names/details only, prefer "summary" with a direct natural-language answer
- Only return rows when the user explicitly asks to show a table, filter rows, sort data, add a column, export, or create a sheet-like result
- For summary answers about matched people, include the important row values in plain English instead of creating a new sheet
- Available columns: ${headers.join(', ')}
- pass threshold is 40 marks
- below 75 means attendance < 75`

  try {
    const result = await callLLM(prompt, null, true, keys)
    res.json(result)
  } catch (e) {
    res.status(500).json({ error: e.message })
  }
})

// Generate Excel from prompt
app.post('/api/llm/generate-excel', async (req, res) => {
  const { prompt: userPrompt, department = 'general', rawRows, stats } = req.body
  const keys = getKeys(req.body)
  if (!keys.gemini && !keys.hf) return res.status(400).json({ error: 'No API keys configured' })

  const ctx = rawRows && rawRows.length ? `\nExisting data:\n${buildDataContext(rawRows, stats)}\nCreate a new sheet based on this data as requested.\n` : ''

  const prompt = `${ctx}
User wants: "${userPrompt}"

Generate spreadsheet data as JSON:
{
  "sheet_name": "SheetName",
  "description": "What this contains",
  "rows": [["Header1","Header2","Header3"],["val1","val2","val3"]]
}

Rules:
- Header row first
- At least 15 realistic rows
- Use realistic Indian student names
- For student data: Roll No, Name, Section, subject marks 0-100, Attendance 60-100
- Numbers must vary realistically
- Return ONLY valid JSON`

  try {
    const result = await callLLM(prompt, null, true, keys)
    if (!result.rows || !result.rows.length) throw new Error('No rows generated')

    const wb = XLSX.utils.book_new()
    const ws = XLSX.utils.aoa_to_sheet(result.rows)
    XLSX.utils.book_append_sheet(wb, ws, result.sheet_name || 'Sheet1')

    const dir = path.join(UPLOAD_DIR, department)
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true })
    const workbookName = buildAiWorkbookName(result.sheet_name || userPrompt, 'ai_generated_sheet')
    const fname = `${Date.now()}_${workbookName}`
    const fpath = path.join(dir, fname)
    XLSX.writeFile(wb, fpath)

    const record = {
      id: `f${Date.now()}`, name: workbookName,
      path: fpath, size: fs.statSync(fpath).size, type: 'xlsx', department,
      uploadedBy: 'AI', uploadedAt: new Date().toISOString(),
      url: `/uploads/${department}/${fname}`
    }
    filesDB.push(record)
    syncFilesDBFromDisk()
    res.json({ success: true, file: record, sheet_name: result.sheet_name, description: result.description, rowCount: result.rows.length - 1, rows: result.rows })
  } catch (e) {
    res.status(500).json({ error: e.message })
  }
})

// New sheet from existing data
app.post('/api/llm/new-sheet', async (req, res) => {
  const { instruction, rawRows, stats, department = 'general' } = req.body
  const keys = getKeys(req.body)
  if (!keys.gemini && !keys.hf) return res.status(400).json({ error: 'No API keys configured' })
  if (!rawRows || !rawRows.length) return res.status(400).json({ error: 'No data' })

  const ctx = buildDataContext(rawRows, stats)
  const prompt = `${ctx}

Instruction: "${instruction}"

Transform the data above into a new spreadsheet. Return JSON:
{
  "sheet_name": "Name",
  "description": "What this contains",
  "rows": [["Header1","Header2",...],["val1","val2",...],...]
}

Examples:
- "top 10 students" → sort by avg, take top 10
- "failed students" → only failed students
- "section summary" → section-wise aggregated stats
- "add Grade column" → all original rows + Grade column computed from avg
- "below 75% attendance" → only those students

Include ALL relevant rows. Return ONLY valid JSON.`

  try {
    const result = await callLLM(prompt, null, true, keys)
    if (!result.rows || !result.rows.length) throw new Error('No rows')

    const wb = XLSX.utils.book_new()
    const ws = XLSX.utils.aoa_to_sheet(result.rows)
    XLSX.utils.book_append_sheet(wb, ws, result.sheet_name || 'Result')
    const dir = path.join(UPLOAD_DIR, department)
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true })
    const workbookName = buildAiWorkbookName(result.sheet_name || instruction, 'ai_sheet')
    const fname = `${Date.now()}_${workbookName}`
    const fpath = path.join(dir, fname)
    XLSX.writeFile(wb, fpath)

    const record = {
      id: `f${Date.now()}`, name: workbookName,
      path: fpath, size: fs.statSync(fpath).size, type: 'xlsx', department,
      uploadedBy: 'AI', uploadedAt: new Date().toISOString(),
      url: `/uploads/${department}/${fname}`
    }
    filesDB.push(record)
    syncFilesDBFromDisk()
    res.json({ success: true, file: record, sheet_name: result.sheet_name, description: result.description, rowCount: result.rows.length - 1, rows: result.rows })
  } catch (e) {
    res.status(500).json({ error: e.message })
  }
})

// Formula generator
app.post('/api/llm/formula', async (req, res) => {
  const { description, headers } = req.body
  const keys = getKeys(req.body)
  if (!keys.gemini && !keys.hf) return res.status(400).json({ error: 'No API keys configured' })

  const prompt = `Excel formula expert. Available columns: ${(headers || []).join(', ')}
User wants: "${description}"

Return JSON:
{
  "formula": "=FORMULA()",
  "explanation": "plain English",
  "example": "example with numbers",
  "alternatives": ["=ALT1()"]
}
Return ONLY valid JSON.`

  try {
    const result = await callLLM(prompt, null, true, keys)
    res.json(result)
  } catch (e) { res.status(500).json({ error: e.message }) }
})

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
  const keys = getKeys(req.body)
  const docTitle = title || slugifyName(data?.prompt || type, 'Report').replace(/_/g, ' ')

  try {
    if (keys.gemini || keys.hf) {
      const prompt = `Create a well-structured ${type} based on the spreadsheet context below.

User request:
${data.prompt || `Create a ${type}`}

Workbook context:
${data.notes || 'No extra notes provided'}

Department: ${data.department || 'general'}
Requested by: ${data.user || 'user'}

Return ONLY valid JSON:
{
  "title": "Document title",
  "content": "Full document content in plain text with headings, sections, and polished writing"
}

Requirements:
- Make the title professional and specific
- Make the content detailed, useful, and ready to use
- Use structure and concrete workbook details when available
- Avoid placeholders
- No markdown fences`

      const result = await callLLM(prompt, null, true, keys)
      const generatedTitle = result.title || docTitle
      const generatedContent = result.content || ''
      const file = saveGeneratedAsset({
        department: data.department || 'general',
        name: buildAiTextName(generatedTitle, 'ai_document', 'txt'),
        content: generatedContent
      })

      return res.json({
        id: `doc${Date.now()}`,
        title: generatedTitle,
        type,
        content: generatedContent,
        createdAt: new Date().toISOString(),
        file
      })
    }
  } catch (e) {
    console.error('Document generation failed:', e)
  }

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
  const keys = getKeys(req.body)
  const defaultDeckTitle = userPrompt || 'AI Presentation'

  try {
    if (keys.gemini || keys.hf) {
      const prompt = `Create a presentation outline from spreadsheet data.

User request:
${userPrompt}

Workbook context:
${data.notes || 'No workbook notes provided'}

Audience: ${audience}
Theme: ${theme}
Goal: ${goal}

Return ONLY valid JSON:
{
  "deck_title": "Presentation title",
  "subtitle": "Short subtitle",
  "slides": [
    {
      "title": "Slide title",
      "content": "Slide body text",
      "chart": "bar|line|pie|none",
      "notes": "Speaker notes"
    }
  ]
}

Requirements:
- Create 5 to 7 slides
- Make the deck data-driven and structured for the chosen audience
- Use charts when supported by the context
- Keep notes concise and useful
- No markdown fences`

      const result = await callLLM(prompt, null, true, keys)
      const deckTitle = result.deck_title || defaultDeckTitle
      const slides = Array.isArray(result.slides) ? result.slides : []
      const file = saveGeneratedAsset({
        department,
        name: buildAiTextName(deckTitle, 'ai_presentation', 'txt'),
        content: JSON.stringify({ deck_title: deckTitle, subtitle: result.subtitle || '', slides }, null, 2)
      })

      return res.json({ deck_title: deckTitle, subtitle: result.subtitle || '', slides, file })
    }
  } catch (e) {
    console.error('PPT generation failed:', e)
  }

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
