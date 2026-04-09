const fs = require('fs')
const path = require('path')
const { spawn } = require('child_process')
const XLSX = require('xlsx')

const ROOT_DIR = path.resolve(__dirname, '..')
const AGENT_DIR = path.join(ROOT_DIR, 'Ai_Agent')

function ensureDir(dir) {
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true })
}

function createWorkbookFromRows(filePath, rows, sheetName = 'SourceData') {
  const workbook = XLSX.utils.book_new()
  const worksheet = XLSX.utils.aoa_to_sheet(rows || [])
  XLSX.utils.book_append_sheet(workbook, worksheet, sheetName)
  XLSX.writeFile(workbook, filePath)
}

function readWorkbookPreview(filePath) {
  const workbook = XLSX.readFile(filePath, { cellDates: true })
  const sheetNames = workbook.SheetNames || []
  const targetSheet = sheetNames[sheetNames.length - 1] || 'Sheet1'
  const rows = targetSheet && workbook.Sheets[targetSheet]
    ? XLSX.utils.sheet_to_json(workbook.Sheets[targetSheet], { header: 1, defval: '' })
    : []

  return {
    sheetNames,
    sheetName: targetSheet,
    rows
  }
}

function stripAnsi(value = '') {
  return String(value).replace(/\x1B\[[0-9;]*m/g, '')
}

function parseAgentError(stderr = '', stdout = '', code) {
  const cleanStderr = stripAnsi(stderr)
    .split(/\r?\n/)
    .filter(line => !/injected env|tip:|dotenv/i.test(line))
    .join('\n')
    .trim()
  const cleanStdout = stripAnsi(stdout)
    .split(/\r?\n/)
    .filter(line => !/injected env|tip:|dotenv/i.test(line))
    .join('\n')
    .trim()
  const combined = [cleanStderr, cleanStdout].filter(Boolean).join('\n')
  const quotaMatch = combined.match(/"code"\s*:\s*429|RESOURCE_EXHAUSTED|quota exceeded|rate.?limit/i)
  const retryMatch = combined.match(/retry in ([0-9.]+)s/i)
  const lines = combined.split(/\r?\n/).map(line => line.trim()).filter(Boolean)
  const meaningfulLine = [...lines].reverse().find(line => !/^\[(agent|context|think|warn|error|save|done|tool|result)\]/i.test(line))
  const message = meaningfulLine || combined || `Agent exited with code ${code}`
  const error = new Error(message)
  if (quotaMatch) {
    error.status = 429
    error.code = 'RESOURCE_EXHAUSTED'
    error.isQuotaExceeded = true
    if (retryMatch) error.retryAfterSeconds = Math.ceil(Number(retryMatch[1]))
  }
  error.agentLog = combined
  return error
}

function runSpreadsheetAgent({ provider = 'gemini', prompt, filePath, geminiKey = '', anthropicKey = '' }) {
  return new Promise((resolve, reject) => {
    const scriptName = provider === 'claude' ? 'agent_claude.mjs' : 'agent_gemini.mjs'
    const scriptPath = path.join(AGENT_DIR, scriptName)
    const env = {
      ...process.env,
      FORCE_COLOR: '0',
      NO_COLOR: '1'
    }
    if (geminiKey) env.GEMINI_API_KEY = geminiKey
    if (anthropicKey) env.ANTHROPIC_API_KEY = anthropicKey

    const child = spawn(process.execPath, [scriptPath, prompt, filePath], {
      cwd: AGENT_DIR,
      env,
      stdio: ['ignore', 'pipe', 'pipe']
    })

    let stdout = ''
    let stderr = ''
    child.stdout.on('data', chunk => { stdout += chunk.toString() })
    child.stderr.on('data', chunk => { stderr += chunk.toString() })
    child.on('error', reject)
    child.on('close', code => {
      if (code !== 0) {
        reject(parseAgentError(stderr, stdout, code))
        return
      }
      resolve({
        stdout: stdout.trim(),
        stderr: stderr.trim(),
        summary: extractSummary(stdout)
      })
    })
  })
}

function extractSummary(output = '') {
  const lines = String(output).split(/\r?\n/).map(line => line.trim()).filter(Boolean)
  const lastMeaningfulLine = [...lines].reverse().find(line => !/^\[(tool|think|context)\]/i.test(line))
  return lastMeaningfulLine || 'Spreadsheet agent completed successfully.'
}

module.exports = {
  AGENT_DIR,
  ensureDir,
  createWorkbookFromRows,
  readWorkbookPreview,
  runSpreadsheetAgent
}
