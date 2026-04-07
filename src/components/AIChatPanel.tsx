import { memo, useState, useRef, useEffect, useCallback } from 'react'
import { useSelector } from 'react-redux'
import { useNavigate } from 'react-router-dom'
import { RootState } from '../index'

interface Msg {
  id: string
  role: 'user' | 'ai' | 'system'
  content: string
  rows?: any[][]
  isStreaming?: boolean
  time: string
}

interface Props {
  currentFile?: any
  rawRows?: any[][]
  stats?: any
  offlineMode?: boolean
  panelWidth?: number
  initialCollapsed?: boolean
  offlineCommands?: Array<{ command: string; effect: string }>
  onGridUpdate?: (rows: string[][], desc: string) => void
  onNewFile?: (file: any, rows: any[][]) => void
  onShowChart?: (config: { chartType: string, chartDataKey: string, title: string }) => void
  onWorkspaceIntent?: (intent: { tab: 'analyze' | 'charts' | 'create'; prompt: string; questions: string[] }) => void
  onOfflineCommand?: (command: string) => void
  onApplyPreview?: () => void
  onDiscardPreview?: () => void
  onUndoLastChange?: () => void
}

const BASE = 'http://localhost:3001/api/llm'

const QUICK = [
  'top 10 students', 'show failed', 'section stats',
  'below 75% attendance', 'subject averages', 'grade distribution',
  'sort by average', 'show passed', 'explain this data', 'who is at risk'
]

let _uid = 0
const uid = () => `${++_uid}_${Date.now()}`
const now = () => new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })
const asText = (value: any): string => {
  if (typeof value === 'string') return value
  if (value == null) return ''
  if (typeof value === 'number' || typeof value === 'boolean') return String(value)
  if (Array.isArray(value)) return value.map(asText).filter(Boolean).join(' ')
  if (typeof value === 'object') {
    if (typeof value.text === 'string') return value.text
    if (typeof value.content === 'string') return value.content
    if (Array.isArray(value.content)) return asText(value.content)
    try { return JSON.stringify(value, null, 2) } catch { return String(value) }
  }
  return String(value)
}

const buildSheetContext = (rawRows?: any[][], stats?: any) => {
  const headers = rawRows?.[0]?.map((h: any) => String(h || '')) || []
  const sampleRows = (rawRows || []).slice(1, 6)
  return {
    headers,
    rowCount: Math.max((rawRows?.length || 1) - 1, 0),
    sampleRows,
    stats: stats || null
  }
}

const normalizeHeader = (value: any) => String(value || '').toLowerCase().replace(/[^a-z0-9]/g, '')
const normalizeSheetRows = (rows?: any[][]) => (rows || []).map(row => row.map(cell => cell == null ? '' : String(cell)))
const findNameColumnIndex = (headers: string[]) => headers.findIndex(header => /name|student/i.test(header))
const findIdentifierColumnIndex = (headers: string[]) => headers.findIndex(header => /roll|id|reg|admission/i.test(header))
const formatRowSummary = (headers: string[], row: string[]) => headers.map((header, index) => `${header}: ${row[index] ?? ''}`).filter(entry => !entry.endsWith(': ')).join(', ')

const parseRowValues = (input: string, headers: string[], baseRow?: string[]) => {
  const nextRow = baseRow ? [...baseRow] : Array(headers.length).fill('')
  const chunks = input.split(/[;,]\s*/).map(chunk => chunk.trim()).filter(Boolean)
  let matched = 0
  let sequentialIndex = 0

  for (const chunk of chunks) {
    const keyed = chunk.match(/^([^:=]+)\s*[:=]\s*(.+)$/)
    if (keyed) {
      const idx = headers.findIndex(header => normalizeHeader(header) === normalizeHeader(keyed[1]))
      if (idx >= 0) {
        nextRow[idx] = keyed[2].trim()
        matched++
        continue
      }
    }
    while (sequentialIndex < headers.length && baseRow && nextRow[sequentialIndex] !== baseRow[sequentialIndex]) sequentialIndex++
    if (sequentialIndex < headers.length) {
      nextRow[sequentialIndex] = chunk
      sequentialIndex++
      matched++
    }
  }

  return { row: nextRow, matched }
}

const applyLocalLookupCommand = (command: string, sourceRows?: any[][]) => {
  const rows = normalizeSheetRows(sourceRows)
  if (rows.length < 2) return null
  const headers = rows[0]
  const dataRows = rows.slice(1)
  const searchMatch = command.match(/\b(find|search|lookup|who\s+is|details\s+for|tell\s+me\s+about)\b\s+(.+)$/i)
  if (!searchMatch) return null

  const rawQuery = searchMatch[2]
    .replace(/\b(in|from|on)\s+(the\s+)?(sheet|excel|table|file)\b/gi, '')
    .replace(/\b(student|name|row|details)\b/gi, '')
    .trim()
  if (!rawQuery) return null

  const query = rawQuery.toLowerCase()
  const matchedRows = dataRows.filter(row => row.some(cell => String(cell || '').toLowerCase().includes(query)))
  if (!matchedRows.length) return { message: `I couldn't find "${rawQuery}" in the current sheet.` }

  const nameIdx = findNameColumnIndex(headers)
  const idIdx = findIdentifierColumnIndex(headers)
  const topMatches = matchedRows.slice(0, 10)
  const summary = topMatches.slice(0, 3).map(row => {
    const label = nameIdx >= 0 ? row[nameIdx] : row[0]
    const identifier = idIdx >= 0 ? ` (${headers[idIdx]}: ${row[idIdx]})` : ''
    return `${label}${identifier}: ${formatRowSummary(headers, row)}`
  }).join('\n')

  return {
    message: matchedRows.length === 1
      ? `I found 1 matching row for "${rawQuery}".\n\n${summary}`
      : `I found ${matchedRows.length} matching rows for "${rawQuery}". Here are the closest matches.\n\n${summary}`,
    rows: [headers, ...topMatches]
  }
}

const applyLocalSheetCommand = (command: string, sourceRows?: any[][]) => {
  const rows = normalizeSheetRows(sourceRows)
  if (rows.length < 1) return null
  const headers = rows[0]
  const addMatch = command.match(/\b(add|insert)\b(?:\s+(?:new\s+)?(?:row|line))?(?:\s+(?:after\s+row\s+|after\s+line\s+|at\s+row\s+|at\s+line\s+)?(\d+)|\s+(?:at|to)\s+(last\s+line|last\s+row|end))?(?:\s*[:,-]?\s*)(.*)$/i)
  if (addMatch) {
    const insertAfter = addMatch[2] ? Number(addMatch[2]) : rows.length - 1
    const details = (addMatch[4] || '').trim() || command.replace(/\b(add|insert)\b/i, '').replace(/\b(at|to)\s+(last\s+line|last\s+row|end)\b/i, '').trim()
    const nameIdx = findNameColumnIndex(headers)
    const { row, matched } = parseRowValues(details, headers)
    const finalRow = [...row]
    const finalMatched = matched || (details && nameIdx >= 0 ? 1 : 0)
    if (!matched && details && nameIdx >= 0) finalRow[nameIdx] = details
    if (!finalMatched) return { error: 'Use something like `add Krishna at last line` or `add row Name=Arun, Section=A` so I can build a preview.' }
    const nextRows = rows.map(rowData => [...rowData])
    const insertIndex = Math.min(Math.max(insertAfter, 1), nextRows.length)
    nextRows.splice(insertIndex, 0, finalRow)
    return { rows: nextRows, message: `Preview ready — added a row ${addMatch[2] ? `after row ${addMatch[2]}` : 'at the end of the sheet'}.` }
  }

  const removeMatch = command.match(/\b(remove|delete)\s+(?:row|line)\s+(\d+)\b/i)
  if (removeMatch) {
    const rowNumber = Number(removeMatch[2])
    if (rowNumber <= 1 || rowNumber >= rows.length + 1) {
      return { error: 'Choose a visible data row number greater than 1 so I can preview the deletion.' }
    }
    const nextRows = rows.map(rowData => [...rowData])
    nextRows.splice(rowNumber - 1, 1)
    return { rows: nextRows, message: `Preview ready — removed row ${rowNumber}. Apply it to make the deletion permanent.` }
  }

  const updateMatch = command.match(/\b(update|edit|change|set)\s+(?:row|line)\s+(\d+)\b[\s:-]*(.*)$/i)
  if (updateMatch) {
    const rowNumber = Number(updateMatch[2])
    if (rowNumber <= 1 || rowNumber >= rows.length + 1) {
      return { error: 'Choose a visible data row number greater than 1 so I can preview the update.' }
    }
    const details = (updateMatch[3] || '').trim()
    const targetRow = [...rows[rowNumber - 1]]
    const { row, matched } = parseRowValues(details, headers, targetRow)
    if (!matched) return { error: 'Use something like `update row 4 Status=Done, Days=22` so I can preview the change.' }
    const nextRows = rows.map(rowData => [...rowData])
    nextRows[rowNumber - 1] = row
    return { rows: nextRows, message: `Preview ready — updated row ${rowNumber}. Review it, then apply or undo from chat.` }
  }

  return null
}

const MiniTable = ({ rows, onExport, onView }: { rows: any[][], onExport: () => void, onView?: () => void }) => {
  if (!rows || rows.length < 2) return null
  const headers = rows[0]
  const data = rows.slice(1, 41)
  const visibleCols = headers.slice(0, 6)
  return (
    <div style={{ marginTop: 8, border: '1px solid var(--border)', borderRadius: 8, overflow: 'hidden', fontSize: '0.7rem' }}>
      <div style={{ overflowX: 'auto', maxHeight: 220 }}>
        <table style={{ width: '100%', borderCollapse: 'collapse' }}>
          <thead>
            <tr style={{ background: 'var(--surface-2)', position: 'sticky', top: 0 }}>
              {visibleCols.map((h: any, i: number) => (
                <th key={i} style={{ padding: '5px 8px', textAlign: 'left', fontWeight: 700, color: 'var(--text-muted)', borderBottom: '1px solid var(--border)', whiteSpace: 'nowrap', fontSize: '0.65rem', textTransform: 'uppercase' }}>{String(h)}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {data.map((row: any[], ri: number) => (
              <tr key={ri} style={{ borderBottom: '1px solid var(--border)' }}>
                {visibleCols.map((_: any, ci: number) => (
                  <td key={ci} style={{ padding: '4px 8px', color: 'var(--text-secondary)', whiteSpace: 'nowrap' }}>{String(row[ci] ?? '')}</td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '4px 8px', background: 'var(--surface-2)', borderTop: '1px solid var(--border)' }}>
        <span style={{ fontSize: '0.62rem', color: 'var(--text-muted)' }}>{data.length} rows{headers.length > 6 ? ` · ${headers.length - 6} more cols hidden` : ''}</span>
        <div style={{ display: 'flex', gap: 6 }}>
          {onView && (
            <button onClick={onView} style={{ background: 'none', border: '1px solid var(--blue)', borderRadius: 4, padding: '2px 8px', fontSize: '0.62rem', color: 'var(--blue)', cursor: 'pointer' }}>View Preview</button>
          )}
          <button onClick={onExport} style={{ background: 'none', border: '1px solid var(--border)', borderRadius: 4, padding: '2px 8px', fontSize: '0.62rem', color: 'var(--text-muted)', cursor: 'pointer' }}>↓ Export .xlsx</button>
        </div>
      </div>
    </div>
  )
}

const MD = ({ text }: { text: string }) => {
  const html = text
    .replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>')
    .replace(/^### (.+)$/gm, '<div style="font-weight:700;font-size:0.82rem;margin:8px 0 3px">$1</div>')
    .replace(/^## (.+)$/gm, '<div style="font-weight:700;font-size:0.86rem;margin:10px 0 4px">$1</div>')
    .replace(/^[-•] (.+)$/gm, '<div style="display:flex;gap:5px;margin:2px 0"><span style="color:var(--accent)">•</span><span>$1</span></div>')
    .replace(/`([^`]+)`/g, '<code style="background:var(--surface-2);padding:1px 4px;border-radius:3px;font-size:0.78em;color:var(--green)">$1</code>')
    .replace(/\n\n/g, '<br/><br/>').replace(/\n/g, '<br/>')
  return <div dangerouslySetInnerHTML={{ __html: html }} />
}

const TypingDots = () => (
  <div style={{ display: 'flex', gap: 6, alignItems: 'center', padding: '8px 12px', background: 'var(--surface-2)', border: '1px solid var(--border)', borderRadius: '10px 10px 10px 2px' }}>
    {[0, 1, 2].map(i => (
      <div key={i} style={{ width: 6, height: 6, borderRadius: '50%', background: 'var(--text-muted)', animation: 'bounce 1s infinite', animationDelay: `${i * 0.15}s` }} />
    ))}
  </div>
)

const ChatComposer = memo(function ChatComposer({
  busy,
  currentFile,
  onSend
}: {
  busy: boolean
  currentFile?: any
  onSend: (text: string) => void
}) {
  const [value, setValue] = useState('')
  const inputRef = useRef<HTMLTextAreaElement>(null)

  const submit = () => {
    const text = value.trim()
    if (!text || busy) return
    onSend(text)
    setValue('')
    if (inputRef.current) {
      inputRef.current.value = ''
      inputRef.current.style.height = 'auto'
    }
  }

  return (
    <div style={{ borderTop: '1px solid var(--border)', padding: '8px 8px 6px', flexShrink: 0 }}>
      {!currentFile && (
        <div style={{ padding: '5px 8px', background: 'rgba(234,179,8,0.08)', border: '1px solid rgba(234,179,8,0.2)', borderRadius: 6, fontSize: '0.65rem', color: 'var(--yellow)', marginBottom: 6 }}>
          Open a file to ask data-specific questions
        </div>
      )}
      <div style={{ display: 'flex', gap: 5, background: 'var(--surface)', border: '1px solid var(--border)', borderRadius: 10, padding: '5px 7px' }}>
        <textarea
          ref={inputRef}
          value={value}
          onChange={e => setValue(e.target.value)}
          onKeyDown={e => {
            if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); submit() }
          }}
          placeholder={currentFile ? 'Ask anything... (Enter to send)' : 'Ask me to create data...'}
          disabled={busy}
          rows={1}
          style={{ flex: 1, border: 'none', background: 'transparent', outline: 'none', resize: 'none', fontSize: '0.76rem', color: 'var(--text)', lineHeight: 1.5, minHeight: 20, maxHeight: 80, fontFamily: 'var(--font-body)', caretColor: 'var(--blue)', opacity: busy ? 0.6 : 1 }}
          onInput={e => { const el = e.target as HTMLTextAreaElement; el.style.height = 'auto'; el.style.height = Math.min(el.scrollHeight, 80) + 'px' }}
        />
        <button
          onClick={submit}
          disabled={busy || !value.trim()}
          style={{ alignSelf: 'flex-end', width: 26, height: 26, borderRadius: 6, border: 'none', background: value.trim() && !busy ? 'var(--blue)' : 'var(--surface-2)', cursor: value.trim() && !busy ? 'pointer' : 'default', display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0, transition: 'background 0.15s' }}
        >
          <svg width={11} height={11} viewBox="0 0 24 24" fill="none" stroke={value.trim() && !busy ? '#fff' : 'var(--text-muted)'} strokeWidth={2.5}><line x1={22} y1={2} x2={11} y2={13}/><polygon points="22 2 15 22 11 13 2 9 22 2"/></svg>
        </button>
      </div>
      <div style={{ fontSize: '0.58rem', color: 'var(--text-muted)', textAlign: 'center', marginTop: 4 }}>â†µ send Â· â‡§â†µ newline</div>
    </div>
  )
})

export default function AIChatPanel({
  currentFile,
  rawRows,
  stats,
  offlineMode = false,
  panelWidth = 360,
  initialCollapsed = false,
  offlineCommands = [],
  onGridUpdate,
  onNewFile,
  onShowChart,
  onWorkspaceIntent,
  onOfflineCommand,
  onApplyPreview,
  onDiscardPreview,
  onUndoLastChange
}: Props) {
  const navigate = useNavigate()
  const user = useSelector((s: RootState) => s.app.user)
  const [collapsed, setCollapsed] = useState(initialCollapsed)
  const [tab, setTab] = useState<'chat' | 'generate' | 'formula' | 'settings'>('chat')

  const [msgs, setMsgs] = useState<Msg[]>([{
    id: 'offline_mode_notice',
    role: 'ai',
    time: now(),
    content: `Hi ${user?.name?.split(' ')[0] || 'there'}! I'm your assistant.\n\n${offlineMode ? 'Offline mode is active. Local workbook help is ready.' : 'AI mode is active. Ask for edits, summaries, formulas, or new workbook ideas.'}`
  }])
  const [inputDisplay, setInputDisplay] = useState('')
  const inputRef = useRef<HTMLTextAreaElement>(null)
  const [busy, setBusy] = useState(false)
  const msgsEndRef = useRef<HTMLDivElement>(null)

  const [formulaDesc, setFormulaDesc] = useState('')
  const [formulaRes, setFormulaRes] = useState<any>(null)
  const [formulaBusy, setFormulaBusy] = useState(false)

  const [genPrompt, setGenPrompt] = useState('')
  const [genBusy, setGenBusy] = useState(false)
  const [sheetPrompt, setSheetPrompt] = useState('')
  const [sheetBusy, setSheetBusy] = useState(false)

  const [geminiKey, setGeminiKey] = useState(() => localStorage.getItem('gemini_key') || '')
  const [hfKey, setHfKey] = useState(() => localStorage.getItem('hf_key') || '')
  const [geminiInput, setGeminiInput] = useState('')
  const [hfInput, setHfInput] = useState('')
  const [showGemini, setShowGemini] = useState(false)
  const [showHF, setShowHF] = useState(false)
  const [keyStatus, setKeyStatus] = useState({ gemini: !!localStorage.getItem('gemini_key'), hf: !!localStorage.getItem('hf_key') })
  const [saving, setSaving] = useState(false)
  const wasOfflineRef = useRef(offlineMode)
  const chatBodyRef = useRef<HTMLDivElement>(null)
  const chatScrollTopRef = useRef(0)
  const shouldStickToBottomRef = useRef(true)

  useEffect(() => {
    if (tab !== 'chat') return
    if (shouldStickToBottomRef.current) {
      msgsEndRef.current?.scrollIntoView({ behavior: 'smooth' })
      return
    }
    if (chatBodyRef.current) chatBodyRef.current.scrollTop = chatScrollTopRef.current
  }, [msgs.length, msgs[msgs.length - 1]?.content?.length, tab])

  useEffect(() => {
    if (tab === 'chat' && chatBodyRef.current) {
      chatBodyRef.current.scrollTop = chatScrollTopRef.current
    }
  }, [tab])

  useEffect(() => {
    if (offlineMode && (tab === 'generate' || tab === 'formula')) setTab('chat')
  }, [offlineMode, tab])

  useEffect(() => {
    const baseGreeting = offlineMode
      ? `Hi ${user?.name?.split(' ')[0] || 'there'}! I'm your assistant.\n\nOffline mode is active. Local workbook help is ready.`
      : `Hi ${user?.name?.split(' ')[0] || 'there'}! I'm your assistant.\n\nAI mode is active. Ask for edits, summaries, formulas, or new workbook ideas.`

    setMsgs(prev => {
      const rest = prev.filter(msg => msg.id !== 'offline_mode_notice' && !/I'm your AI Excel assistant|Open a file and ask me anything/i.test(msg.content))
      return [{ id: 'offline_mode_notice', role: 'ai', time: now(), content: baseGreeting }, ...rest]
    })
  }, [offlineMode, user?.name])

  useEffect(() => {
    if (offlineMode && !wasOfflineRef.current) {
      setMsgs(prev => {
        const keep = prev.filter(msg => {
          if (msg.id === 'offline_mode_notice') return false
          if (msg.role === 'user') return true
          if (msg.role === 'system') return /Offline mode keeps chat local/i.test(msg.content)
          return /Offline (analysis|chart|workbook) mode is active/i.test(msg.content)
        })
        return [{
          id: 'offline_mode_notice',
          role: 'ai',
          time: now(),
          content: `Hi ${user?.name?.split(' ')[0] || 'there'}! I'm your assistant.\n\nOffline mode is active. Local workbook help is ready.`
        }, ...keep]
      })
    }
    wasOfflineRef.current = offlineMode
  }, [offlineMode, user?.name])

  useEffect(() => {
    const g = localStorage.getItem('gemini_key')
    const h = localStorage.getItem('hf_key')
    if (g || h) {
      fetch('http://localhost:3001/api/llm/save-keys', {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ geminiKey: g || '', hfKey: h || '' })
      }).catch(() => {})
    }
  }, [])

  const addMsg = useCallback((role: Msg['role'], content: string, rows?: any[][]): string => {
    const id = uid()
    setMsgs(prev => [...prev, { id, role, content, rows, time: now() }])
    return id
  }, [])

  const updateMsg = useCallback((id: string, content: string, done = false) => {
    setMsgs(prev => prev.map(m => m.id === id ? { ...m, content, isStreaming: !done } : m))
  }, [])

  const saveKeys = async () => {
    setSaving(true)
    const newGemini = geminiInput.trim() || geminiKey
    const newHF = hfInput.trim() || hfKey
    try {
      await fetch('http://localhost:3001/api/llm/save-keys', {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ geminiKey: newGemini, hfKey: newHF })
      })
      if (geminiInput.trim()) { setGeminiKey(newGemini); localStorage.setItem('gemini_key', newGemini); setGeminiInput('') }
      if (hfInput.trim()) { setHfKey(newHF); localStorage.setItem('hf_key', newHF); setHfInput('') }
      setKeyStatus({ gemini: !!newGemini, hf: !!newHF })
      addMsg('system', `✓ Keys saved. ${newGemini ? 'Gemini (primary)' : ''}${newHF ? ' + HuggingFace (fallback)' : ''} active.`)
      setTab('chat')
    } catch { addMsg('system', '✗ Failed to save keys') }
    setSaving(false)
  }

  const exportRows = async (rows: any[][]) => {
    if (!rows?.length) return
    try {
      const headers = rows[0]
      const data = rows.slice(1).map(row => Object.fromEntries(row.map((v, i) => [String(headers[i] || `col${i}`), v])))
      const res = await fetch('http://localhost:3001/api/excel/generate', {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ data, filename: 'ai_result.xlsx', department: user?.department || 'general' })
      })
      const j = await res.json()
      window.open(`http://localhost:3001${j.file.url}`, '_blank')
    } catch { addMsg('system', '✗ Export failed') }
  }

  const streamChat = useCallback(async (text: string) => {
    const aiId = uid()
    setMsgs(prev => [...prev, { id: aiId, role: 'ai', content: '', isStreaming: true, time: now() }])
    try {
      const res = await fetch(`${BASE}/stream`, {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ message: text, rawRows, stats })
      })
      if (!res.ok || !res.body) { updateMsg(aiId, `⚠ Stream failed`, true); return }
      const reader = res.body.getReader()
      const decoder = new TextDecoder()
      let full = '', buf = ''
      while (true) {
        const { done, value } = await reader.read()
        if (done) break
        buf += decoder.decode(value, { stream: true })
        const lines = buf.split('\n'); buf = lines.pop() || ''
        for (const line of lines) {
          if (!line.startsWith('data: ')) continue
          try {
            const d = JSON.parse(line.slice(6))
            if (d.error) { updateMsg(aiId, `⚠ ${d.error}`, true); return }
            if (d.text) { full += d.text; updateMsg(aiId, full, false) }
            if (d.done) updateMsg(aiId, full, true)
          } catch {}
        }
      }
      if (full) updateMsg(aiId, full, true)
    } catch (e: any) { updateMsg(aiId, `⚠ ${e.message}`, true) }
  }, [rawRows, stats, updateMsg])

  const send = useCallback(async (msg?: string) => {
    const text = (msg || '').trim()
    if (!text || busy) return
    setBusy(true)
    addMsg('user', text)
    const lowerText = text.toLowerCase()
    if (/\b(apply|confirm)\b.*\b(preview|changes?)\b/i.test(text) && onApplyPreview) {
      onApplyPreview()
      addMsg('system', 'Applied the current preview to the sheet.')
      setBusy(false)
      return
    }
    if (/\b(discard|cancel)\b.*\b(preview|changes?)\b/i.test(text) && onDiscardPreview) {
      onDiscardPreview()
      addMsg('system', 'Discarded the current preview.')
      setBusy(false)
      return
    }
    if (/\bundo\b.*\b(change|preview|last)\b/i.test(text) && onUndoLastChange) {
      onUndoLastChange()
      addMsg('system', 'Undid the last AI-applied change.')
      setBusy(false)
      return
    }
    const localLookup = applyLocalLookupCommand(text, rawRows)
    if (localLookup) {
      addMsg('ai', localLookup.message, localLookup.rows)
      setBusy(false)
      return
    }
    const localSheetCommand = applyLocalSheetCommand(text, rawRows)
    if (localSheetCommand) {
      if ('error' in localSheetCommand) addMsg('system', localSheetCommand.error)
      else if (onGridUpdate) {
        const msgId = addMsg('ai', localSheetCommand.message, localSheetCommand.rows)
        setMsgs(prev => prev.map(m => m.id === msgId ? { ...m, rows: localSheetCommand.rows } : m))
        onGridUpdate(localSheetCommand.rows, localSheetCommand.message)
      }
      setBusy(false)
      return
    }
    if (offlineMode) {
      if (/\b(analy[sz]e|analysis|insight|summary|summarize|explain)\b/i.test(text) && onWorkspaceIntent) {
        onWorkspaceIntent({
          tab: 'analyze',
          prompt: text,
          questions: ['analyze data', 'summary statistics', 'show null values', 'quick audit']
        })
        onOfflineCommand?.(text)
        addMsg('ai', 'Offline analysis is active. I opened Analysis and sent this command to the local workbook tools.')
        setBusy(false)
        return
      }
      if (/\b(chart|graph|plot|visuali[sz]e)\b/i.test(text) && onShowChart) {
        const offlineChartType = lowerText.includes('line') ? 'line' : lowerText.includes('pie') || lowerText.includes('donut') ? 'pie' : 'bar'
        const offlineChartKey =
          lowerText.includes('subject') ? 'subject' :
          lowerText.includes('grade') ? 'grade' :
          lowerText.includes('score') || lowerText.includes('distribution') ? 'score' :
          'section'
        onShowChart({ chartType: offlineChartType, chartDataKey: offlineChartKey, title: `Offline chart: ${text}` })
        onWorkspaceIntent?.({
          tab: 'charts',
          prompt: text,
          questions: ['create chart', 'bar chart', 'line chart', 'chart top categories']
        })
        onOfflineCommand?.(text)
        addMsg('ai', 'Offline chart mode is active. I opened Charts and sent this command to the local workbook tools.')
        setBusy(false)
        return
      }
      if (/\b(create|generate|make|clean|report|template|split|segment|export)\b/i.test(text)) {
        onWorkspaceIntent?.({
          tab: 'create',
          prompt: text,
          questions: ['clean data', 'generate report', 'advanced excel', 'template']
        })
        onOfflineCommand?.(text)
        addMsg('ai', 'Offline workbook mode is active. I sent this command to the local workbook tools.')
        setBusy(false)
        return
      }
      const examples = offlineCommands.slice(0, 8).map(item => `- \`${item.command}\`: ${item.effect}`).join('\n')
      addMsg('system', `Offline mode keeps chat local, so cloud AI replies are paused.\n\nYou can still use row edits like \`add Krishna at last line\`, chart requests like \`create chart\`, and workbook actions such as:\n${examples}`)
      setBusy(false)
      return
    }
    const contextPayload = {
      prompt: text,
      fileName: currentFile?.name || 'Current sheet',
      data: buildSheetContext(rawRows, stats),
      user: { name: user?.name, department: user?.department }
    }
    if (/\b(create|generate|make)\b.*\b(doc|document|report|notice|proposal)\b/i.test(text)) {
      sessionStorage.setItem('ai_document_draft', JSON.stringify(contextPayload))
      navigate('/documents', { state: { aiDraft: contextPayload } })
      addMsg('ai', 'Moved this request to Documents and prepared an AI draft from the current sheet.')
      setBusy(false)
      return
    }
    if (/\b(create|generate|make)\b.*\b(ppt|powerpoint|presentation|deck|slides)\b/i.test(text)) {
      const pptQuestions = [
        'Who is the audience for this deck?',
        'What is the main goal of the presentation?',
        'How many slides do you want roughly?',
        'Should the tone be academic, executive, or investor-focused?'
      ]
      sessionStorage.setItem('ai_ppt_draft', JSON.stringify(contextPayload))
      sessionStorage.setItem('ai_ppt_questions', JSON.stringify(pptQuestions))
      navigate('/powerpoint', { state: { aiDraft: contextPayload } })
      addMsg('ai', `Moved this request to PPT Generator with your current sheet context.\n\nAnswer these first:\n- ${pptQuestions.join('\n- ')}`)
      setBusy(false)
      return
    }
    if (/\b(analy[sz]e|analysis|insight|summary|summarize|explain)\b/i.test(text) && onWorkspaceIntent) {
      onWorkspaceIntent({
        tab: 'analyze',
        prompt: text,
        questions: [
          'Which metric matters most here?',
          'Show top risks from this sheet',
          'Compare sections or categories',
          'What should I visualize next?'
        ]
      })
      addMsg('ai', 'Opened the Analysis workspace and loaded follow-up questions based on your request.')
      setBusy(false)
      return
    }
    if (/\b(create|generate|make)\b.*\b(sheet|table|excel|spreadsheet)\b/i.test(text) && onWorkspaceIntent) {
      onWorkspaceIntent({
        tab: 'create',
        prompt: text,
        questions: [
          'Should the sheet be a summary or a raw table?',
          'Do you want formulas included?',
          'Should I color-code or rank the rows?',
          'Which columns are required?'
        ]
      })
      setTab('generate')
      addMsg('ai', 'Opened the Create workspace and prepared your request so you can guide the next sheet with AI.')
      setBusy(false)
      return
    }
    const chartMatch = text.match(/\b(bar|graph|chart|visuali[sz]e|plot)\b/i)
    if (chartMatch && stats) {
      const chartType = lowerText.includes('line') ? 'line' : lowerText.includes('pie') || lowerText.includes('donut') ? 'pie' : 'bar'
      const chartDataKey =
        lowerText.includes('subject') ? 'subject' :
        lowerText.includes('grade') ? 'grade' :
        lowerText.includes('score') || lowerText.includes('distribution') ? 'score' :
        'section'
      if (onShowChart) {
        onShowChart({ chartType, chartDataKey, title: `AI chart: ${text}` })
        onWorkspaceIntent?.({
          tab: 'charts',
          prompt: text,
          questions: [
            'Which field should be on the X-axis?',
            'Do you want bar, line, or pie?',
            'Should the chart compare sections, grades, or scores?',
            'Do you want analysis text with the chart?'
          ]
        })
        addMsg('ai', `Prepared a ${chartType} chart preview for ${chartDataKey} data. Review it in Excel, then apply or discard it.`)
      } else {
        addMsg('ai', 'Chart view is not connected yet.')
      }
      setBusy(false)
      return
    }
    const isCmd = rawRows && rawRows.length > 1 &&
      /top\s*\d*|fail|pass|section|attend|below|above|sort|filter|search|add.*col|grade|subject|rank|list|show|find|who|which|compare|avg|average|highest|lowest/i.test(text)
    if (isCmd) {
      try {
        const res = await fetch(`${BASE}/command`, {
          method: 'POST', headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ command: text, rawRows, stats })
        })
        const result = await res.json()
        if (!res.ok) throw new Error(result.error)
        if (result.action === 'summary') {
          addMsg('ai', asText(result.answer || result.description || 'Done.'))
        } else if (result.rows?.length > 1) {
          const summary = [asText(result.description), asText(result.answer)].filter(Boolean).join('\n\n')
          const msgId = addMsg('ai', summary || `**Result: ${result.rows.length - 1} rows**`)
          setMsgs(prev => prev.map(m => m.id === msgId ? { ...m, rows: result.rows } : m))
          if (result.action === 'add_column' && onGridUpdate) {
            onGridUpdate(result.rows, asText(result.description || 'Column added'))
            addMsg('system', `✓ Preview ready — column "${result.column_name}" can be reviewed before applying`)
          }
        } else { await streamChat(text) }
      } catch { await streamChat(text) }
    } else { await streamChat(text) }
    setBusy(false)
  }, [busy, rawRows, stats, offlineMode, offlineCommands, onGridUpdate, onShowChart, onWorkspaceIntent, onOfflineCommand, onApplyPreview, onDiscardPreview, onUndoLastChange, addMsg, streamChat])

  const handleKeyDown = useCallback((e: React.KeyboardEvent<HTMLTextAreaElement>) => {
    if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); send(inputDisplay) }
  }, [send, inputDisplay])

  const generateFormula = async () => {
    if (!formulaDesc.trim()) return
    setFormulaBusy(true); setFormulaRes(null)
    try {
      const headers = rawRows?.[0]?.map((h: any) => String(h)) || []
      const res = await fetch(`${BASE}/formula`, {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ description: formulaDesc, headers })
      })
      setFormulaRes(await res.json())
    } catch (e: any) { setFormulaRes({ error: e.message }) }
    setFormulaBusy(false)
  }

  const generateExcel = async () => {
    if (!genPrompt.trim()) return
    setGenBusy(true)
    try {
      const res = await fetch(`${BASE}/generate-excel`, {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ prompt: genPrompt, department: user?.department || 'general' })
      })
      const result = await res.json()
      if (!res.ok) throw new Error(result.error)
      if (onNewFile && result.file) onNewFile(result.file, result.rows)
      setGenPrompt(''); setTab('chat')
      const msgId = addMsg('ai', `✓ **Created "${result.sheet_name}"** — ${result.rowCount} rows\n\n${result.description}\n\nThe workbook is now linked back to Excel Sheet so you can open, edit, and save it.`)
      if (result.rows?.length > 1) setMsgs(prev => prev.map(m => m.id === msgId ? { ...m, rows: result.rows } : m))
    } catch (e: any) { addMsg('system', `✗ ${e.message}`); setTab('chat') }
    setGenBusy(false)
  }

  const generateSheet = async () => {
    if (!sheetPrompt.trim() || !rawRows?.length) return
    setSheetBusy(true)
    try {
      const res = await fetch(`${BASE}/new-sheet`, {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ instruction: sheetPrompt, rawRows, stats, department: user?.department })
      })
      const result = await res.json()
      if (!res.ok) throw new Error(result.error)
      if (onNewFile && result.file) onNewFile(result.file, result.rows)
      setSheetPrompt(''); setTab('chat')
      const msgId = addMsg('ai', `✓ **"${result.sheet_name}"** — ${result.rowCount} rows\n\n${result.description}\n\nIt has been sent back to Excel Sheet so you can review and keep editing manually.`)
      if (result.rows?.length > 1) setMsgs(prev => prev.map(m => m.id === msgId ? { ...m, rows: result.rows } : m))
    } catch (e: any) { addMsg('system', `✗ ${e.message}`); setTab('chat') }
    setSheetBusy(false)
  }

  const anyKeySet = offlineMode || keyStatus.gemini || keyStatus.hf
  const hasHiddenStreamingAi = msgs.some(m => m.role === 'ai' && m.isStreaming && !m.content.trim())
  const hasVisibleStreamingAi = msgs.some(m => m.role === 'ai' && m.isStreaming && !!m.content.trim())

  if (collapsed) {
    return (
      <div onClick={() => setCollapsed(false)} style={{ width: 36, background: 'var(--bg)', borderLeft: '1px solid var(--border)', display: 'flex', flexDirection: 'column', alignItems: 'center', paddingTop: 14, gap: 8, cursor: 'pointer', flexShrink: 0 }}>
        <span style={{ fontSize: '0.58rem', color: 'var(--text-muted)', writingMode: 'vertical-lr' as any, textTransform: 'uppercase', letterSpacing: '0.06em' }}>{offlineMode ? 'Offline' : 'AI Panel'}</span>
        <div style={{ width: 6, height: 6, borderRadius: '50%', background: offlineMode ? 'var(--warning)' : anyKeySet ? 'var(--green)' : 'var(--orange)' }} />
      </div>
    )
  }

  return (
    <aside style={{ width: panelWidth, minWidth: 280, maxWidth: 620, flexShrink: 0, background: 'var(--bg)', borderLeft: '1px solid var(--border)', display: 'flex', flexDirection: 'column' }}>
      <div style={{ height: 44, display: 'flex', alignItems: 'center', padding: '0 10px', borderBottom: '1px solid var(--border)', gap: 7, flexShrink: 0 }}>
        <div style={{ width: 24, height: 24, background: offlineMode ? 'linear-gradient(135deg, #f59e0b, #b45309)' : 'var(--blue)', borderRadius: 5, display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: '0.6rem', fontWeight: 800, color: '#fff', flexShrink: 0 }}>
          {offlineMode ? 'OF' : 'AI'}
        </div>
        <div style={{ flex: 1, minWidth: 0 }}>
          <div style={{ fontSize: '0.75rem', fontWeight: 600, color: 'var(--text)' }}>{offlineMode ? 'Xtron Offline' : 'Xtron'}</div>
          <div style={{ fontSize: '0.6rem', display: 'flex', alignItems: 'center', gap: 3, color: offlineMode ? 'var(--orange)' : anyKeySet ? 'var(--green)' : 'var(--orange)' }}>
            <div style={{ width: 4, height: 4, borderRadius: '50%', background: offlineMode ? 'var(--orange)' : anyKeySet ? 'var(--green)' : 'var(--orange)', flexShrink: 0 }} />
            {offlineMode ? 'Offline mode active' : anyKeySet ? `${keyStatus.gemini ? 'Gemini' : ''}${keyStatus.gemini && keyStatus.hf ? ' + ' : ''}${keyStatus.hf ? 'HuggingFace' : ''} · Ready` : 'AI mode · Add API key in Settings'}
          </div>
        </div>
        {currentFile && <div style={{ fontSize: '0.58rem', color: 'var(--text-muted)', maxWidth: 60, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>📄 {currentFile.name}</div>}
        <button onClick={() => setCollapsed(true)} style={{ background: 'none', border: 'none', color: 'var(--text-muted)', cursor: 'pointer', fontSize: 18, lineHeight: 1, padding: '0 2px', flexShrink: 0 }}>›</button>
      </div>

      <div style={{ display: 'flex', borderBottom: '1px solid var(--border)', flexShrink: 0 }}>
        {([['chat','Chat'], ...(!offlineMode ? [['generate','Generate'],['formula','Formula']] : []), ['settings','⚙']] as Array<[string, string]>).map(([id, label]) => (
          <button key={id} onClick={() => setTab(id as any)}
            style={{ flex: 1, padding: '6px 0', background: 'none', border: 'none', borderBottom: tab === id ? '2px solid var(--blue)' : '2px solid transparent', cursor: 'pointer', fontSize: '0.7rem', fontWeight: tab === id ? 600 : 400, color: tab === id ? 'var(--text)' : 'var(--text-muted)' }}>
            {label}
          </button>
        ))}
      </div>

      {tab === 'chat' && (
        <>
          {!offlineMode && !anyKeySet && (
            <div style={{ padding: '8px 10px', background: 'rgba(234,179,8,0.08)', borderBottom: '1px solid rgba(234,179,8,0.2)', fontSize: '0.65rem', color: 'var(--orange)', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
              <span>⚠ No API key set</span>
              <button onClick={() => setTab('settings')} style={{ background: 'none', border: '1px solid var(--orange)', borderRadius: 4, padding: '2px 8px', fontSize: '0.62rem', color: 'var(--orange)', cursor: 'pointer' }}>Add Key</button>
            </div>
          )}
          {rawRows && rawRows.length > 1 && !offlineMode && (
            <div style={{ display: 'flex', gap: 4, padding: '6px 8px', borderBottom: '1px solid var(--border)', overflowX: 'auto', flexShrink: 0 }}>
              {QUICK.map(q => (
                <button key={q} onClick={() => !busy && send(q)} disabled={busy}
                  style={{ padding: '3px 8px', background: 'var(--surface)', border: '1px solid var(--border)', borderRadius: 10, fontSize: '0.6rem', color: 'var(--text-secondary)', cursor: busy ? 'default' : 'pointer', whiteSpace: 'nowrap', flexShrink: 0, opacity: busy ? 0.5 : 1 }}>
                  {q}
                </button>
              ))}
            </div>
          )}
          <div
            ref={chatBodyRef}
            onScroll={() => {
              if (!chatBodyRef.current) return
              chatScrollTopRef.current = chatBodyRef.current.scrollTop
              const { scrollTop, scrollHeight, clientHeight } = chatBodyRef.current
              shouldStickToBottomRef.current = scrollHeight - (scrollTop + clientHeight) < 48
            }}
            style={{ flex: 1, overflowY: 'auto', padding: '10px 8px', display: 'flex', flexDirection: 'column', gap: 10 }}
          >
            {msgs.filter(m => !(m.role === 'ai' && m.isStreaming && !m.content.trim())).map(m => (
              <div key={m.id} style={{ display: 'flex', gap: 6, flexDirection: m.role === 'user' ? 'row-reverse' : 'row', alignItems: 'flex-start' }}>
                {m.role !== 'user' && (
                  <div style={{ width: 22, height: 22, borderRadius: 5, background: m.role === 'ai' ? 'var(--blue)' : 'var(--surface-2)', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: '0.55rem', fontWeight: 800, color: m.role === 'ai' ? '#fff' : 'var(--text-muted)', flexShrink: 0, marginTop: 2 }}>
                    {m.role === 'ai' ? 'AI' : '⚙'}
                  </div>
                )}
                <div style={{ maxWidth: '88%', display: 'flex', flexDirection: 'column', gap: 3, alignItems: m.role === 'user' ? 'flex-end' : 'flex-start' }}>
                  <div style={{ padding: '7px 10px', background: m.role === 'user' ? 'var(--blue)' : 'var(--surface-2)', color: m.role === 'user' ? '#fff' : 'var(--text)', borderRadius: m.role === 'user' ? '10px 10px 2px 10px' : '10px 10px 10px 2px', border: '1px solid var(--border)', fontSize: '0.76rem', lineHeight: 1.55, wordBreak: 'break-word' }}>
                    {m.role === 'user' ? m.content : <MD text={m.content} />}
                    {m.isStreaming && <span style={{ display: 'inline-block', width: 7, height: 13, background: 'var(--accent)', animation: 'blink 0.7s infinite', marginLeft: 2, verticalAlign: 'text-bottom', borderRadius: 1 }} />}
                  </div>
                  {m.rows && m.rows.length > 1 && <MiniTable rows={m.rows} onExport={() => exportRows(m.rows!)} onView={onGridUpdate ? () => onGridUpdate(m.rows as string[][], 'Viewed AI result') : undefined} />}
                  <div style={{ fontSize: '0.58rem', color: 'var(--text-muted)' }}>{m.time}</div>
                </div>
              </div>
            ))}
            {busy && !hasVisibleStreamingAi && (hasHiddenStreamingAi || msgs[msgs.length-1]?.role === 'user') && (
              <div style={{ display: 'flex', gap: 6 }}>
                <div style={{ width: 22, height: 22, borderRadius: 5, background: 'var(--blue)', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: '0.55rem', fontWeight: 800, color: '#fff' }}>AI</div>
                <TypingDots />
              </div>
            )}
            <div ref={msgsEndRef} />
          </div>
          <div style={{ display: 'none' }}>
            {!currentFile && (
              <div style={{ padding: '5px 8px', background: 'rgba(234,179,8,0.08)', border: '1px solid rgba(234,179,8,0.2)', borderRadius: 6, fontSize: '0.65rem', color: 'var(--yellow)', marginBottom: 6 }}>
                Open a file to ask data-specific questions
              </div>
            )}
            <div style={{ display: 'flex', gap: 5, background: 'var(--surface)', border: '1px solid var(--border)', borderRadius: 10, padding: '5px 7px' }}>
              <textarea ref={inputRef} onKeyDown={handleKeyDown} onChange={e => setInputDisplay(e.target.value)}
                placeholder={currentFile ? 'Ask anything... (Enter to send)' : 'Ask me to create data...'}
                disabled={busy} rows={1}
                style={{ flex: 1, border: 'none', background: 'transparent', outline: 'none', resize: 'none', fontSize: '0.76rem', color: 'var(--text)', lineHeight: 1.5, minHeight: 20, maxHeight: 80, fontFamily: 'var(--font-body)', caretColor: 'var(--blue)', opacity: busy ? 0.6 : 1 }}
                onInput={e => { const el = e.target as HTMLTextAreaElement; el.style.height = 'auto'; el.style.height = Math.min(el.scrollHeight, 80) + 'px' }}
              />
              <button onClick={() => send()} disabled={busy || !inputDisplay.trim()}
                style={{ alignSelf: 'flex-end', width: 26, height: 26, borderRadius: 6, border: 'none', background: inputDisplay.trim() && !busy ? 'var(--blue)' : 'var(--surface-2)', cursor: inputDisplay.trim() && !busy ? 'pointer' : 'default', display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0, transition: 'background 0.15s' }}>
                <svg width={11} height={11} viewBox="0 0 24 24" fill="none" stroke={inputDisplay.trim() && !busy ? '#fff' : 'var(--text-muted)'} strokeWidth={2.5}><line x1={22} y1={2} x2={11} y2={13}/><polygon points="22 2 15 22 11 13 2 9 22 2"/></svg>
              </button>
            </div>
            <div style={{ fontSize: '0.58rem', color: 'var(--text-muted)', textAlign: 'center', marginTop: 4 }}>↵ send · ⇧↵ newline</div>
          </div>
          <ChatComposer busy={busy} currentFile={currentFile} onSend={send} />
        </>
      )}

      {!offlineMode && tab === 'generate' && (
        <div style={{ flex: 1, overflowY: 'auto', padding: 12, display: 'flex', flexDirection: 'column', gap: 12 }}>
          {offlineMode && (
            <div style={{ background: 'rgba(245,158,11,0.08)', border: '1px solid rgba(245,158,11,0.2)', borderRadius: 10, padding: 12, fontSize: '0.7rem', color: 'var(--text-secondary)', lineHeight: 1.6 }}>
              Offline mode keeps this panel focused on local workbook commands. Use the Chat tab to run commands like `analyze data`, `create chart`, `clean data`, or `advanced excel`.
            </div>
          )}
          <div style={{ background: 'var(--surface)', border: '1px solid var(--border)', borderRadius: 10, padding: 12 }}>
            <div style={{ fontSize: '0.65rem', color: 'var(--text-muted)', textTransform: 'uppercase', marginBottom: 7 }}>✨ Create Excel from scratch</div>
            <textarea className="input" value={genPrompt} onChange={e => setGenPrompt(e.target.value)} placeholder="e.g., 'Student grade sheet for 30 students with Maths, Physics, Chemistry'" style={{ width: '100%', minHeight: 70, fontSize: '0.76rem', resize: 'vertical', marginBottom: 7 }} />
            <button className="btn btn-primary btn-sm" onClick={generateExcel} disabled={offlineMode || genBusy || !genPrompt.trim()} style={{ width: '100%', justifyContent: 'center' }}>
              {genBusy ? '⟳ Generating...' : '✨ Generate File'}
            </button>
            <div style={{ display: 'flex', flexWrap: 'wrap', gap: 4, marginTop: 7 }}>
              {['Student grade sheet 30 students', 'Attendance tracker semester', 'Monthly sales report'].map(ex => (
                <button key={ex} className="btn btn-xs btn-outline" onClick={() => setGenPrompt(ex)} style={{ fontSize: '0.6rem' }}>{ex}</button>
              ))}
            </div>
          </div>
          <div style={{ background: 'var(--surface)', border: '1px solid var(--border)', borderRadius: 10, padding: 12, opacity: !currentFile ? 0.55 : 1 }}>
            <div style={{ fontSize: '0.65rem', color: 'var(--text-muted)', textTransform: 'uppercase', marginBottom: 5 }}>⊞ New sheet from current data</div>
            {!currentFile && <div style={{ fontSize: '0.65rem', color: 'var(--orange)', marginBottom: 5 }}>Open a file first</div>}
            <textarea className="input" value={sheetPrompt} onChange={e => setSheetPrompt(e.target.value)} disabled={!currentFile} placeholder={'e.g., "Top 10 students only"\n"Failed students"\n"Section summary"'} style={{ width: '100%', minHeight: 70, fontSize: '0.76rem', resize: 'vertical', marginBottom: 7 }} />
            <button className="btn btn-outline btn-sm" onClick={generateSheet} disabled={offlineMode || sheetBusy || !sheetPrompt.trim() || !currentFile} style={{ width: '100%', justifyContent: 'center' }}>
              {sheetBusy ? '⟳ Creating...' : '⊞ Create New Sheet'}
            </button>
          </div>
        </div>
      )}

      {!offlineMode && tab === 'formula' && (
        <div style={{ flex: 1, overflowY: 'auto', padding: 12, display: 'flex', flexDirection: 'column', gap: 10 }}>
          {offlineMode && (
            <div style={{ background: 'rgba(245,158,11,0.08)', border: '1px solid rgba(245,158,11,0.2)', borderRadius: 10, padding: 12, fontSize: '0.7rem', color: 'var(--text-secondary)', lineHeight: 1.6 }}>
              Formula AI is paused offline. You can still edit cells manually in the sheet and use chat for local row, chart, and analysis commands.
            </div>
          )}
          <div style={{ fontSize: '0.65rem', color: 'var(--text-muted)', textTransform: 'uppercase', marginBottom: 3 }}>Describe what you want to calculate</div>
          <textarea className="input" value={formulaDesc} onChange={e => setFormulaDesc(e.target.value)} onKeyDown={e => e.key === 'Enter' && e.ctrlKey && generateFormula()} placeholder={'e.g., "Average of all mark columns"\n"Count students who passed"'} style={{ width: '100%', minHeight: 80, fontSize: '0.78rem', resize: 'vertical' }} />
          {rawRows?.[0] && <div style={{ fontSize: '0.6rem', color: 'var(--text-muted)' }}>Columns: {rawRows[0].map((h: any) => String(h)).filter(Boolean).join(', ')}</div>}
          <button className="btn btn-primary btn-sm" onClick={generateFormula} disabled={offlineMode || formulaBusy || !formulaDesc.trim()} style={{ justifyContent: 'center' }}>
            {formulaBusy ? '⟳ Generating...' : 'Generate Formula'}
          </button>
          {formulaRes && !formulaRes.error && (
            <div style={{ background: 'var(--surface)', border: '1px solid var(--border)', borderRadius: 10, padding: 12, display: 'flex', flexDirection: 'column', gap: 8 }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', background: 'var(--surface-2)', border: '1px solid var(--border)', borderRadius: 6, padding: '6px 10px' }}>
                <code style={{ fontFamily: 'var(--font-mono)', fontSize: '0.78rem', color: 'var(--green)', wordBreak: 'break-all' }}>{formulaRes.formula}</code>
                <button onClick={() => navigator.clipboard.writeText(formulaRes.formula)} style={{ background: 'none', border: '1px solid var(--border)', borderRadius: 4, padding: '2px 6px', fontSize: '0.6rem', color: 'var(--text-muted)', cursor: 'pointer', flexShrink: 0, marginLeft: 8 }}>Copy</button>
              </div>
              <p style={{ fontSize: '0.74rem', color: 'var(--text-secondary)', lineHeight: 1.6, margin: 0 }}>{formulaRes.explanation}</p>
              {formulaRes.example && <div style={{ fontSize: '0.68rem', color: 'var(--text-muted)', fontStyle: 'italic', background: 'var(--surface-2)', borderRadius: 6, padding: '5px 8px' }}>{formulaRes.example}</div>}
            </div>
          )}
          {formulaRes?.error && <div style={{ padding: '8px', background: 'rgba(239,68,68,0.1)', border: '1px solid rgba(239,68,68,0.3)', borderRadius: 7, fontSize: '0.72rem', color: 'var(--red)' }}>{formulaRes.error}</div>}
          <div style={{ borderTop: '1px solid var(--border)', paddingTop: 10 }}>
            <div style={{ fontSize: '0.6rem', color: 'var(--text-muted)', textTransform: 'uppercase', marginBottom: 6 }}>Quick formulas</div>
            {['Average of all mark columns', 'Count passed students (≥40)', 'Rank by total score', 'Grade: O/A+/A/B+/B/C/F', 'Sum if attendance > 75'].map(ex => (
              <button key={ex} onClick={() => setFormulaDesc(ex)} style={{ display: 'block', width: '100%', textAlign: 'left', padding: '5px 8px', marginBottom: 3, background: 'var(--surface-2)', border: '1px solid var(--border)', borderRadius: 6, cursor: 'pointer', fontSize: '0.7rem', color: 'var(--text-secondary)' }}>{ex}</button>
            ))}
          </div>
        </div>
      )}

      {tab === 'settings' && (
        <div style={{ flex: 1, overflowY: 'auto', padding: 12, display: 'flex', flexDirection: 'column', gap: 12 }}>
          <div style={{ display: 'flex', gap: 6 }}>
            {[['Gemini', 'Primary · Free · 60 req/min', keyStatus.gemini], ['HuggingFace', 'Fallback · Free · No quota', keyStatus.hf]].map(([name, sub, active]) => (
              <div key={name as string} style={{ flex: 1, padding: '6px 10px', borderRadius: 8, background: active ? 'rgba(34,197,94,0.08)' : 'rgba(239,68,68,0.06)', border: `1px solid ${active ? 'rgba(34,197,94,0.25)' : 'rgba(239,68,68,0.2)'}`, textAlign: 'center' }}>
                <div style={{ fontSize: '0.6rem', color: 'var(--text-muted)', marginBottom: 2 }}>{name as string}</div>
                <div style={{ fontSize: '0.72rem', fontWeight: 600, color: active ? 'var(--green)' : 'var(--red)' }}>{active ? '✓ Active' : '✗ Not set'}</div>
              </div>
            ))}
          </div>

          {keyStatus.gemini && keyStatus.hf && (
            <div style={{ padding: '6px 10px', background: 'rgba(59,130,246,0.08)', border: '1px solid rgba(59,130,246,0.2)', borderRadius: 8, fontSize: '0.68rem', color: 'var(--blue)' }}>
              ⚡ Auto-fallback on — Gemini quota hit? HuggingFace kicks in automatically
            </div>
          )}

          {/* Gemini */}
          <div style={{ background: 'var(--surface)', border: '1px solid var(--border)', borderRadius: 10, padding: 12 }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 8 }}>
              <div>
                <div style={{ fontSize: '0.72rem', fontWeight: 600, color: 'var(--text)' }}>Gemini API Key</div>
                <div style={{ fontSize: '0.6rem', color: 'var(--text-muted)' }}>Primary · Free · 60 req/min</div>
              </div>
              {keyStatus.gemini && <button onClick={() => { setGeminiKey(''); localStorage.removeItem('gemini_key'); setKeyStatus(s => ({ ...s, gemini: false })) }} style={{ background: 'none', border: 'none', color: 'var(--red)', cursor: 'pointer', fontSize: '0.62rem' }}>Remove</button>}
            </div>
            <div style={{ position: 'relative' }}>
              <input type={showGemini ? 'text' : 'password'} value={geminiInput} onChange={e => setGeminiInput(e.target.value)}
                placeholder={keyStatus.gemini ? '••••••••••• (saved)' : 'AIza...'}
                style={{ width: '100%', padding: '7px 44px 7px 10px', background: 'var(--surface-2)', border: '1px solid var(--border)', borderRadius: 6, color: 'var(--text)', fontSize: '0.76rem', boxSizing: 'border-box' as any }} />
              <button onClick={() => setShowGemini(p => !p)} style={{ position: 'absolute', right: 8, top: '50%', transform: 'translateY(-50%)', background: 'none', border: 'none', color: 'var(--text-muted)', cursor: 'pointer', fontSize: '0.58rem', fontWeight: 700 }}>{showGemini ? 'HIDE' : 'SHOW'}</button>
            </div>
            <div style={{ fontSize: '0.62rem', color: 'var(--text-muted)', marginTop: 5 }}>Get free at <strong style={{ color: 'var(--blue)' }}>aistudio.google.com</strong> → Get API key</div>
          </div>

          {/* HuggingFace */}
          <div style={{ background: 'var(--surface)', border: '1px solid var(--border)', borderRadius: 10, padding: 12 }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 8 }}>
              <div>
                <div style={{ fontSize: '0.72rem', fontWeight: 600, color: 'var(--text)' }}>HuggingFace Token</div>
                <div style={{ fontSize: '0.6rem', color: 'var(--text-muted)' }}>Fallback · Free · No quota</div>
              </div>
              {keyStatus.hf && <button onClick={() => { setHfKey(''); localStorage.removeItem('hf_key'); setKeyStatus(s => ({ ...s, hf: false })) }} style={{ background: 'none', border: 'none', color: 'var(--red)', cursor: 'pointer', fontSize: '0.62rem' }}>Remove</button>}
            </div>
            <div style={{ position: 'relative' }}>
              <input type={showHF ? 'text' : 'password'} value={hfInput} onChange={e => setHfInput(e.target.value)}
                placeholder={keyStatus.hf ? '••••••••••• (saved)' : 'hf_...'}
                style={{ width: '100%', padding: '7px 44px 7px 10px', background: 'var(--surface-2)', border: '1px solid var(--border)', borderRadius: 6, color: 'var(--text)', fontSize: '0.76rem', boxSizing: 'border-box' as any }} />
              <button onClick={() => setShowHF(p => !p)} style={{ position: 'absolute', right: 8, top: '50%', transform: 'translateY(-50%)', background: 'none', border: 'none', color: 'var(--text-muted)', cursor: 'pointer', fontSize: '0.58rem', fontWeight: 700 }}>{showHF ? 'HIDE' : 'SHOW'}</button>
            </div>
            <div style={{ fontSize: '0.62rem', color: 'var(--text-muted)', marginTop: 5 }}>Get free at <strong style={{ color: 'var(--blue)' }}>huggingface.co/settings/tokens</strong> → New token → Read</div>
          </div>

          {(geminiInput.trim() || hfInput.trim()) && (
            <button onClick={saveKeys} disabled={saving} style={{ padding: '10px', background: 'var(--blue)', color: '#fff', border: 'none', borderRadius: 8, cursor: 'pointer', fontSize: '0.8rem', fontWeight: 600 }}>
              {saving ? '⟳ Saving...' : '💾 Save Keys'}
            </button>
          )}

          <div style={{ background: 'var(--surface)', border: '1px solid var(--border)', borderRadius: 10, padding: 12 }}>
            <div style={{ fontSize: '0.65rem', color: 'var(--text-muted)', textTransform: 'uppercase', marginBottom: 8 }}>How fallback works</div>
            <div style={{ fontSize: '0.7rem', color: 'var(--text-secondary)', lineHeight: 1.8 }}>
              <div>1️⃣ Every request tries <strong>Gemini</strong> first</div>
              <div>2️⃣ Quota hit? Auto-switches to <strong>HuggingFace</strong></div>
              <div>3️⃣ No more quota errors</div>
            </div>
          </div>
        </div>
      )}

      <style>{`@keyframes blink{0%,100%{opacity:1}50%{opacity:0}} @keyframes bounce{0%,100%{transform:translateY(0)}50%{transform:translateY(-4px)}}`}</style>
    </aside>
  )
}
