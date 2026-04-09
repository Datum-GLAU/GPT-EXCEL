import { memo, useState, useRef, useEffect, useCallback } from 'react'
import { useSelector } from 'react-redux'
import { RootState } from '../index'

interface Msg {
  id: string
  role: 'user' | 'ai'
  content: string
  rows?: any[][]
  isStreaming?: boolean
  time: string
}

interface Props {
  currentFile?: any
  rawRows?: any[][]
  stats?: any
  panelWidth?: number
  initialCollapsed?: boolean
  onGridUpdate?: (rows: string[][], desc: string) => void
  onNewFile?: (file: any, rows: any[][]) => void
  onShowChart?: (config: { chartType: string, chartDataKey: string, title: string }) => void
  onAnalysisResult?: (result: any) => void
  onDocumentResult?: (doc: { id: string; title: string; type: string; content: string; createdAt: string; length: 'short' | 'medium' | 'long' }) => void
  onApplyPreview?: () => void
  onDiscardPreview?: () => void
  onUndoLastChange?: () => void
}

const BASE = 'http://localhost:3001/api/llm'

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
    .replace(/\[([^\]]+)\]\((https?:\/\/[^\s)]+)\)/g, '<a href="$2" target="_blank" rel="noreferrer" style="color:var(--blue);text-decoration:underline">$1</a>')
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
    </div>
  )
})

export default function AIChatPanel({
  currentFile,
  rawRows,
  stats,
  panelWidth = 360,
  initialCollapsed = false,
  onGridUpdate,
  onNewFile,
  onShowChart,
  onAnalysisResult,
  onDocumentResult,
  onApplyPreview,
  onDiscardPreview,
  onUndoLastChange
}: Props) {
  const user = useSelector((s: RootState) => s.app.user)
  const chatStorageKey = `ai_chat_history_${currentFile?.id || 'global'}`
  const [tab, setTab] = useState<'chat'>('chat')
  const [collapsed, setCollapsed] = useState(initialCollapsed)
  const [msgs, setMsgs] = useState<Msg[]>([])
  const [busy, setBusy] = useState(false)
  const msgsEndRef = useRef<HTMLDivElement>(null)

  const [genPrompt, setGenPrompt] = useState('')
  const [genBusy, setGenBusy] = useState(false)
  const [sheetPrompt, setSheetPrompt] = useState('')
  const [sheetBusy, setSheetBusy] = useState(false)
  const [pendingDocRequest, setPendingDocRequest] = useState<{ text: string; contextPayload: any } | null>(null)

  const [provider, setProvider] = useState<'gemini' | 'claude'>('gemini')
  const [geminiKey, setGeminiKey] = useState('')
  const [anthropicKey, setAnthropicKey] = useState('')
  const [geminiInput, setGeminiInput] = useState('')
  const [anthropicInput, setAnthropicInput] = useState('')
  const [saving, setSaving] = useState(false)
  const [keyStatus, setKeyStatus] = useState({
    provider: 'gemini' as 'gemini' | 'claude',
    gemini: false,
    anthropic: false,
    hf: false
  })
  const chatBodyRef = useRef<HTMLDivElement>(null)
  const chatScrollTopRef = useRef(0)
  const shouldStickToBottomRef = useRef(true)

  useEffect(() => {
    if (shouldStickToBottomRef.current) {
      msgsEndRef.current?.scrollIntoView({ behavior: 'smooth' })
      return
    }
    if (chatBodyRef.current) chatBodyRef.current.scrollTop = chatScrollTopRef.current
  }, [msgs.length, msgs[msgs.length - 1]?.content?.length])

  const refreshKeyStatus = useCallback(async () => {
    try {
      const res = await fetch('http://localhost:3001/api/llm/status')
      const result = await res.json()
      if (!res.ok) throw new Error(result.error || 'Status fetch failed')
      const nextProvider = result.provider === 'claude' ? 'claude' : 'gemini'
      setKeyStatus({
        provider: nextProvider,
        gemini: !!result.gemini,
        anthropic: !!result.anthropic,
        hf: !!result.anthropic
      })
    } catch {}
  }, [])

  useEffect(() => {
    refreshKeyStatus()
  }, [refreshKeyStatus])

  useEffect(() => {
    try {
      const saved = JSON.parse(localStorage.getItem(chatStorageKey) || '[]')
      setMsgs(Array.isArray(saved)
        ? saved.map((msg: any) => ({
            id: String(msg.id || uid()),
            role: msg.role === 'user' ? 'user' : 'ai',
            content: String(msg.content || ''),
            rows: Array.isArray(msg.rows) ? msg.rows : undefined,
            time: String(msg.time || now()),
            isStreaming: false
          }))
        : [])
    } catch {
      setMsgs([])
    }
  }, [chatStorageKey])

  useEffect(() => {
    try {
      const persistable = msgs
        .filter(msg => !msg.isStreaming)
        .map(msg => ({
          id: msg.id,
          role: msg.role,
          content: msg.content,
          rows: msg.rows,
          time: msg.time
        }))
      localStorage.setItem(chatStorageKey, JSON.stringify(persistable))
    } catch {}
  }, [chatStorageKey, msgs])

  const anyKeySet = keyStatus.gemini || keyStatus.anthropic

  const addMsg = useCallback((role: Msg['role'], content: string, rows?: any[][]): string => {
    const id = uid()
    setMsgs(prev => [...prev, { id, role, content, rows, time: now() }])
    return id
  }, [])

  const updateMsg = useCallback((id: string, content: string, done = false) => {
    setMsgs(prev => prev.map(m => m.id === id ? { ...m, content, isStreaming: !done } : m))
  }, [])

  const startTimedAiStatus = useCallback((label: string) => {
    const id = uid()
    const startedAt = Date.now()
    setMsgs(prev => [...prev, { id, role: 'ai', content: `${label}... 0s`, isStreaming: true, time: now() }])
    const timer = window.setInterval(() => {
      const seconds = Math.max(1, Math.floor((Date.now() - startedAt) / 1000))
      updateMsg(id, `${label}... ${seconds}s`, false)
    }, 1000)
    return {
      id,
      stop(finalContent?: string) {
        window.clearInterval(timer)
        if (finalContent) updateMsg(id, finalContent, true)
        else {
          setMsgs(prev => prev.filter(msg => msg.id !== id))
        }
      }
    }
  }, [updateMsg])

  const generateDocumentFromPrompt = useCallback(async (
    requestText: string,
    contextPayload: any,
    requestedLength: 'short' | 'medium' | 'long'
  ) => {
    const docType =
      /\bproposal\b/i.test(requestText) ? 'proposal' :
      /\bnotice\b/i.test(requestText) ? 'notice' :
      /\breport\b/i.test(requestText) ? 'report' :
      'document'

    const res = await fetch('http://localhost:3001/api/documents/generate', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        type: docType,
        data: {
          prompt: `${requestText}\n\nLength: ${requestedLength}. Keep the document ${requestedLength} unless the user explicitly asks otherwise.`,
          notes: [
            `Source sheet: ${contextPayload.fileName}`,
            `Rows: ${contextPayload.data.rowCount}`,
            contextPayload.data.headers?.length ? `Columns: ${contextPayload.data.headers.join(', ')}` : '',
            contextPayload.data.sampleRows?.length
              ? `Sample rows:\n${contextPayload.data.sampleRows.map((row: any[]) => row.map(cell => String(cell ?? '')).join(' | ')).join('\n')}`
              : ''
          ].filter(Boolean).join('\n'),
          user: user?.name,
          department: user?.department
        },
        title: requestText.slice(0, 60)
      })
    })
    const result = await res.json()
    if (!res.ok) throw new Error(result.error || 'Document generation failed')

    const doc = {
      ...result,
      id: result.id || `doc_${Date.now()}`,
      title: asText(result.title || requestText.slice(0, 50)),
      type: result.type || docType,
      content: asText(result.content),
      createdAt: result.createdAt || new Date().toISOString(),
      length: requestedLength
    } as const

    try {
      const existing = JSON.parse(localStorage.getItem('ai_documents') || '[]')
      localStorage.setItem('ai_documents', JSON.stringify([doc, ...existing.filter((item: any) => item.id !== doc.id)]))
      localStorage.setItem('ai_documents_last_open', doc.id)
    } catch {}

    onDocumentResult?.(doc)
    addMsg('ai', `Created **${doc.title}** as a ${requestedLength} document. The preview is ready in Analysis.`)
  }, [addMsg, onDocumentResult, user?.department, user?.name])

  const persistAiSettings = useCallback(async (nextProvider: 'gemini' | 'claude', nextGemini: string, nextAnthropic: string) => {
    const res = await fetch('http://localhost:3001/api/llm/save-keys', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ provider: nextProvider, geminiKey: nextGemini, anthropicKey: nextAnthropic })
    })
    const result = await res.json()
    if (!res.ok) throw new Error(result.error || 'Failed to save AI settings')
    return result
  }, [])

  const saveKeys = async () => {
    setSaving(true)
    const newGemini = geminiInput.trim() || geminiKey
    const nextProvider = provider
    const newAnthropic = anthropicInput.trim() || anthropicKey
    try {
      await persistAiSettings(nextProvider, newGemini, newAnthropic)
      localStorage.setItem('ai_provider', nextProvider)
      if (geminiInput.trim()) { setGeminiKey(newGemini); localStorage.setItem('gemini_key', newGemini); setGeminiInput('') }
      if (anthropicInput.trim()) { setAnthropicKey(newAnthropic); localStorage.setItem('anthropic_key', newAnthropic); setAnthropicInput('') }
      setKeyStatus({ provider: nextProvider, gemini: !!newGemini, anthropic: !!newAnthropic, hf: !!newAnthropic })
      addMsg('ai', `AI settings saved. ${nextProvider === 'claude' ? 'Claude' : 'Gemini'} is selected.`)
      setTab('chat')
    } catch { addMsg('ai', 'Failed to save keys.') }
    setSaving(false)
  }

  const removeSavedKey = useCallback(async (key: 'gemini' | 'anthropic') => {
    try {
      const nextGemini = key === 'gemini' ? '' : geminiKey
      const nextAnthropic = key === 'anthropic' ? '' : anthropicKey
      await persistAiSettings(provider, nextGemini, nextAnthropic)
      if (key === 'gemini') {
        setGeminiKey('')
        setGeminiInput('')
        localStorage.removeItem('gemini_key')
      } else {
        setAnthropicKey('')
        setAnthropicInput('')
        localStorage.removeItem('anthropic_key')
      }
      setKeyStatus({ provider, gemini: !!nextGemini, anthropic: !!nextAnthropic, hf: !!nextAnthropic })
      addMsg('ai', `${key === 'gemini' ? 'Gemini' : 'Claude'} key removed.`)
    } catch {
      addMsg('ai', 'Failed to remove key.')
    }
  }, [addMsg, anthropicKey, geminiKey, persistAiSettings, provider])

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
    } catch { addMsg('ai', 'Export failed.') }
  }

  const streamChat = useCallback(async (text: string) => {
    const aiId = uid()
    setMsgs(prev => [...prev, { id: aiId, role: 'ai', content: '', isStreaming: true, time: now() }])
    if (!anyKeySet) {
      updateMsg(aiId, 'AI is not ready. Add `GEMINI_API_KEY` or `ANTHROPIC_API_KEY` in `Ai_Agent/.env`, then restart the backend.', true)
      return
    }
    try {
      const res = await fetch(`${BASE}/stream`, {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ message: text, rawRows, stats })
      })
      if (!res.ok || !res.body) { updateMsg(aiId, 'AI could not send a reply right now. Please try again.', true); return }
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
            if (d.error) { updateMsg(aiId, `AI error: ${d.error}`, true); return }
            if (d.text) { full += d.text; updateMsg(aiId, full, false) }
            if (d.done) updateMsg(aiId, full || 'AI could not send a reply. Please try again.', true)
          } catch {}
        }
      }
      updateMsg(aiId, full || 'AI could not send a reply. Please try again.', true)
    } catch (e: any) { updateMsg(aiId, `AI error: ${e.message}`, true) }
  }, [anyKeySet, rawRows, stats, updateMsg])

  const send = useCallback(async (msg?: string) => {
    const text = (msg || '').trim()
    if (!text || busy) return
    setBusy(true)
    addMsg('user', text)
    const requestedLength = /\bshort\b/i.test(text)
      ? 'short'
      : /\b(mid|medium)\b/i.test(text)
        ? 'medium'
        : /\blong\b/i.test(text)
          ? 'long'
          : null

    if (pendingDocRequest) {
      if (!requestedLength) {
        addMsg('ai', 'Reply with one document size: `short`, `medium`, or `long`.')
        setBusy(false)
        return
      }
      try {
        await generateDocumentFromPrompt(pendingDocRequest.text, pendingDocRequest.contextPayload, requestedLength)
      } catch (e: any) {
        addMsg('ai', `Document request failed: ${e.message}`)
      }
      setPendingDocRequest(null)
      setBusy(false)
      return
    }
    const lowerText = text.toLowerCase()
    if (/\b(apply|confirm)\b.*\b(preview|changes?)\b/i.test(text) && onApplyPreview) {
      onApplyPreview()
      addMsg('ai', 'Applied the current preview to the sheet.')
      setBusy(false)
      return
    }
    if (/\b(discard|cancel)\b.*\b(preview|changes?)\b/i.test(text) && onDiscardPreview) {
      onDiscardPreview()
      addMsg('ai', 'Discarded the current preview.')
      setBusy(false)
      return
    }
    if (/\bundo\b.*\b(change|preview|last)\b/i.test(text) && onUndoLastChange) {
      onUndoLastChange()
      addMsg('ai', 'Undid the last AI-applied change.')
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
      let progress: ReturnType<typeof startTimedAiStatus> | null = null
      try {
        if (!requestedLength) {
          setPendingDocRequest({ text, contextPayload })
          addMsg('ai', 'What size should the document be: `short`, `medium`, or `long`?')
        } else {
          progress = startTimedAiStatus('Creating document')
          await generateDocumentFromPrompt(text, contextPayload, requestedLength)
          progress.stop()
        }
      } catch (e: any) {
        progress?.stop()
        addMsg('ai', `Document request failed: ${e.message}`)
      }
      setBusy(false)
      return
    }
    if ((/\b(formula|excel formula|spreadsheet formula|function)\b/i.test(text) || /\b(sumif|sumifs|countif|countifs|averageif|averageifs|vlookup|xlookup|index match|iferror)\b/i.test(text)) && rawRows?.[0]?.length) {
      let progress: ReturnType<typeof startTimedAiStatus> | null = null
      try {
        progress = startTimedAiStatus('Creating formula')
        const headers = rawRows[0].map((h: any) => String(h))
        const res = await fetch(`${BASE}/formula`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ description: text, headers })
        })
        const result = await res.json()
        if (!res.ok) throw new Error(result.error || 'Formula generation failed')
        const parts = [
          result.formula ? `**Formula**\n\`${result.formula}\`` : '',
          result.explanation ? `**Explanation**\n${result.explanation}` : '',
          result.example ? `**Example**\n${result.example}` : ''
        ].filter(Boolean)
        progress.stop()
        addMsg('ai', parts.join('\n\n') || 'I generated a formula suggestion.')
      } catch (e: any) {
        progress?.stop()
        addMsg('ai', `Formula request failed: ${e.message}`)
      }
      setBusy(false)
      return
    }
    if (/\b(create|generate|make)\b.*\b(sheet|table|excel|spreadsheet)\b/i.test(text)) {
      let progress: ReturnType<typeof startTimedAiStatus> | null = null
      try {
        progress = startTimedAiStatus(rawRows?.length ? 'Creating new sheet from current workbook' : 'Creating workbook')
        const endpoint = rawRows?.length ? `${BASE}/new-sheet` : `${BASE}/generate-excel`
        const payload = rawRows?.length
          ? { instruction: text, rawRows, stats, department: user?.department }
          : { prompt: text, department: user?.department || 'general' }
        const res = await fetch(endpoint, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(payload)
        })
        const result = await res.json()
        if (!res.ok) throw new Error(result.error || 'Spreadsheet generation failed')
        progress.stop()
        if (onNewFile && result.file) onNewFile(result.file, result.rows)
        onAnalysisResult?.({
          kind: 'file',
          title: asText(result.sheet_name || result.file?.name || 'New sheet'),
          description: asText(result.description || 'A new workbook was created from chat.'),
          rows: result.rows || [],
          data: [],
          createdAt: new Date().toISOString()
        })
        const rowCount = result.rowCount ? ` (${result.rowCount} rows)` : ''
        const openUrl = result.file?.url ? `http://localhost:3001${result.file.url}` : ''
        const openLine = openUrl ? `\n\n[Open workbook](${openUrl})` : ''
        addMsg('ai', `Created **${asText(result.sheet_name || result.file?.name || 'New sheet')}**${rowCount} from your chat request.${result.fallback ? '\n\nBuilt from the current sheet using local fallback because the AI quota is exhausted.' : ''}${openLine}`, result.rows)
      } catch (e: any) {
        progress?.stop()
        addMsg('ai', `Spreadsheet request failed: ${e.message}`)
      }
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
        const progress = startTimedAiStatus('Preparing chart')
        onShowChart({ chartType, chartDataKey, title: `AI chart: ${text}` })
        progress.stop()
        addMsg('ai', `Prepared a ${chartType} chart preview for ${chartDataKey} data. Review it in the sheet area, then apply or discard it.`)
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
          const summary = asText(result.answer || result.description || 'Done.')
          onAnalysisResult?.({
            kind: 'analysis',
            title: 'AI analysis result',
            description: summary,
            data: result.data || [],
            createdAt: new Date().toISOString()
          })
          addMsg('ai', summary)
        } else if (result.rows?.length > 1) {
          const summary = [asText(result.description), asText(result.answer)].filter(Boolean).join('\n\n')
          const msgId = addMsg('ai', summary || `**Result: ${result.rows.length - 1} rows**`)
          setMsgs(prev => prev.map(m => m.id === msgId ? { ...m, rows: result.rows } : m))
          onAnalysisResult?.({
            kind: 'analysis',
            title: 'AI analysis result',
            description: summary || 'Chat generated a tabular result.',
            rows: result.rows,
            data: result.data || [],
            createdAt: new Date().toISOString()
          })
          if (result.action === 'add_column' && onGridUpdate) {
            onGridUpdate(result.rows, asText(result.description || 'Column added'))
            addMsg('ai', `Preview ready. Column "${result.column_name}" can be reviewed before applying.`)
          }
        } else { await streamChat(text) }
      } catch { await streamChat(text) }
    } else { await streamChat(text) }
    setBusy(false)
  }, [busy, pendingDocRequest, rawRows, stats, onGridUpdate, onNewFile, onShowChart, onAnalysisResult, onApplyPreview, onDiscardPreview, onUndoLastChange, addMsg, streamChat, currentFile?.name, user?.department, user?.name, generateDocumentFromPrompt])

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
    } catch (e: any) { addMsg('ai', `Request failed: ${e.message}`); setTab('chat') }
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
    } catch (e: any) { addMsg('ai', `Request failed: ${e.message}`); setTab('chat') }
    setSheetBusy(false)
  }

  const hasHiddenStreamingAi = msgs.some(m => m.role === 'ai' && m.isStreaming && !m.content.trim())
  const hasVisibleStreamingAi = msgs.some(m => m.role === 'ai' && m.isStreaming && !!m.content.trim())

  if (collapsed) {
    return (
      <div onClick={() => setCollapsed(false)} style={{ width: 36, background: 'var(--bg)', borderLeft: '1px solid var(--border)', display: 'flex', flexDirection: 'column', alignItems: 'center', paddingTop: 14, gap: 8, cursor: 'pointer', flexShrink: 0 }}>
        <span style={{ fontSize: '0.58rem', color: 'var(--text-muted)', writingMode: 'vertical-lr' as any, textTransform: 'uppercase', letterSpacing: '0.06em' }}>AI Panel</span>
        <div style={{ width: 6, height: 6, borderRadius: '50%', background: anyKeySet ? 'var(--green)' : 'var(--orange)' }} />
      </div>
    )
  }

  return (
    <aside style={{ width: panelWidth, minWidth: 280, maxWidth: 620, flexShrink: 0, background: 'var(--bg)', borderLeft: '1px solid var(--border)', display: 'flex', flexDirection: 'column' }}>
      <div style={{ height: 44, display: 'flex', alignItems: 'center', padding: '0 10px', borderBottom: '1px solid var(--border)', gap: 7, flexShrink: 0 }}>
        <div style={{ width: 24, height: 24, background: 'var(--blue)', borderRadius: 5, display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: '0.6rem', fontWeight: 800, color: '#fff', flexShrink: 0 }}>
          AI
        </div>
        <div style={{ flex: 1, minWidth: 0 }}>
          <div style={{ fontSize: '0.75rem', fontWeight: 600, color: 'var(--text)' }}>Xtron</div>
          <div style={{ fontSize: '0.6rem', display: 'flex', alignItems: 'center', gap: 3, color: anyKeySet ? 'var(--green)' : 'var(--orange)' }}>
            <div style={{ width: 4, height: 4, borderRadius: '50%', background: anyKeySet ? 'var(--green)' : 'var(--orange)', flexShrink: 0 }} />
            {anyKeySet ? `${keyStatus.provider === 'claude' ? 'Claude' : 'Gemini'} · Ready` : 'AI mode · Configure Ai_Agent/.env'}
          </div>
        </div>
        {currentFile && <div style={{ fontSize: '0.58rem', color: 'var(--text-muted)', maxWidth: 60, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>📄 {currentFile.name}</div>}
        <button onClick={() => setCollapsed(true)} style={{ background: 'none', border: 'none', color: 'var(--text-muted)', cursor: 'pointer', fontSize: 18, lineHeight: 1, padding: '0 2px', flexShrink: 0 }}>›</button>
      </div>

      {tab === 'chat' && (
        <>
          {!anyKeySet && (
            <div style={{ padding: '8px 10px', background: 'rgba(234,179,8,0.08)', borderBottom: '1px solid rgba(234,179,8,0.2)', fontSize: '0.65rem', color: 'var(--orange)' }}>
              <span>Set `GEMINI_API_KEY` or `ANTHROPIC_API_KEY` in `Ai_Agent/.env` or your terminal environment, then restart the backend.</span>
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
            style={{ flex: 1, overflowY: 'auto', padding: '10px 8px', display: 'flex', flexDirection: 'column', gap: 10, scrollbarWidth: 'none', msOverflowStyle: 'none' }}
          >
            {msgs.filter(m => !(m.role === 'ai' && m.isStreaming && !m.content.trim())).map(m => (
              <div key={m.id} style={{ display: 'flex', gap: 6, flexDirection: m.role === 'user' ? 'row-reverse' : 'row', alignItems: 'flex-start' }}>
                {m.role !== 'user' && (
                  <div style={{ width: 22, height: 22, borderRadius: 5, background: m.role === 'ai' ? 'var(--blue)' : 'var(--surface-2)', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: '0.55rem', fontWeight: 800, color: m.role === 'ai' ? '#fff' : 'var(--text-muted)', flexShrink: 0, marginTop: 2 }}>
                    {m.role === 'ai' ? 'AI' : '⚙'}
                  </div>
                )}
                <div style={{ maxWidth: '88%', display: 'flex', flexDirection: 'column', gap: 3, alignItems: m.role === 'user' ? 'flex-end' : 'flex-start' }}>
                  <div style={{ padding: '7px 10px', background: m.role === 'user' ? 'var(--blue)' : 'var(--surface-2)', color: m.role === 'user' ? '#fff' : 'var(--text)', borderRadius: m.role === 'user' ? '10px 10px 2px 10px' : '10px 10px 10px 2px', border: '1px solid var(--border)', fontSize: '0.76rem', lineHeight: 1.55, wordBreak: 'break-word', userSelect: 'text', WebkitUserSelect: 'text', cursor: 'text' }}>
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
          <ChatComposer busy={busy} currentFile={currentFile} onSend={send} />
        </>
      )}
      <style>{`@keyframes blink{0%,100%{opacity:1}50%{opacity:0}} @keyframes bounce{0%,100%{transform:translateY(0)}50%{transform:translateY(-4px)}} aside ::-webkit-scrollbar{display:none}`}</style>
    </aside>
  )
}





