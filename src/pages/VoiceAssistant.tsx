import { useEffect, useMemo, useRef, useState } from 'react'
import { useNavigate } from 'react-router-dom'
import Sidebar from '../components/Sidebar'
import Header from '../components/Header'

type Preset = 'lady' | 'deep' | 'british' | 'boy'
type Msg = { id: string; role: 'user' | 'ai' | 'system'; content: string; time: string }
type Rec = {
  continuous: boolean
  interimResults: boolean
  lang: string
  onstart: null | (() => void)
  onend: null | (() => void)
  onerror: null | ((e: { error?: string }) => void)
  onresult: null | ((e: any) => void)
  start: () => void
  stop: () => void
}
type RecCtor = new () => Rec

declare global {
  interface Window {
    webkitSpeechRecognition?: RecCtor
    SpeechRecognition?: RecCtor
  }
}

const BASE = 'http://localhost:3001/api/llm'
const KEYS = { hist: 'voice_hist_v3', preset: 'voice_preset', auto: 'voice_auto' }
const PRESETS: Array<{ id: Preset; title: string; hint: string; match: string[]; lang?: string }> = [
  { id: 'lady', title: 'Lady', hint: 'Warm', match: ['zira', 'aria', 'jenny', 'samantha', 'female'] },
  { id: 'deep', title: 'Deep', hint: 'Low', match: ['david', 'guy', 'male', 'brian', 'mark'] },
  { id: 'british', title: 'British', hint: 'UK', match: ['uk', 'british', 'hazel', 'arthur'], lang: 'en-GB' },
  { id: 'boy', title: 'Boy', hint: 'Light', match: ['ryan', 'young', 'junior', 'tom'] },
]
const QUICK = ['Open Excel sheet', 'Create bar chart', 'Analyze this workbook', 'Create document draft', 'Create PPT from sheet', 'Open File Manager']
let uid = 0
const id = () => `voice_${++uid}_${Date.now()}`
const tm = () => new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })
const txt = (v: any): string => typeof v === 'string' ? v : v == null ? '' : Array.isArray(v) ? v.map(txt).filter(Boolean).join('\n') : typeof v === 'object' ? (typeof v.text === 'string' ? v.text : typeof v.content === 'string' ? v.content : JSON.stringify(v, null, 2)) : String(v)
const intent = (tab: 'analyze' | 'charts' | 'create', prompt: string, questions: string[]) => ({ tab, prompt, questions })
const voicePick = (voices: SpeechSynthesisVoice[], preset: Preset) => {
  const cfg = PRESETS.find(v => v.id === preset) || PRESETS[0]
  const matchers = cfg.match.map(v => v.toLowerCase())
  return voices.find(v => {
    const hay = `${v.name} ${v.lang}`.toLowerCase()
    return (!cfg.lang || v.lang.toLowerCase().startsWith(cfg.lang.toLowerCase())) && matchers.some(m => hay.includes(m))
  }) || voices.find(v => cfg.lang ? v.lang.toLowerCase().startsWith(cfg.lang.toLowerCase()) : v.lang.toLowerCase().startsWith('en')) || voices[0] || null
}

export default function VoiceAssistant() {
  const navigate = useNavigate()
  const [sidebarOpen, setSidebarOpen] = useState(true)
  const [msgs, setMsgs] = useState<Msg[]>(() => {
    try {
      const saved = JSON.parse(localStorage.getItem(KEYS.hist) || '[]')
      if (Array.isArray(saved) && saved.length) return saved
    } catch {}
    return [{ id: id(), role: 'ai', content: 'Voice assistant is ready. Speak or type a command to open workspaces, prepare charts, create docs or PPT drafts, or ask AI questions.', time: tm() }]
  })
  const [input, setInput] = useState('')
  const [draft, setDraft] = useState('')
  const [listening, setListening] = useState(false)
  const [busy, setBusy] = useState(false)
  const [supported, setSupported] = useState(false)
  const [backend, setBackend] = useState<boolean | null>(null)
  const [keysReady, setKeysReady] = useState(false)
  const [status, setStatus] = useState('Ready for voice or typed commands')
  const [preset, setPreset] = useState<Preset>(() => (localStorage.getItem(KEYS.preset) as Preset) || 'lady')
  const [autoSpeak, setAutoSpeak] = useState(() => localStorage.getItem(KEYS.auto) !== 'false')
  const [voices, setVoices] = useState<SpeechSynthesisVoice[]>([])
  const recRef = useRef<Rec | null>(null)
  const bodyRef = useRef<HTMLDivElement>(null)

  useEffect(() => { localStorage.setItem(KEYS.hist, JSON.stringify(msgs.slice(-30))) }, [msgs])
  useEffect(() => { localStorage.setItem(KEYS.preset, preset) }, [preset])
  useEffect(() => { localStorage.setItem(KEYS.auto, String(autoSpeak)) }, [autoSpeak])
  useEffect(() => { bodyRef.current?.scrollTo({ top: bodyRef.current.scrollHeight, behavior: 'smooth' }) }, [msgs, draft, busy])

  useEffect(() => {
    const Ctor = window.SpeechRecognition || window.webkitSpeechRecognition
    if (!Ctor) { setSupported(false); setStatus('Speech recognition is not available here. Typed AI still works.'); return }
    setSupported(true)
    const rec = new Ctor()
    rec.continuous = false
    rec.interimResults = true
    rec.lang = 'en-US'
    rec.onstart = () => { setListening(true); setDraft(''); setStatus('Listening') }
    rec.onend = () => { setListening(false); setStatus(s => s === 'Listening' ? 'Mic stopped' : s) }
    rec.onerror = e => { setListening(false); setStatus(e?.error === 'not-allowed' ? 'Microphone permission was denied.' : `Voice capture stopped: ${e?.error || 'unknown'}`) }
    rec.onresult = (e: any) => {
      let interim = '', finalText = ''
      for (let i = e.resultIndex; i < e.results.length; i++) {
        const chunk = e.results[i][0]?.transcript || ''
        if (e.results[i].isFinal) finalText += chunk
        else interim += chunk
      }
      if (interim) setDraft(interim.trim())
      if (finalText.trim()) { setDraft(''); setInput(finalText.trim()); void submit(finalText.trim()) }
    }
    recRef.current = rec
    return () => rec.stop()
  }, [])

  useEffect(() => {
    const load = () => setVoices(window.speechSynthesis?.getVoices?.() || [])
    load()
    window.speechSynthesis?.addEventListener?.('voiceschanged', load)
    return () => window.speechSynthesis?.removeEventListener?.('voiceschanged', load)
  }, [])

  useEffect(() => {
    const check = async () => {
      try {
        const res = await fetch(`${BASE}/key-status`)
        if (!res.ok) throw new Error('down')
        const data = await res.json()
        setBackend(true)
        setKeysReady(!!(data.gemini || data.hf))
      } catch {
        setBackend(false)
        setKeysReady(!!(localStorage.getItem('gemini_key') || localStorage.getItem('hf_key')))
      }
    }
    void check()
  }, [])

  const voice = useMemo(() => voicePick(voices, preset), [voices, preset])
  const add = (role: Msg['role'], content: string) => setMsgs(prev => [...prev, { id: id(), role, content, time: tm() }].slice(-30))
  const patchAi = (content: string) => setMsgs(prev => {
    const copy = [...prev]
    for (let i = copy.length - 1; i >= 0; i--) if (copy[i].role === 'ai') { copy[i] = { ...copy[i], content }; return copy }
    return [...copy, { id: id(), role: 'ai', content, time: tm() }]
  })
  const speak = (content: string) => {
    if (!content.trim() || !window.speechSynthesis) return
    window.speechSynthesis.cancel()
    const u = new SpeechSynthesisUtterance(content)
    u.voice = voice
    u.lang = voice?.lang || 'en-US'
    u.rate = preset === 'boy' ? 1.04 : preset === 'deep' ? 0.92 : 0.98
    u.pitch = preset === 'deep' ? 0.78 : preset === 'boy' ? 1.16 : 1
    window.speechSynthesis.speak(u)
  }

  const runLocal = (raw: string) => {
    const t = raw.toLowerCase()
    if (/\b(open|go to|show)\b.*\bdashboard\b/.test(t)) { navigate('/dashboard'); return 'Opened Dashboard.' }
    if (/\b(open|go to|show)\b.*\b(file manager|files)\b/.test(t)) { navigate('/file-manager'); return 'Opened File Manager.' }
    if (/\b(open|go to|show)\b.*\bsettings\b/.test(t)) { navigate('/settings'); return 'Opened Settings.' }
    if (/\b(open|go to|show)\b.*\bworkflow\b/.test(t)) { navigate('/workflow'); return 'Opened Workflow.' }
    if (/\b(open|go to|show)\b.*\bdocuments?\b/.test(t)) { navigate('/documents'); return 'Opened Documents.' }
    if (/\b(open|go to|show)\b.*\b(ppt|powerpoint|presentation|slides|deck)\b/.test(t)) { navigate('/powerpoint'); return 'Opened PPT workspace.' }
    if (/\b(open|go to|show)\b.*\b(excel|sheet|worksheet)\b/.test(t)) { navigate('/excel'); return 'Opened Excel Sheet.' }
    if (/\b(create|generate|make)\b.*\b(doc|document|report|notice|proposal)\b/.test(t)) {
      const draftDoc = { prompt: raw, fileName: 'Voice request', data: { rowCount: 0, headers: [], sampleRows: [] } }
      sessionStorage.setItem('ai_document_draft', JSON.stringify(draftDoc))
      navigate('/documents', { state: { aiDraft: draftDoc } })
      return 'Moved this request to Documents and prepared a draft.'
    }
    if (/\b(create|generate|make)\b.*\b(ppt|powerpoint|presentation|deck|slides)\b/.test(t)) {
      const draftPpt = { prompt: raw, fileName: 'Voice request', data: { rowCount: 0, headers: [], sampleRows: [] } }
      sessionStorage.setItem('ai_ppt_draft', JSON.stringify(draftPpt))
      navigate('/powerpoint', { state: { aiDraft: draftPpt } })
      return 'Moved this request to PPT and prepared the deck flow.'
    }
    if (/\b(chart|graph|bar chart|line chart|pie chart|visuali[sz]e)\b/.test(t)) {
      sessionStorage.setItem('excel_workspace_intent', JSON.stringify(intent('charts', raw, ['Which metric should I visualize?', 'Compare sections with a bar chart', 'Show score distribution as line', 'Make a pie chart for grades'])))
      navigate('/excel')
      return 'Opened Excel charts workspace with your prompt ready.'
    }
    if (/\b(analy[sz]e|analysis|summari[sz]e|explain)\b/.test(t)) {
      sessionStorage.setItem('excel_workspace_intent', JSON.stringify(intent('analyze', raw, ['Highlight top risks', 'Summarize key trends', 'Find outliers in the sheet', 'Show performance by section'])))
      navigate('/excel')
      return 'Opened Excel analysis workspace with your prompt ready.'
    }
    if (/\b(create|generate|make|filter|sort|table|new sheet)\b/.test(t)) {
      sessionStorage.setItem('excel_workspace_intent', JSON.stringify(intent('create', raw, ['Add a ranked result table', 'Build a filtered sheet', 'Create a summary table', 'Prepare a new exportable sheet'])))
      navigate('/excel')
      return 'Opened Excel create workspace so you can preview the AI result.'
    }
    return null
  }

  const askAi = async (raw: string) => {
    const res = await fetch(`${BASE}/stream`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ message: raw }) })
    if (!res.ok || !res.body) throw new Error((await res.text().catch(() => '')) || `AI request failed (${res.status})`)
    add('ai', '')
    const reader = res.body.getReader()
    const dec = new TextDecoder()
    let buf = '', full = ''
    while (true) {
      const { done, value } = await reader.read()
      if (done) break
      buf += dec.decode(value, { stream: true })
      const lines = buf.split('\n')
      buf = lines.pop() || ''
      for (const line of lines) {
        if (!line.startsWith('data: ')) continue
        const d = JSON.parse(line.slice(6))
        if (d.error) throw new Error(txt(d.error))
        if (d.text) { full += txt(d.text); patchAi(full) }
      }
    }
    return full.trim()
  }

  const submit = async (raw = input) => {
    const text = raw.trim()
    if (!text || busy) return
    window.speechSynthesis?.cancel()
    setInput('')
    setDraft('')
    setBusy(true)
    setStatus('Processing your request')
    add('user', text)
    const local = runLocal(text)
    if (local) {
      add('system', local)
      if (autoSpeak) speak(local)
      setStatus('Command completed')
      setBusy(false)
      return
    }
    try {
      const reply = await askAi(text)
      const spoken = reply || 'I did not receive a full response. Please try again.'
      setStatus('AI response ready')
      if (autoSpeak) speak(spoken)
    } catch (e: any) {
      const err = e?.message || 'Voice assistant request failed.'
      add('system', err)
      setStatus(err)
    }
    setBusy(false)
  }

  const start = () => {
    if (!recRef.current) { setStatus('Speech recognition is not available here.'); return }
    try { window.speechSynthesis?.cancel(); setDraft(''); recRef.current.start() } catch { setStatus('Microphone is already active.') }
  }
  const stop = () => { recRef.current?.stop(); window.speechSynthesis?.cancel(); setListening(false); setStatus('Stopped voice capture and playback') }
  const clear = () => { window.speechSynthesis?.cancel(); setMsgs([{ id: id(), role: 'ai', content: 'Fresh voice session ready. Speak or type a command to continue.', time: tm() }]); setStatus('Conversation cleared') }
  const startListening = start
  const transcript = draft || input
  const response = [...msgs].reverse().find(m => m.role === 'ai' && m.content.trim())?.content || ''
  const history = msgs.reduce<{ text: string; response: string }[]>((acc, msg) => {
    if (msg.role === 'user') acc.push({ text: msg.content, response: '' })
    else if ((msg.role === 'ai' || msg.role === 'system') && acc.length) {
      acc[acc.length - 1].response = acc[acc.length - 1].response ? `${acc[acc.length - 1].response}\n${msg.content}` : msg.content
    }
    return acc
  }, []).slice(-10).reverse()

  return (
    <div style={{ display: 'flex', flexDirection: 'column', height: '100vh', background: 'var(--bg)' }}>
      <Header toggleSidebar={() => setSidebarOpen(p => !p)} />
      <div style={{ display: 'flex', flex: 1, overflow: 'hidden' }}>
        <Sidebar isOpen={sidebarOpen} />
        <main style={{ flex: 1, overflowY: 'auto', padding: '24px 28px', background: 'linear-gradient(180deg, var(--bg) 0%, var(--bg-2) 100%)' }}>
          <h1 style={{ fontSize: '1.5rem', letterSpacing: -0.5, marginBottom: 8 }}>Voice Assistant</h1>
          <p style={{ fontSize: '0.8rem', color: 'var(--text-secondary)', marginBottom: 32 }}>Hands‑free commands with Whisper API integration</p>
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4, minmax(0, 1fr))', gap: 10, marginBottom: 18 }}>
            {[
              { label: 'Mic', value: supported ? (listening ? 'Listening' : 'Ready') : 'Typed only' },
              { label: 'AI', value: busy ? 'Processing' : 'Standby' },
              { label: 'Backend', value: backend === null ? 'Checking' : backend ? 'Connected' : 'Offline' },
              { label: 'Keys', value: keysReady ? 'Ready' : 'Missing' }
            ].map(card => (
              <div key={card.label} style={{ padding: '12px 14px', borderRadius: 10, background: 'var(--surface)', border: '1px solid var(--border)' }}>
                <div style={{ fontSize: '0.66rem', color: 'var(--text-muted)', textTransform: 'uppercase', letterSpacing: '0.08em', marginBottom: 4 }}>{card.label}</div>
                <div style={{ fontSize: '0.9rem', color: 'var(--text)', fontWeight: 700 }}>{card.value}</div>
              </div>
            ))}
          </div>

          <div
            style={{
              background: 'var(--surface)',
              border: '1px solid var(--border)',
              borderRadius: 12,
              padding: '32px',
              textAlign: 'center',
              marginBottom: 32,
            }}
          >
            <button
              onClick={startListening}
              disabled={listening || busy}
              style={{
                width: 80,
                height: 80,
                borderRadius: '50%',
                background: listening ? 'var(--error)' : 'var(--accent)',
                border: 'none',
                cursor: 'pointer',
                marginBottom: 20,
                transition: 'all var(--tr)',
              }}
            >
              <svg width={32} height={32} viewBox="0 0 24 24" fill="none" stroke="white" strokeWidth={2}>
                <path d="M12 1a3 3 0 0 0-3 3v8a3 3 0 0 0 6 0V4a3 3 0 0 0-3-3z" />
                <path d="M19 10v2a7 7 0 0 1-14 0v-2" />
                <line x1={12} y1={19} x2={12} y2={23} />
                <line x1={8} y1={23} x2={16} y2={23} />
              </svg>
            </button>
            <div style={{ fontSize: '1.1rem', marginBottom: 8 }}>
              {listening ? 'Listening...' : transcript ? `"${transcript}"` : status}
            </div>
            {response && (
              <div
                style={{
                  background: 'var(--surface-2)',
                  border: '1px solid var(--border)',
                  borderRadius: 8,
                  padding: '16px',
                  marginTop: 20,
                  textAlign: 'left',
                }}
              >
                <div style={{ fontWeight: 600, marginBottom: 8 }}>Response:</div>
                <p style={{ fontSize: '0.9rem' }}>{response}</p>
              </div>
            )}

            <div style={{ marginTop: 18, textAlign: 'left' }}>
              <div style={{ display: 'flex', flexWrap: 'wrap', gap: 8, marginBottom: 10 }}>
                {QUICK.map(q => (
                  <button key={q} className="btn btn-ghost btn-xs" onClick={() => setInput(q)}>{q}</button>
                ))}
              </div>
              <textarea
                className="input"
                value={input}
                onChange={e => setInput(e.target.value)}
                onKeyDown={e => {
                  if (e.key === 'Enter' && !e.shiftKey) {
                    e.preventDefault()
                    void submit()
                  }
                }}
                placeholder="Type a command or question. Example: create bar chart for section performance"
                style={{ width: '100%', minHeight: 92, resize: 'vertical', marginBottom: 10 }}
              />
              <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap', marginBottom: 10 }}>
                <button className="btn btn-primary" disabled={busy || !input.trim()} onClick={() => void submit()}>{busy ? 'Processing...' : 'Send'}</button>
                <button className="btn btn-secondary" disabled={busy} onClick={start}>Speak</button>
                <button className="btn btn-outline" onClick={stop}>Stop</button>
                <button className="btn btn-ghost" onClick={clear}>Clear</button>
              </div>
              <div style={{ display: 'flex', flexWrap: 'wrap', gap: 8, marginBottom: 10 }}>
                {PRESETS.map(v => (
                  <button key={v.id} className={preset === v.id ? 'btn btn-primary btn-xs' : 'btn btn-secondary btn-xs'} onClick={() => setPreset(v.id)}>
                    {v.title}
                  </button>
                ))}
              </div>
              <label style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: '0.8rem', color: 'var(--text-secondary)' }}>
                <input type="checkbox" checked={autoSpeak} onChange={e => setAutoSpeak(e.target.checked)} />
                Speak AI replies automatically
              </label>
              <div style={{ marginTop: 8, fontSize: '0.75rem', color: 'var(--text-muted)' }}>
                {status}{voice ? ` · Voice: ${voice.name} (${voice.lang})` : ''}
              </div>
            </div>
          </div>

          <h3 style={{ fontSize: '1rem', marginBottom: 12 }}>Command History</h3>
          <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
            {history.length === 0 ? (
              <p style={{ color: 'var(--text-muted)' }}>No commands yet.</p>
            ) : (
              history.map((h, i) => (
                <div
                  key={i}
                  style={{
                    background: 'var(--surface)',
                    border: '1px solid var(--border)',
                    borderRadius: 8,
                    padding: '12px 16px',
                  }}
                >
                  <div style={{ fontWeight: 500, marginBottom: 4 }}>🎤 {h.text}</div>
                  <div style={{ fontSize: '0.8rem', color: 'var(--text-secondary)', whiteSpace: 'pre-wrap', lineHeight: 1.55 }}>{h.response || 'Waiting for a paired AI response...'}</div>
                </div>
              ))
            )}
          </div>
        </main>
      </div>
    </div>
  );
}
