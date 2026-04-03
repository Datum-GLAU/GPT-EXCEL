import { useState } from 'react'

interface AssistantMsg {
  id: string
  role: 'user' | 'ai' | 'system'
  content: string
}

interface Props {
  title: string
  subtitle: string
  placeholder: string
  initialMessage: string
  onSubmit: (text: string) => Promise<string>
  suggestions?: string[]
  busyLabel?: string
}

let assistantUid = 0
const nextId = () => `assistant_${++assistantUid}_${Date.now()}`

export default function WorkspaceAssistant({ title, subtitle, placeholder, initialMessage, onSubmit, suggestions = [], busyLabel = 'Thinking...' }: Props) {
  const [messages, setMessages] = useState<AssistantMsg[]>([
    { id: nextId(), role: 'ai', content: initialMessage }
  ])
  const [value, setValue] = useState('')
  const [busy, setBusy] = useState(false)

  const send = async () => {
    const text = value.trim()
    if (!text || busy) return
    setBusy(true)
    setMessages(prev => [...prev, { id: nextId(), role: 'user', content: text }])
    setValue('')
    try {
      const reply = await onSubmit(text)
      setMessages(prev => [...prev, { id: nextId(), role: 'ai', content: reply }])
    } catch (e: any) {
      setMessages(prev => [...prev, { id: nextId(), role: 'system', content: e.message || 'Request failed' }])
    }
    setBusy(false)
  }

  return (
    <aside style={{ width: 330, flexShrink: 0, borderLeft: '1px solid var(--border)', background: 'var(--bg)', display: 'flex', flexDirection: 'column' }}>
      <div style={{ padding: '12px 14px', borderBottom: '1px solid var(--border)' }}>
        <div style={{ fontSize: '0.86rem', fontWeight: 700, color: 'var(--text)' }}>{title}</div>
        <div style={{ fontSize: '0.7rem', color: 'var(--text-muted)', marginTop: 3 }}>{subtitle}</div>
        <div style={{ marginTop: 10, display: 'flex', alignItems: 'center', gap: 8 }}>
          <div style={{ width: 8, height: 8, borderRadius: '50%', background: busy ? 'var(--orange)' : 'var(--green)', boxShadow: `0 0 0 6px ${busy ? 'rgba(249,115,22,0.12)' : 'rgba(34,197,94,0.12)'}` }} />
          <div style={{ fontSize: '0.66rem', color: 'var(--text-muted)' }}>{busy ? busyLabel : 'Ready for follow-up questions'}</div>
        </div>
      </div>
      <div style={{ flex: 1, overflowY: 'auto', padding: 12, display: 'flex', flexDirection: 'column', gap: 10 }}>
        {suggestions.length > 0 && (
          <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6 }}>
            {suggestions.map(suggestion => (
              <button
                key={suggestion}
                onClick={() => !busy && setValue(suggestion)}
                style={{ padding: '5px 8px', borderRadius: 999, border: '1px solid var(--border)', background: 'var(--surface)', color: 'var(--text-secondary)', fontSize: '0.65rem', cursor: busy ? 'default' : 'pointer', opacity: busy ? 0.55 : 1 }}
              >
                {suggestion}
              </button>
            ))}
          </div>
        )}
        {messages.map(msg => (
          <div key={msg.id} style={{ alignSelf: msg.role === 'user' ? 'flex-end' : 'flex-start', maxWidth: '92%' }}>
            <div style={{ padding: '8px 10px', borderRadius: msg.role === 'user' ? '12px 12px 4px 12px' : '12px 12px 12px 4px', background: msg.role === 'user' ? 'var(--blue)' : msg.role === 'system' ? 'rgba(239,68,68,0.1)' : 'var(--surface-2)', color: msg.role === 'user' ? '#fff' : msg.role === 'system' ? 'var(--red)' : 'var(--text)', border: '1px solid var(--border)', fontSize: '0.75rem', lineHeight: 1.55, whiteSpace: 'pre-wrap' }}>
              {msg.content}
            </div>
          </div>
        ))}
      </div>
      <div style={{ borderTop: '1px solid var(--border)', padding: 10 }}>
        <textarea
          className="input"
          value={value}
          onChange={e => setValue(e.target.value)}
          onKeyDown={e => {
            if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); send() }
          }}
          placeholder={placeholder}
          style={{ width: '100%', minHeight: 78, resize: 'vertical', marginBottom: 8 }}
        />
        <button className="btn btn-primary btn-sm" onClick={send} disabled={busy || !value.trim()} style={{ width: '100%', justifyContent: 'center' }}>
          {busy ? busyLabel : 'Send to AI'}
        </button>
      </div>
    </aside>
  )
}
