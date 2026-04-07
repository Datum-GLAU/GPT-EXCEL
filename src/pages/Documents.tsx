import { useState, useEffect } from 'react'
import { useLocation } from 'react-router-dom'
import { useSelector } from 'react-redux'
import { RootState } from '../index'
import Sidebar from '../components/Sidebar'
import Header from '../components/Header'
import WorkspaceAssistant from '../components/WorkspaceAssistant'
import { api } from '../api'

interface Doc { id: string; title: string; type: string; content: string; createdAt: string }

const asText = (value: any): string => {
  if (typeof value === 'string') return value
  if (value == null) return ''
  if (typeof value === 'number' || typeof value === 'boolean') return String(value)
  if (Array.isArray(value)) return value.map(asText).filter(Boolean).join('\n')
  if (typeof value === 'object') {
    if (typeof value.text === 'string') return value.text
    if (typeof value.content === 'string') return value.content
    try { return JSON.stringify(value, null, 2) } catch { return String(value) }
  }
  return String(value)
}

const TEMPLATES = [
  { type: 'report', icon: '📊', label: 'Student Performance Report', prompt: 'Generate a comprehensive student performance report with pass/fail analysis' },
  { type: 'proposal', icon: '📝', label: 'Academic Proposal', prompt: 'Write a formal academic proposal for curriculum improvement' },
  { type: 'notice', icon: '📋', label: 'Exam Notice', prompt: 'Create an official exam notification for students' },
  { type: 'minutes', icon: '📄', label: 'Meeting Minutes', prompt: 'Generate faculty meeting minutes template' },
  { type: 'circular', icon: '🔔', label: 'Department Circular', prompt: 'Write a departmental circular for faculty' },
  { type: 'certificate', icon: '🏆', label: 'Certificate Template', prompt: 'Create a student achievement certificate' },
]

export default function Documents() {
  const location = useLocation()
  const user = useSelector((s: RootState) => s.app.user)
  const userDepartment = (user as any)?.department as string | undefined
  const [sidebarOpen, setSidebarOpen] = useState(true)
  const [docs, setDocs] = useState<Doc[]>(() => {
    try { return JSON.parse(localStorage.getItem('ai_documents') || '[]') } catch { return [] }
  })
  const [selectedDoc, setSelectedDoc] = useState<Doc | null>(null)
  const [openedDoc, setOpenedDoc] = useState<Doc | null>(null)
  const [prompt, setPrompt] = useState('')
  const [docType, setDocType] = useState('report')
  const [generating, setGenerating] = useState(false)
  const [error, setError] = useState('')
  const [aiNotes, setAiNotes] = useState('')
  const [pendingAutoPrompt, setPendingAutoPrompt] = useState('')
  const [generationStage, setGenerationStage] = useState('Ready for a new document')
  const [previewMode, setPreviewMode] = useState<'read' | 'present'>('read')

  useEffect(() => {
    localStorage.setItem('ai_documents', JSON.stringify(docs))
  }, [docs])

  useEffect(() => {
    const lastId = localStorage.getItem('ai_documents_last_open')
    if (!docs.length) return
    if (selectedDoc && docs.some(doc => doc.id === selectedDoc.id)) return
    const nextDoc = docs.find(doc => doc.id === lastId) || docs[0]
    setSelectedDoc(nextDoc)
  }, [docs, selectedDoc])

  useEffect(() => {
    if (!openedDoc) return
    if (docs.some(doc => doc.id === openedDoc.id)) return
    setOpenedDoc(null)
  }, [docs, openedDoc])

  useEffect(() => {
    if (selectedDoc) {
      localStorage.setItem('ai_documents_last_open', selectedDoc.id)
    }
  }, [selectedDoc])

  useEffect(() => {
    setPreviewMode('read')
  }, [openedDoc?.id])

  useEffect(() => {
    const stateDraft = (location.state as any)?.aiDraft
    const storedDraft = sessionStorage.getItem('ai_document_draft')
    const draft = stateDraft || (storedDraft ? JSON.parse(storedDraft) : null)
    if (!draft) return

    const summary = [
      draft.prompt ? `Request: ${draft.prompt}` : '',
      draft.fileName ? `Source sheet: ${draft.fileName}` : '',
      draft.data?.rowCount ? `Rows: ${draft.data.rowCount}` : '',
      draft.data?.headers?.length ? `Columns: ${draft.data.headers.join(', ')}` : '',
      draft.data?.sampleRows?.length ? `Sample rows:\n${draft.data.sampleRows.map((row: any[]) => row.map(cell => String(cell ?? '')).join(' | ')).join('\n')}` : '',
      draft.data?.stats?.passRate !== undefined ? `Pass Rate: ${draft.data.stats.passRate}%` : ''
    ].filter(Boolean).join('\n')

    setPrompt(prev => prev || `Create a professional ${docType} using this spreadsheet data.\n\n${summary}`)
    setAiNotes(summary)
    setPendingAutoPrompt(`Create a professional ${docType} using this spreadsheet data.\n\n${summary}`)
    sessionStorage.removeItem('ai_document_draft')
  }, [location.state, docType])

  useEffect(() => {
    if (!pendingAutoPrompt || generating) return
    generate(pendingAutoPrompt, docType)
    setPendingAutoPrompt('')
  }, [pendingAutoPrompt, generating, docType])

  const generate = async (customPrompt?: string, customType?: string) => {
    const p = customPrompt || prompt
    if (!p.trim()) return
    setGenerating(true); setError('')
    setGenerationStage('AI is reading the sheet context and drafting a structured document')
    try {
      const doc = await api.generateDoc(customType || docType, { prompt: p, notes: aiNotes, user: user?.name, department: userDepartment }, p.slice(0, 60))
      const normalizedDoc = { ...doc, title: asText(doc.title || p.slice(0, 50)), content: asText(doc.content) }
      setDocs(prev => [normalizedDoc, ...prev])
      setSelectedDoc(normalizedDoc)
      setPrompt('')
      setGenerationStage('Document preview is ready to review, refine, or download')
    } catch (e: any) {
      setError('Backend offline — run: cd server && npm start')
      // fallback: local mock
      const mock: Doc = {
        id: `doc${Date.now()}`, title: p.slice(0, 50), type: customType || docType,
        content: `${p.slice(0,60)}\n\nGenerated: ${new Date().toLocaleString()}\nBy: ${user?.name || 'User'}\nDepartment: ${userDepartment || 'N/A'}\n\n[Backend offline — connect server for full AI generation]`,
        createdAt: new Date().toISOString()
      }
      setDocs(prev => [mock, ...prev])
      setPrompt('')
      setGenerationStage('Created a local fallback draft because the backend was unavailable')
    }
    setGenerating(false)
  }

  const downloadDoc = (doc: Doc) => {
    const blob = new Blob([doc.content], { type: 'text/plain' })
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a'); a.href = url; a.download = `${doc.title}.txt`; a.click()
    URL.revokeObjectURL(url)
  }

  const shareDoc = async (doc: Doc) => {
    const payload = `${doc.title}\n${doc.type}\n\n${doc.content}`
    try {
      if (navigator.share) {
        await navigator.share({ title: doc.title, text: payload })
      } else {
        await navigator.clipboard.writeText(payload)
      }
      setGenerationStage('Document share content is ready. It was shared or copied to the clipboard.')
    } catch (error) {
      console.error('Share failed:', error)
      setGenerationStage('Could not share the document right now. Try download instead.')
    }
  }

  const TYPE_COLORS: Record<string, string> = { report: 'var(--blue)', proposal: 'var(--purple)', notice: 'var(--orange)', minutes: 'var(--teal)', circular: 'var(--green)', certificate: 'var(--yellow)' }
  const TYPE_ICONS: Record<string, string> = { report: '📊', proposal: '📝', notice: '📋', minutes: '📄', circular: '🔔', certificate: '🏆' }

  return (
    <div style={{ display: 'flex', flexDirection: 'column', height: '100vh', background: 'var(--bg)' }}>
      <Header toggleSidebar={() => setSidebarOpen(p => !p)} />
      <div style={{ display: 'flex', flex: 1, overflow: 'hidden' }}>
        <Sidebar isOpen={sidebarOpen} />
        <main style={{ flex: 1, overflowY: 'auto', padding: '20px 24px', background: 'var(--bg-2)' }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: 20 }}>
            <div>
              <h1 style={{ fontSize: '1.3rem', letterSpacing: -0.5, marginBottom: 4 }}>Documents</h1>
              <p style={{ fontSize: '0.78rem', color: 'var(--text-secondary)' }}>Generate reports, notices, proposals, and academic documents</p>
            </div>
          </div>

          {error && <div style={{ padding: '8px 14px', background: 'rgba(249,115,22,0.1)', border: '1px solid rgba(249,115,22,0.3)', borderRadius: 8, fontSize: '0.78rem', color: 'var(--orange)', marginBottom: 16 }}>{error}</div>}

          {aiNotes && (
            <div style={{ padding: '10px 14px', background: 'rgba(59,130,246,0.08)', border: '1px solid rgba(59,130,246,0.2)', borderRadius: 10, fontSize: '0.78rem', color: 'var(--blue)', marginBottom: 16, whiteSpace: 'pre-wrap' }}>
              AI context loaded from Excel Sheet:
              {'\n'}
              {aiNotes}
            </div>
          )}

          <div style={{ display: 'grid', gridTemplateColumns: '1.2fr 0.8fr', gap: 14, marginBottom: 20 }}>
            <div style={{ padding: '16px 18px', borderRadius: 16, background: 'linear-gradient(135deg, color-mix(in srgb, var(--surface-3) 82%, var(--blue) 18%), color-mix(in srgb, var(--surface-3) 60%, var(--blue) 40%))', border: '1px solid color-mix(in srgb, var(--border-hi) 74%, var(--blue) 26%)', boxShadow: '0 12px 28px var(--accent-dim2)' }}>
              <div style={{ fontSize: '0.72rem', color: 'var(--blue)', textTransform: 'uppercase', letterSpacing: '0.08em', marginBottom: 6 }}>AI Document Studio</div>
              <div style={{ fontSize: '1.05rem', color: 'var(--text)', fontWeight: 700, marginBottom: 6 }}>Structured writing from workbook data</div>
              <div style={{ fontSize: '0.8rem', color: 'var(--text-secondary)', lineHeight: 1.65 }}>
                Ask for reports, notices, proposals, summaries, or policy-style documents. The AI uses sheet context, department context, and your prompt to draft something you can refine in chat.
              </div>
            </div>
            <div style={{ padding: '16px 18px', borderRadius: 16, background: 'var(--surface)', border: '1px solid var(--border)' }}>
              <div style={{ fontSize: '0.72rem', color: 'var(--text-muted)', textTransform: 'uppercase', letterSpacing: '0.08em', marginBottom: 6 }}>Generation Status</div>
              <div style={{ fontSize: '0.95rem', color: 'var(--text)', fontWeight: 700, marginBottom: 6 }}>{generating ? 'Processing with AI' : 'Ready'}</div>
              <div style={{ fontSize: '0.78rem', color: 'var(--text-secondary)', lineHeight: 1.6 }}>{generationStage}</div>
            </div>
          </div>

          {/* Generate box */}
          <div style={{ background: 'var(--surface)', border: '1px solid var(--border)', borderRadius: 10, padding: 16, marginBottom: 20 }}>
            <div style={{ display: 'flex', gap: 8, marginBottom: 10 }}>
              <select className="input" value={docType} onChange={e => setDocType(e.target.value)} style={{ width: 160, height: 34, padding: '0 10px', fontSize: '0.8rem' }}>
                {TEMPLATES.map(t => <option key={t.type} value={t.type}>{t.label}</option>)}
              </select>
              <span style={{ fontSize: '0.72rem', color: 'var(--text-muted)', alignSelf: 'center' }}>for {userDepartment || 'your department'}</span>
            </div>
            <div style={{ display: 'flex', gap: 10 }}>
              <textarea className="input" value={prompt} onChange={e => setPrompt(e.target.value)} placeholder="Describe the document you need... e.g., 'Q4 result analysis report for CSE section A with 45 students, pass rate 78%'" style={{ flex: 1, minHeight: 80, resize: 'vertical', fontSize: '0.85rem' }} />
              <button className="btn btn-primary" onClick={() => generate()} disabled={generating || !prompt.trim()} style={{ alignSelf: 'flex-end', minWidth: 100 }}>
                {generating ? <><span className="spinner" style={{ width: 12, height: 12 }} /> Generating</> : 'Generate →'}
              </button>
            </div>
          </div>

          {generating && (
            <div style={{ padding: '18px 20px', borderRadius: 16, background: 'linear-gradient(135deg, color-mix(in srgb, var(--surface-3) 84%, var(--blue) 16%), color-mix(in srgb, var(--surface-3) 62%, var(--blue) 38%))', border: '1px solid color-mix(in srgb, var(--border-hi) 78%, var(--blue) 22%)', marginBottom: 20, boxShadow: '0 12px 28px var(--accent-dim2)' }}>
              <div style={{ fontSize: '0.72rem', color: 'var(--blue)', textTransform: 'uppercase', letterSpacing: '0.08em', marginBottom: 8 }}>AI Processing</div>
              <div style={{ fontSize: '1rem', color: 'var(--text)', fontWeight: 700, marginBottom: 6 }}>Drafting from the Excel sheet context</div>
              <div style={{ fontSize: '0.82rem', color: 'var(--text-secondary)', lineHeight: 1.65 }}>
                The AI is using workbook details, the selected document type, and your prompt to build a structured draft.
              </div>
            </div>
          )}

          {selectedDoc && (
            <div style={{ display: 'grid', gridTemplateColumns: '1.15fr 0.85fr', gap: 14, marginBottom: 20 }}>
              <div style={{ padding: '18px 20px', borderRadius: 18, background: 'linear-gradient(145deg, rgba(30,41,59,0.98), rgba(15,23,42,0.95))', border: '1px solid rgba(148,163,184,0.16)' }}>
                <div style={{ fontSize: '0.7rem', color: '#93c5fd', textTransform: 'uppercase', letterSpacing: '0.08em', marginBottom: 6 }}>Current Draft</div>
                <div style={{ fontSize: '1.2rem', color: '#f8fafc', fontWeight: 800, marginBottom: 8 }}>{selectedDoc.title}</div>
                <div style={{ fontSize: '0.78rem', color: '#dbeafe', lineHeight: 1.7, whiteSpace: 'pre-wrap', maxHeight: 220, overflowY: 'auto' }}>
                  {selectedDoc.content}
                </div>
                <div style={{ display: 'flex', gap: 8, marginTop: 14, flexWrap: 'wrap' }}>
                  <button className="btn btn-sm btn-primary" onClick={() => setOpenedDoc(selectedDoc)}>Open Reader</button>
                  <button className="btn btn-sm btn-outline" onClick={() => shareDoc(selectedDoc)}>Share</button>
                  <button className="btn btn-sm btn-outline" onClick={() => downloadDoc(selectedDoc)}>Download .txt</button>
                </div>
              </div>
              <div style={{ padding: '18px 20px', borderRadius: 18, background: 'var(--surface)', border: '1px solid var(--border)' }}>
                <div style={{ fontSize: '0.72rem', color: 'var(--text-muted)', textTransform: 'uppercase', letterSpacing: '0.08em', marginBottom: 8 }}>Workbook Data Used</div>
                <div style={{ fontSize: '0.78rem', color: 'var(--text-secondary)', lineHeight: 1.65, whiteSpace: 'pre-wrap', maxHeight: 220, overflowY: 'auto' }}>
                  {aiNotes || 'No workbook context was available for this document.'}
                </div>
              </div>
            </div>
          )}

          {/* Templates */}
          <div style={{ marginBottom: 20 }}>
            <div style={{ fontSize: '0.72rem', fontFamily: 'var(--font-mono)', color: 'var(--text-muted)', textTransform: 'uppercase', letterSpacing: '0.08em', marginBottom: 10 }}>Quick Templates</div>
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(200px, 1fr))', gap: 8 }}>
              {TEMPLATES.map(t => (
                <button key={t.type} onClick={() => generate(t.prompt, t.type)} disabled={generating} style={{ display: 'flex', alignItems: 'center', gap: 10, padding: '10px 12px', background: 'var(--surface)', border: '1px solid var(--border)', borderRadius: 8, cursor: 'pointer', transition: 'all 0.15s', textAlign: 'left' }}
                  onMouseEnter={e => { const el = e.currentTarget as HTMLElement; el.style.borderColor = TYPE_COLORS[t.type] || 'var(--accent)'; el.style.background = 'var(--surface-2)' }}
                  onMouseLeave={e => { const el = e.currentTarget as HTMLElement; el.style.borderColor = 'var(--border)'; el.style.background = 'var(--surface)' }}
                >
                  <span style={{ fontSize: 18 }}>{t.icon}</span>
                  <span style={{ fontSize: '0.78rem', color: 'var(--text)', fontWeight: 500 }}>{t.label}</span>
                </button>
              ))}
            </div>
          </div>

          {/* Docs list */}
          {docs.length > 0 && (
            <div>
              <div style={{ fontSize: '0.72rem', fontFamily: 'var(--font-mono)', color: 'var(--text-muted)', textTransform: 'uppercase', letterSpacing: '0.08em', marginBottom: 10 }}>Generated ({docs.length})</div>
              <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(280px, 1fr))', gap: 12 }}>
                {docs.map(doc => (
                  <div key={doc.id} style={{ background: 'var(--surface)', border: '1px solid var(--border)', borderRadius: 10, padding: 16, cursor: 'pointer', transition: 'all var(--tr)' }}
                    onMouseEnter={e => { (e.currentTarget as HTMLElement).style.borderColor = TYPE_COLORS[doc.type] || 'var(--accent)' }}
                    onMouseLeave={e => { (e.currentTarget as HTMLElement).style.borderColor = 'var(--border)' }}
                    onClick={() => { setSelectedDoc(doc); setOpenedDoc(doc) }}
                  >
                    <div style={{ display: 'flex', alignItems: 'flex-start', gap: 10, marginBottom: 8 }}>
                      <span style={{ fontSize: 22, flexShrink: 0 }}>{TYPE_ICONS[doc.type] || '📄'}</span>
                      <div style={{ flex: 1, minWidth: 0 }}>
                        <div style={{ fontSize: '0.85rem', fontWeight: 600, color: 'var(--text)', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{doc.title}</div>
                        <div style={{ fontSize: '0.68rem', color: 'var(--text-muted)', marginTop: 2 }}>{doc.type} · {new Date(doc.createdAt).toLocaleString()}</div>
                      </div>
                    </div>
                    <p style={{ fontSize: '0.75rem', color: 'var(--text-muted)', overflow: 'hidden', textOverflow: 'ellipsis', display: '-webkit-box', WebkitLineClamp: 3, WebkitBoxOrient: 'vertical' as any }}>{doc.content.slice(0, 120)}...</p>
                    <div style={{ display: 'flex', gap: 6, marginTop: 10 }}>
                      <button className="btn btn-xs btn-outline" onClick={e => { e.stopPropagation(); downloadDoc(doc) }}>↓ Download</button>
                      <button className="btn btn-xs btn-outline" onClick={e => { e.stopPropagation(); shareDoc(doc) }}>Share</button>
                      <button className="btn btn-xs btn-ghost" onClick={e => { e.stopPropagation(); setDocs(prev => prev.filter(d => d.id !== doc.id)) }} style={{ color: 'var(--red)' }}>✕</button>
                    </div>
                  </div>
                ))}
              </div>
            </div>
          )}

          {/* Modal */}
          {openedDoc && (
            <div className="modal-backdrop" onClick={() => setOpenedDoc(null)}>
              <div className="modal-box" style={{ width: 620, maxHeight: '80vh', display: 'flex', flexDirection: 'column' }} onClick={e => e.stopPropagation()}>
                <div style={{ display: 'flex', alignItems: 'center', gap: 10, marginBottom: 16 }}>
                  <span style={{ fontSize: 24 }}>{TYPE_ICONS[openedDoc.type] || '📄'}</span>
                  <div>
                    <div className="modal-title" style={{ margin: 0 }}>{openedDoc.title}</div>
                    <div style={{ fontSize: '0.7rem', color: 'var(--text-muted)' }}>{openedDoc.type} · {new Date(openedDoc.createdAt).toLocaleString()}</div>
                  </div>
                </div>
                <div style={{ display: 'flex', gap: 8, marginBottom: 12 }}>
                  <button className={`btn btn-sm ${previewMode === 'read' ? 'btn-primary' : 'btn-outline'}`} onClick={() => setPreviewMode('read')}>Read</button>
                  <button className={`btn btn-sm ${previewMode === 'present' ? 'btn-primary' : 'btn-outline'}`} onClick={() => setPreviewMode('present')}>Present</button>
                </div>
                <div style={{ flex: 1, overflowY: 'auto', background: previewMode === 'present' ? 'linear-gradient(145deg, color-mix(in srgb, var(--surface-3) 84%, var(--blue) 16%), color-mix(in srgb, var(--surface-3) 60%, var(--blue) 40%))' : 'var(--surface-2)', borderRadius: 8, padding: previewMode === 'present' ? 24 : 16, fontFamily: previewMode === 'present' ? 'var(--font-body)' : 'var(--font-mono)', fontSize: previewMode === 'present' ? '0.98rem' : '0.78rem', lineHeight: previewMode === 'present' ? 1.9 : 1.7, color: previewMode === 'present' ? 'var(--text)' : 'var(--text-secondary)', whiteSpace: 'pre-wrap', maxHeight: 400, border: previewMode === 'present' ? '1px solid var(--border-hi)' : 'none' }}>
                  {openedDoc.content}
                </div>
                <div style={{ display: 'flex', gap: 8, justifyContent: 'flex-end', marginTop: 16 }}>
                  <button className="btn btn-secondary" onClick={() => setOpenedDoc(null)}>Close</button>
                  <button className="btn btn-outline" onClick={() => shareDoc(openedDoc)}>Share</button>
                  <button className="btn btn-primary" onClick={() => downloadDoc(openedDoc)}>↓ Download .txt</button>
                </div>
              </div>
            </div>
          )}
        </main>
        <WorkspaceAssistant
          title="Xtron"
          subtitle="Refine drafts, regenerate, or create formal documents from workbook context."
          placeholder="Ask AI to rewrite, expand, formalize, shorten, or create another document..."
          initialMessage="I can help you create polished reports, notices, proposals, and formal documents from your data."
          suggestions={[
            'Turn this into a formal report',
            'Make it shorter and more executive',
            'Add a conclusion and action items',
            'Rewrite this as a notice'
          ]}
          busyLabel="Drafting with AI..."
          onSubmit={async (text) => {
            const refinedPrompt = selectedDoc
              ? `${text}\n\nCurrent document title: ${selectedDoc.title}\nCurrent content:\n${selectedDoc.content}`
              : text
            const doc = await api.generateDoc(docType, { prompt: refinedPrompt, notes: aiNotes, user: user?.name, department: userDepartment }, refinedPrompt.slice(0, 60))
            const normalizedDoc = { ...doc, title: asText(doc.title || refinedPrompt.slice(0, 50)), content: asText(doc.content) }
            setDocs(prev => [normalizedDoc, ...prev])
            setSelectedDoc(normalizedDoc)
            return `Created "${normalizedDoc.title}" and opened it in the preview.`
          }}
        />
      </div>
    </div>
  )
}
