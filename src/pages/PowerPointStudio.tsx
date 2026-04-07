import { useEffect, useState } from 'react'
import { useLocation } from 'react-router-dom'
import { useSelector } from 'react-redux'
import Sidebar from '../components/Sidebar'
import Header from '../components/Header'
import { RootState } from '../index'
import WorkspaceAssistant from '../components/WorkspaceAssistant'
import { api } from '../api'

interface Slide {
  id: number
  title: string
  content: string
  chart?: string
  notes?: string
}

interface DeckDraftState {
  prompt: string
  theme: string
  audience: string
  goal: string
}

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

export default function PowerPointStudio() {
  const location = useLocation()
  const user = useSelector((s: RootState) => s.app.user)
  const [sidebarOpen, setSidebarOpen] = useState(true)
  const [prompt, setPrompt] = useState('')
  const [generating, setGenerating] = useState(false)
  const [slides, setSlides] = useState<Slide[]>(() => {
    try { return JSON.parse(localStorage.getItem('ai_ppt_slides') || '[]') } catch { return [] }
  })
  const [theme, setTheme] = useState('Executive Blue')
  const [audience, setAudience] = useState('Management')
  const [goal, setGoal] = useState('Explain performance and recommendations')
  const [aiNotes, setAiNotes] = useState('')
  const [generationStage, setGenerationStage] = useState('Ready to storyboard a new presentation')
  const [presenting, setPresenting] = useState(false)
  const [currentSlide, setCurrentSlide] = useState(0)
  const [draftQuestions, setDraftQuestions] = useState<string[]>([])

  useEffect(() => {
    try {
      const saved = JSON.parse(localStorage.getItem('ai_ppt_draft_meta') || 'null') as DeckDraftState | null
      if (!saved) return
      setPrompt(prev => prev || saved.prompt || '')
      setTheme(saved.theme || 'Executive Blue')
      setAudience(saved.audience || 'Management')
      setGoal(saved.goal || 'Explain performance and recommendations')
    } catch {}
  }, [])

  useEffect(() => {
    try {
      const savedContext = localStorage.getItem('ai_ppt_context')
      if (!savedContext) return
      setAiNotes(prev => prev || savedContext)
    } catch {}
  }, [])

  useEffect(() => {
    localStorage.setItem('ai_ppt_slides', JSON.stringify(slides))
  }, [slides])

  useEffect(() => {
    localStorage.setItem('ai_ppt_draft_meta', JSON.stringify({ prompt, theme, audience, goal }))
  }, [prompt, theme, audience, goal])

  useEffect(() => {
    localStorage.setItem('ai_ppt_context', aiNotes || '')
  }, [aiNotes])

  useEffect(() => {
    const stateDraft = (location.state as any)?.aiDraft
    const storedDraft = sessionStorage.getItem('ai_ppt_draft')
    const draft = stateDraft || (storedDraft ? JSON.parse(storedDraft) : null)
    if (!draft) return

    const summary = [
      draft.prompt ? `Deck request: ${draft.prompt}` : '',
      draft.fileName ? `Source sheet: ${draft.fileName}` : '',
      draft.data?.rowCount ? `Rows available: ${draft.data.rowCount}` : '',
      draft.data?.headers?.length ? `Columns: ${draft.data.headers.join(', ')}` : '',
      draft.data?.sampleRows?.length ? `Sample rows:\n${draft.data.sampleRows.map((row: any[]) => row.map(cell => String(cell ?? '')).join(' | ')).join('\n')}` : '',
      draft.data?.stats?.avgScore !== undefined ? `Average: ${draft.data.stats.avgScore}` : '',
      draft.data?.stats?.passRate !== undefined ? `Pass rate: ${draft.data.stats.passRate}%` : ''
    ].filter(Boolean).join('\n')

    setPrompt(prev => prev || draft.prompt || 'Create a presentation from this spreadsheet data')
    setAiNotes(summary)
    try {
      const savedQuestions = JSON.parse(sessionStorage.getItem('ai_ppt_questions') || '[]')
      if (Array.isArray(savedQuestions)) setDraftQuestions(savedQuestions.filter(Boolean))
    } catch {}
    setGenerationStage('Workbook context loaded. Answer the planning questions, adjust the deck settings, then generate the presentation.')
    sessionStorage.removeItem('ai_ppt_draft')
    sessionStorage.removeItem('ai_ppt_questions')
  }, [location.state])

  const generateDeck = async () => {
    if (!prompt.trim()) return
    setGenerating(true)
    setGenerationStage('AI is shaping the storyline, chart ideas, and speaker notes')
    try {
      const result = await api.generatePpt({
        prompt,
        data: { notes: aiNotes, user: user?.name, department: user?.department },
        audience,
        theme,
        goal,
        department: user?.department || 'general'
      })
      const deck = (result.slides || []).map((slide: any, index: number) => ({
        id: index + 1,
        title: asText(slide.title || `Slide ${index + 1}`),
        content: asText(slide.content),
        chart: asText(slide.chart),
        notes: asText(slide.notes)
      }))
      setSlides(deck)
      setGenerationStage(`Deck preview ready with ${deck.length} slides. Review the story, then refine or export.`)
    } catch (e) {
      console.error('PPT generation failed:', e)
      setGenerationStage('AI could not generate the deck right now. Please retry after the backend is available.')
    }
    setGenerating(false)
  }

  const exportDeck = () => {
    const content = [
      `Title: ${prompt}`,
      `Theme: ${theme}`,
      `Audience: ${audience}`,
      `Goal: ${goal}`,
      '',
      ...slides.map(slide => `Slide ${slide.id}: ${slide.title}\n${slide.content}\nNotes: ${slide.notes || ''}\n`)
    ].join('\n')
    const blob = new Blob([content], { type: 'text/plain' })
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url
    a.download = `${prompt.slice(0, 40).replace(/[^a-zA-Z0-9]+/g, '_') || 'ai_presentation'}.txt`
    a.click()
    URL.revokeObjectURL(url)
  }

  const shareDeck = async () => {
    const payload = [
      `Presentation: ${prompt || 'AI Presentation'}`,
      `Theme: ${theme}`,
      `Audience: ${audience}`,
      `Goal: ${goal}`,
      '',
      ...slides.map(slide => `Slide ${slide.id}: ${slide.title}\n${slide.content}`)
    ].join('\n')
    try {
      if (navigator.share) {
        await navigator.share({ title: prompt || 'AI Presentation', text: payload })
      } else {
        await navigator.clipboard.writeText(payload)
      }
      setGenerationStage('Presentation share content is ready. It was shared or copied to the clipboard.')
    } catch (error) {
      console.error('Share failed:', error)
      setGenerationStage('Could not share the deck. Try export instead.')
    }
  }

  const openPresentation = () => {
    if (!slides.length) return
    setCurrentSlide(0)
    setPresenting(true)
    setGenerationStage('Presentation preview opened. Use Next and Previous to walk through the deck.')
  }

  return (
    <div style={{ display: 'flex', flexDirection: 'column', height: '100vh', background: 'var(--bg)' }}>
      <Header toggleSidebar={() => setSidebarOpen(p => !p)} />
      <div style={{ display: 'flex', flex: 1, overflow: 'hidden' }}>
        <Sidebar isOpen={sidebarOpen} />
        <main style={{ flex: 1, overflowY: 'auto', padding: '24px 28px', background: 'var(--bg-2)' }}>
          <div style={{ marginBottom: 24 }}>
            <h1 style={{ fontSize: '1.5rem', letterSpacing: -0.5 }}>PowerPoint Generator</h1>
            <p style={{ fontSize: '0.8rem', color: 'var(--text-secondary)' }}>AI-powered slide decks with charts, themes, questions, and preview before export</p>
          </div>

          {aiNotes && (
            <div style={{ padding: '10px 14px', background: 'rgba(59,130,246,0.08)', border: '1px solid rgba(59,130,246,0.2)', borderRadius: 10, fontSize: '0.78rem', color: 'var(--blue)', marginBottom: 16, whiteSpace: 'pre-wrap' }}>
              AI context loaded from Excel Sheet:
              {'\n'}
              {aiNotes}
            </div>
          )}

          <div style={{ display: 'grid', gridTemplateColumns: '1.1fr 0.9fr', gap: 14, marginBottom: 24 }}>
            <div style={{ padding: '18px 20px', borderRadius: 18, background: 'linear-gradient(135deg, color-mix(in srgb, var(--surface-3) 82%, var(--orange) 18%), color-mix(in srgb, var(--surface-3) 60%, var(--orange) 40%))', border: '1px solid color-mix(in srgb, var(--border-hi) 74%, var(--orange) 26%)', boxShadow: '0 12px 28px var(--accent-dim2)' }}>
              <div style={{ fontSize: '0.72rem', color: 'var(--orange)', textTransform: 'uppercase', letterSpacing: '0.08em', marginBottom: 6 }}>AI Presentation Studio</div>
              <div style={{ fontSize: '1.08rem', color: 'var(--text)', fontWeight: 700, marginBottom: 6 }}>Turn workbook data into a deck with a narrative</div>
              <div style={{ fontSize: '0.8rem', color: 'var(--text-secondary)', lineHeight: 1.65 }}>
                Give the AI a goal, audience, and theme. It will build a clean slide flow, suggest chart moments, and prepare speaker notes you can keep refining from the assistant panel.
              </div>
            </div>
            <div style={{ padding: '18px 20px', borderRadius: 18, background: 'var(--surface)', border: '1px solid var(--border)' }}>
              <div style={{ fontSize: '0.72rem', color: 'var(--text-muted)', textTransform: 'uppercase', letterSpacing: '0.08em', marginBottom: 6 }}>Generation Status</div>
              <div style={{ fontSize: '0.95rem', color: 'var(--text)', fontWeight: 700, marginBottom: 6 }}>{generating ? 'Designing with AI' : 'Ready'}</div>
              <div style={{ fontSize: '0.78rem', color: 'var(--text-secondary)', lineHeight: 1.6 }}>{generationStage}</div>
            </div>
          </div>

          <div style={{ background: 'var(--surface)', border: '1px solid var(--border)', borderRadius: 10, padding: '20px', marginBottom: 24 }}>
            {draftQuestions.length > 0 && (
              <div style={{ marginBottom: 14 }}>
                <div style={{ fontSize: '0.72rem', color: 'var(--text-muted)', textTransform: 'uppercase', letterSpacing: '0.08em', marginBottom: 8 }}>Questions Before Creating</div>
                <div style={{ display: 'flex', flexWrap: 'wrap', gap: 8, marginBottom: 10 }}>
                  {draftQuestions.map(question => (
                    <button key={question} className="btn btn-outline btn-sm" onClick={() => setPrompt(prev => prev ? `${prev}\n${question}` : question)}>
                      {question}
                    </button>
                  ))}
                </div>
                <div style={{ fontSize: '0.76rem', color: 'var(--text-secondary)', lineHeight: 1.6 }}>
                  Start by answering these in the prompt box, then generate the deck.
                </div>
              </div>
            )}
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, minmax(160px, 1fr))', gap: 10, marginBottom: 14 }}>
              <select className="input" value={theme} onChange={e => setTheme(e.target.value)} style={{ height: 40 }}>
                <option>Executive Blue</option>
                <option>Academic Clean</option>
                <option>Investor Dark</option>
                <option>Modern Orange</option>
              </select>
              <select className="input" value={audience} onChange={e => setAudience(e.target.value)} style={{ height: 40 }}>
                <option>Management</option>
                <option>Students</option>
                <option>Faculty</option>
                <option>Investors</option>
              </select>
              <input className="input" value={goal} onChange={e => setGoal(e.target.value)} placeholder="What should the deck achieve?" style={{ height: 40 }} />
            </div>
            <textarea
              className="input"
              value={prompt}
              onChange={e => setPrompt(e.target.value)}
              placeholder="Describe your presentation... e.g., Q4 investor update with revenue charts and growth metrics"
              style={{ width: '100%', minHeight: 100, marginBottom: 16 }}
            />
            <button className="btn btn-primary" onClick={generateDeck} disabled={generating}>
              {generating ? 'Generating...' : 'Generate Presentation'}
            </button>
          </div>

          {generating && (
            <div style={{ padding: '18px 20px', borderRadius: 16, background: 'linear-gradient(135deg, color-mix(in srgb, var(--surface-3) 84%, var(--orange) 16%), color-mix(in srgb, var(--surface-3) 62%, var(--orange) 38%))', border: '1px solid color-mix(in srgb, var(--border-hi) 78%, var(--orange) 22%)', marginBottom: 24, boxShadow: '0 12px 28px var(--accent-dim2)' }}>
              <div style={{ fontSize: '0.72rem', color: 'var(--orange)', textTransform: 'uppercase', letterSpacing: '0.08em', marginBottom: 8 }}>AI Processing</div>
              <div style={{ fontSize: '1rem', color: 'var(--text)', fontWeight: 700, marginBottom: 6 }}>Building the deck from your workbook context</div>
              <div style={{ fontSize: '0.82rem', color: 'var(--text-secondary)', lineHeight: 1.6 }}>
                The AI is reading the Excel notes, drafting slide structure, and preparing chart moments and speaker notes.
              </div>
            </div>
          )}

          {slides.length > 0 && (
            <div>
              <div style={{ display: 'grid', gridTemplateColumns: '1.2fr 0.8fr', gap: 14, marginBottom: 18 }}>
                <div style={{ padding: '18px 20px', borderRadius: 18, background: 'linear-gradient(145deg, var(--surface-3), color-mix(in srgb, var(--surface-3) 78%, var(--orange) 22%))', border: '1px solid var(--border-hi)', boxShadow: '0 12px 28px var(--accent-dim2)' }}>
                  <div style={{ fontSize: '0.7rem', color: 'var(--orange)', textTransform: 'uppercase', letterSpacing: '0.08em', marginBottom: 6 }}>Deck Summary</div>
                  <div style={{ fontSize: '1.25rem', color: 'var(--text)', fontWeight: 800, marginBottom: 8 }}>{prompt || 'AI Presentation'}</div>
                  <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
                    {[`Theme: ${theme}`, `Audience: ${audience}`, `Goal: ${goal}`, `${slides.length} slides`].map(item => (
                      <span key={item} style={{ padding: '5px 9px', borderRadius: 999, background: 'color-mix(in srgb, var(--surface) 76%, var(--orange) 24%)', color: 'var(--text)', fontSize: '0.68rem', border: '1px solid color-mix(in srgb, var(--border-hi) 72%, var(--orange) 28%)' }}>{item}</span>
                    ))}
                  </div>
                </div>
                <div style={{ padding: '18px 20px', borderRadius: 18, background: 'var(--surface)', border: '1px solid var(--border)' }}>
                  <div style={{ fontSize: '0.72rem', color: 'var(--text-muted)', textTransform: 'uppercase', letterSpacing: '0.08em', marginBottom: 8 }}>Workbook Data Used</div>
                  <div style={{ fontSize: '0.78rem', color: 'var(--text-secondary)', lineHeight: 1.65, whiteSpace: 'pre-wrap', maxHeight: 132, overflowY: 'auto' }}>
                    {aiNotes || 'No workbook context was available for this deck.'}
                  </div>
                </div>
              </div>
              <h3 style={{ marginBottom: 16 }}>Generated Slides ({slides.length})</h3>
              <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
                {slides.map(slide => (
                  <div key={slide.id} style={{ background: 'linear-gradient(180deg, rgba(255,255,255,0.02), rgba(255,255,255,0))', border: '1px solid var(--border)', borderRadius: 16, padding: '16px 20px', position: 'relative', overflow: 'hidden' }}>
                    <div style={{ position: 'absolute', top: 0, right: 0, width: 140, height: 140, background: 'radial-gradient(circle at top right, rgba(59,130,246,0.16), transparent 60%)', pointerEvents: 'none' }} />
                    <div style={{ fontSize: '0.7rem', color: 'var(--text-muted)', textTransform: 'uppercase', letterSpacing: '0.08em', marginBottom: 6 }}>Slide {slide.id}</div>
                    <div style={{ fontSize: '1rem', fontWeight: 600, marginBottom: 8 }}>{slide.title}</div>
                    <p style={{ fontSize: '0.85rem', color: 'var(--text-secondary)', marginBottom: 8, whiteSpace: 'pre-wrap' }}>{slide.content}</p>
                    {slide.chart && (
                      <div style={{ height: 100, background: 'linear-gradient(135deg, rgba(59,130,246,0.08), rgba(249,115,22,0.08))', borderRadius: 10, display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: '0.7rem', color: 'var(--text-muted)', border: '1px solid var(--border)' }}>
                        Chart moment: {slide.chart}
                      </div>
                    )}
                    {slide.notes && <div style={{ marginTop: 8, fontSize: '0.72rem', color: 'var(--text-muted)' }}>Speaker notes: {slide.notes}</div>}
                  </div>
                ))}
              </div>
              <div style={{ display: 'flex', gap: 12, marginTop: 24, flexWrap: 'wrap' }}>
                <button className="btn btn-secondary" onClick={openPresentation}>Present Preview</button>
                <button className="btn btn-outline" onClick={shareDeck}>Share / Copy</button>
                <button className="btn btn-primary" onClick={exportDeck}>Export Deck Outline</button>
                <button className="btn btn-secondary" onClick={() => setSlides([])}>Clear Preview</button>
              </div>
            </div>
          )}
        </main>
        <WorkspaceAssistant
          title="Xtron"
          subtitle="Ask for stronger structure, different audiences, more charts, or sharper slide writing."
          placeholder="Ask AI to regenerate slides, change tone, add sections, or rewrite for a new audience..."
          initialMessage="I can help turn your workbook data into a better slide deck, one refinement at a time."
          suggestions={[
            'Add a slide with recommendations',
            'Make it suitable for management',
            'Add stronger chart moments',
            'Shorten the deck to 5 slides'
          ]}
          busyLabel="Designing slides..."
          onSubmit={async (text) => {
            const result = await api.generatePpt({
              prompt: `${prompt}\n\nRefinement request: ${text}`,
              data: { notes: aiNotes, user: user?.name, department: user?.department },
              audience,
              theme,
              goal,
              department: user?.department || 'general'
            })
            const deck = (result.slides || []).map((slide: any, index: number) => ({
              id: index + 1,
              title: asText(slide.title || `Slide ${index + 1}`),
              content: asText(slide.content),
              chart: asText(slide.chart),
              notes: asText(slide.notes)
            }))
            setSlides(deck)
            return `Updated the presentation with ${deck.length} slides for ${audience}.`
          }}
        />
      </div>

      {presenting && slides[currentSlide] && (
        <div className="modal-backdrop" onClick={() => setPresenting(false)}>
          <div className="modal-box" onClick={e => e.stopPropagation()} style={{ width: 860, maxWidth: '92vw', maxHeight: '88vh', background: 'linear-gradient(145deg, #0f172a, #111827)', color: '#f8fafc', border: '1px solid rgba(148,163,184,0.18)', padding: 0, overflow: 'hidden' }}>
            <div style={{ padding: '18px 22px', borderBottom: '1px solid rgba(148,163,184,0.16)', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
              <div>
                <div style={{ fontSize: '0.72rem', color: '#93c5fd', textTransform: 'uppercase', letterSpacing: '0.08em' }}>Presentation Preview</div>
                <div style={{ fontSize: '1.05rem', fontWeight: 700, marginTop: 4 }}>{prompt || 'AI Presentation'}</div>
              </div>
              <button className="btn btn-secondary" onClick={() => setPresenting(false)}>Close</button>
            </div>
            <div style={{ padding: '28px 32px', minHeight: 420, display: 'flex', flexDirection: 'column', justifyContent: 'space-between', gap: 18 }}>
              <div>
                <div style={{ fontSize: '0.75rem', color: '#94a3b8', marginBottom: 10 }}>Slide {currentSlide + 1} of {slides.length}</div>
                <div style={{ fontSize: '1.9rem', fontWeight: 800, marginBottom: 16 }}>{slides[currentSlide].title}</div>
                <div style={{ fontSize: '1rem', lineHeight: 1.8, color: '#e2e8f0', whiteSpace: 'pre-wrap' }}>{slides[currentSlide].content}</div>
              </div>
              {slides[currentSlide].chart && (
                <div style={{ padding: '14px 16px', borderRadius: 14, border: '1px solid rgba(148,163,184,0.16)', background: 'rgba(30,41,59,0.72)', color: '#cbd5e1' }}>
                  Chart moment: {slides[currentSlide].chart}
                </div>
              )}
              {slides[currentSlide].notes && (
                <div style={{ fontSize: '0.82rem', color: '#94a3b8', borderTop: '1px solid rgba(148,163,184,0.16)', paddingTop: 14 }}>
                  Speaker notes: {slides[currentSlide].notes}
                </div>
              )}
            </div>
            <div style={{ padding: '16px 22px', borderTop: '1px solid rgba(148,163,184,0.16)', display: 'flex', justifyContent: 'space-between' }}>
              <button className="btn btn-secondary" onClick={() => setCurrentSlide(prev => Math.max(prev - 1, 0))} disabled={currentSlide === 0}>Previous</button>
              <button className="btn btn-primary" onClick={() => setCurrentSlide(prev => Math.min(prev + 1, slides.length - 1))} disabled={currentSlide === slides.length - 1}>Next</button>
            </div>
          </div>
        </div>
      )}
    </div>
  )
}
