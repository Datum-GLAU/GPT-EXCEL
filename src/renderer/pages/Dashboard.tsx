
import { useState, useEffect, useRef } from 'react'
import { useNavigate } from 'react-router-dom'
import Sidebar from '../components/Sidebar'
import Header from '../components/Header'
import AIChatPanel from '../components/AIChatPanel'

interface File { id: string; name: string; type: 'excel'|'doc'|'pdf'|'csv'|'pptx'; modified: string; size: string; starred: boolean; color: string }
interface Activity { id: string; icon: string; text: string; time: string; type: 'success'|'info'|'warning' }
interface KPI { label: string; value: string; change: string; up: boolean; sparkline: number[]; color: string }

const FILES: File[] = [
  { id:'1', name:'Q4 Financial Model.xlsx', type:'excel', modified:'2 min ago', size:'1.2 MB', starred:true, color:'var(--green)' },
  { id:'2', name:'Sales Dashboard 2026.xlsx', type:'excel', modified:'1 hr ago', size:'840 KB', starred:false, color:'var(--green)' },
  { id:'3', name:'Project Proposal.docx', type:'doc', modified:'3 hr ago', size:'320 KB', starred:true, color:'var(--blue)' },
  { id:'4', name:'Customer Data Export.csv', type:'csv', modified:'Yesterday', size:'4.1 MB', starred:false, color:'var(--yellow)' },
  { id:'5', name:'Annual Report 2025.pdf', type:'pdf', modified:'2 days ago', size:'2.8 MB', starred:false, color:'var(--red)' },
  { id:'6', name:'Q4 Pitch Deck.pptx', type:'pptx', modified:'3 days ago', size:'5.2 MB', starred:true, color:'var(--orange)' },
]

const ACTIVITIES: Activity[] = [
  { id:'1', icon:'⊞', text:'Generated "Q4 Revenue Forecast" — 12 sheets, 5 charts', time:'2m ago', type:'success' },
  { id:'2', icon:'◱', text:'Document "Project Proposal" exported as PDF', time:'1h ago', type:'info' },
  { id:'3', icon:'◎', text:'Voice command: "Create pivot table for sales data"', time:'2h ago', type:'success' },
  { id:'4', icon:'◈', text:'43 files categorized and tagged automatically', time:'4h ago', type:'success' },
  { id:'5', icon:'⌘', text:'Workflow "Weekly Report" scheduled for Monday 09:00', time:'Yesterday', type:'info' },
  { id:'6', icon:'◬', text:'API rate limit warning — 89% of free tier used', time:'Yesterday', type:'warning' },
]

const KPIS: KPI[] = [
  { label:'Files Generated', value:'248', change:'+12 today', up:true, sparkline:[4,6,5,8,7,9,11,10,12,14,13,16], color:'var(--blue)' },
  { label:'Tokens Used', value:'1.2M', change:'+84k today', up:true, sparkline:[20,22,18,25,30,28,35,33,40,38,42,45], color:'var(--purple)' },
  { label:'Storage Used', value:'3.4 GB', change:'+120 MB', up:true, sparkline:[10,12,13,14,15,17,19,20,22,24,25,27], color:'var(--orange)' },
  { label:'Free Tier Left', value:'23%', change:'27 remaining', up:false, sparkline:[80,74,68,62,55,49,43,38,34,30,27,23], color:'var(--yellow)' },
]

const Sparkline = ({ data, color }: { data: number[]; color: string }) => {
  const max = Math.max(...data), min = Math.min(...data)
  const norm = (v: number) => 1 - (v - min) / (max - min || 1)
  const w = 60, h = 22
  const pts = data.map((v, i) => `${(i / (data.length - 1)) * w},${norm(v) * h}`).join(' ')
  return (
    <svg width={w} height={h} style={{ overflow: 'visible' }}>
      <polyline points={pts} fill="none" stroke={color} strokeWidth={1.5} strokeLinejoin="round" strokeLinecap="round"/>
    </svg>
  )
}

const MiniBar = ({ data, colors }: { data: number[]; colors: string[] }) => {
  const max = Math.max(...data)
  const w = 120, h = 36, barW = (w - (data.length - 1) * 2) / data.length
  return (
    <svg width={w} height={h}>
      {data.map((v, i) => {
        const bh = (v / max) * h
        return <rect key={i} x={i * (barW + 2)} y={h - bh} width={barW} height={bh} fill={colors[i % colors.length]} opacity={0.85} rx={2}/>
      })}
    </svg>
  )
}

const MiniDonut = ({ vals, colors }: { vals: number[]; colors: string[] }) => {
  const total = vals.reduce((a, b) => a + b, 0)
  const R = 22, r = 14
  let angle = -Math.PI / 2
  const slices = vals.map((v, i) => {
    const sweep = (v / total) * 2 * Math.PI
    const x1 = R + R * Math.cos(angle), y1 = R + R * Math.sin(angle)
    angle += sweep
    const x2 = R + R * Math.cos(angle), y2 = R + R * Math.sin(angle)
    const large = sweep > Math.PI ? 1 : 0
    return { d: `M${R},${R} L${x1},${y1} A${R},${R} 0 ${large} 1 ${x2},${y2} Z`, color: colors[i], pct: Math.round((v / total) * 100) }
  })
  return (
    <svg width={44} height={44}>
      {slices.map((s, i) => <path key={i} d={s.d} fill={s.color} opacity={0.85}/>)}
      <circle cx={R} cy={R} r={r} fill="var(--surface-2)"/>
    </svg>
  )
}

export default function Dashboard() {
  const nav = useNavigate()
  const [sidebarOpen, setSidebarOpen] = useState(true)
  const [files, setFiles] = useState<File[]>(FILES)
  const [search, setSearch] = useState('')
  const [activeTab, setActiveTab] = useState<'recent'|'starred'|'shared'>('recent')
  const [selected, setSelected] = useState<string|null>(null)
  const [greeting, setGreeting] = useState('')
  const [prompt, setPrompt] = useState('')
  const [generating, setGenerating] = useState(false)
  const [showUpgrade, setShowUpgrade] = useState(false)
  const promptRef = useRef<HTMLTextAreaElement>(null)
  const user = (() => { try { return JSON.parse(localStorage.getItem('gpe_user') || 'null') } catch { return null } })()

  useEffect(() => {
    const h = new Date().getHours()
    setGreeting(h < 12 ? 'Good morning' : h < 17 ? 'Good afternoon' : 'Good evening')
  }, [])

  const filtered = files
    .filter(f => activeTab === 'starred' ? f.starred : true)
    .filter(f => f.name.toLowerCase().includes(search.toLowerCase()))

  const toggleStar = (id: string) => setFiles(p => p.map(f => f.id === id ? { ...f, starred: !f.starred } : f))

  const handlePrompt = async () => {
    if (!prompt.trim()) return
    setGenerating(true)
    await new Promise(r => setTimeout(r, 1600))
    setGenerating(false)
    nav('/excel')
  }

  const typeIcon = (type: File['type']) => ({
    excel: { icon: 'XL', color: 'var(--green)', bg: 'var(--green-dim)', border: 'var(--green-border)' },
    doc: { icon: 'W', color: 'var(--blue)', bg: 'var(--blue-dim)', border: 'var(--blue-border)' },
    pdf: { icon: 'PDF', color: 'var(--red)', bg: 'var(--red-dim)', border: 'var(--red-border)' },
    csv: { icon: 'CSV', color: 'var(--yellow)', bg: 'var(--yellow-dim)', border: 'var(--yellow-border)' },
    pptx: { icon: 'PPT', color: 'var(--orange)', bg: 'var(--orange-dim)', border: 'rgba(249,115,22,0.2)' },
  }[type])

  return (
    <div style={{ display: 'flex', flexDirection: 'column', height: '100vh', background: 'var(--bg)', fontFamily: 'var(--font-body)' }}>
      <style>{`
        .file-row:hover { background: var(--surface) !important; }
        .file-row:hover .file-actions { opacity: 1 !important; }
        .file-actions { opacity: 0; transition: opacity 0.14s; display: flex; gap: 4px; align-items: center; }
        .metric-card:hover { border-color: var(--border-2) !important; transform: translateY(-1px); box-shadow: var(--shadow); }
        .quick-action:hover { border-color: var(--blue-border) !important; background: var(--blue-dim) !important; }
        .quick-action:hover .qa-icon { color: var(--blue) !important; }
        .quick-action:hover .qa-label { color: var(--blue) !important; }
        .prompt-footer { opacity: 0; transition: opacity 0.2s; }
        .prompt-area:focus-within .prompt-footer { opacity: 1 !important; }
      `}</style>

      <Header toggleSidebar={() => setSidebarOpen(p => !p)} />
      <div style={{ display: 'flex', flex: 1, overflow: 'hidden' }}>
        <Sidebar isOpen={sidebarOpen} />
        <main style={{ flex: 1, overflowY: 'auto', background: 'var(--bg-2)', display: 'flex', flexDirection: 'column' }}>

          {/* TOP SECTION */}
          <div style={{ padding: '22px 28px 0', borderBottom: '1px solid var(--border)', background: 'var(--bg)' }}>
            <div style={{ display: 'flex', alignItems: 'flex-start', justifyContent: 'space-between', marginBottom: 16 }}>
              <div>
                <div style={{ fontSize: '0.7rem', fontFamily: 'var(--font-mono)', letterSpacing: '0.08em', color: 'var(--text-muted)', textTransform: 'uppercase', marginBottom: 4 }}>
                  {new Date().toLocaleDateString('en-US', { weekday: 'long', month: 'long', day: 'numeric', year: 'numeric' })}
                </div>
                <h1 style={{ fontFamily: 'var(--font-display)', fontSize: '1.5rem', letterSpacing: -0.5, color: 'var(--text)' }}>
                  {greeting}, {user?.name?.split(' ')[0] || 'there'} 👋
                </h1>
                <p style={{ fontSize: '0.78rem', color: 'var(--text-muted)', marginTop: 2 }}>
                  You have <span style={{ color: 'var(--blue)', fontWeight: 600 }}>23 generations</span> remaining on the free plan.
                </p>
              </div>
              <div style={{ display: 'flex', gap: 8 }}>
                <button className="btn btn-blue btn-sm" onClick={() => setShowUpgrade(true)}>↑ Upgrade to Pro</button>
                <button className="btn btn-outline btn-sm" onClick={() => nav('/excel')}>+ New File</button>
              </div>
            </div>

            {/* Prompt */}
            <div className="prompt-area" style={{ background: 'var(--surface)', border: '1px solid var(--border-2)', borderRadius: 10, overflow: 'hidden', marginBottom: 18 }}>
              <div style={{ display: 'flex', alignItems: 'flex-start', gap: 10, padding: '12px 14px' }}>
                <div style={{ width: 26, height: 26, background: 'var(--blue-dim)', border: '1px solid var(--blue-border)', borderRadius: 6, display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 14, flexShrink: 0, marginTop: 2, color: 'var(--blue)' }}>⊞</div>
                <textarea ref={promptRef} className="input" value={prompt} onChange={e => setPrompt(e.target.value)} onKeyDown={e => { if (e.key === 'Enter' && (e.metaKey || e.ctrlKey)) handlePrompt() }}
                  placeholder='Ask GPT-EXCEL anything… "Create a Q4 sales forecast" or "Build a KPI dashboard"'
                  style={{ flex: 1, border: 'none', background: 'transparent', resize: 'none', minHeight: 20, maxHeight: 120, boxShadow: 'none', fontSize: '0.875rem', padding: 0, lineHeight: 1.6, color: 'var(--text)' }} rows={1}
                  onInput={e => { const el = e.currentTarget; el.style.height = 'auto'; el.style.height = el.scrollHeight + 'px' }}
                />
              </div>
              <div className="prompt-footer" style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '8px 14px', borderTop: '1px solid var(--border)', background: 'var(--bg-3)' }}>
                <div style={{ display: 'flex', gap: 6 }}>
                  {['Excel', 'Document', 'Chart', 'Pivot', 'Dashboard'].map(t => (
                    <button key={t} className="btn btn-xs btn-outline" onClick={() => setPrompt(p => p ? p + ` [${t}]` : `Create a ${t.toLowerCase()}: `)}>{t}</button>
                  ))}
                </div>
                <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                  <span style={{ fontSize: '0.7rem', color: 'var(--text-muted)', fontFamily: 'var(--font-mono)' }}><kbd>⌘</kbd><kbd>↵</kbd></span>
                  <button className="btn btn-primary btn-sm" onClick={handlePrompt} disabled={!prompt.trim() || generating} style={{ minWidth: 80 }}>
                    {generating ? <><span className="spinner" style={{ width: 12, height: 12 }}/> Generating</> : 'Generate →'}
                  </button>
                </div>
              </div>
            </div>

            {/* Quick actions bar */}
            <div style={{ display: 'flex', gap: 0, borderTop: '1px solid var(--border)', marginLeft: -28, marginRight: -28, paddingLeft: 28, overflowX: 'auto' }}>
              {[
                { icon: '⊞', label: 'New Spreadsheet', color: 'var(--green)' },
                { icon: '◱', label: 'New Document', color: 'var(--blue)' },
                { icon: '◬', label: 'New Chart', color: 'var(--yellow)' },
                { icon: '◻', label: 'New Slide Deck', color: 'var(--purple)' },
                { icon: '⌘', label: 'New Workflow', color: 'var(--orange)' },
                { icon: '◈', label: 'Organize Files', color: 'var(--teal)' },
                { icon: '◎', label: 'Voice Input', color: 'var(--pink)' },
              ].map((a, i) => (
                <button key={i} className="quick-action btn btn-ghost btn-sm"
                  style={{ borderRadius: 0, height: 36, borderRight: '1px solid var(--border)', fontSize: '0.73rem', gap: 5, flexShrink: 0, whiteSpace: 'nowrap', transition: 'all var(--tr)' }}
                  onClick={() => nav('/excel')}
                >
                  <span className="qa-icon" style={{ color: a.color, transition: 'color var(--tr)' }}>{a.icon}</span>
                  <span className="qa-label" style={{ color: 'var(--text-sec)', transition: 'color var(--tr)' }}>{a.label}</span>
                </button>
              ))}
            </div>
          </div>

          {/* BODY */}
          <div style={{ flex: 1, padding: '18px 28px', display: 'flex', gap: 18 }}>

            {/* LEFT */}
            <div style={{ flex: 1, display: 'flex', flexDirection: 'column', gap: 14, minWidth: 0 }}>

              {/* KPI cards */}
              <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4, 1fr)', gap: 10 }}>
                {KPIS.map((k, i) => (
                  <div key={i} className="metric-card" style={{ background: 'var(--surface)', border: '1px solid var(--border)', borderRadius: 9, padding: '13px 14px', transition: 'all 0.2s', cursor: 'pointer' }}>
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start' }}>
                      <div>
                        <div style={{ fontFamily: 'var(--font-display)', fontSize: '1.5rem', fontWeight: 800, letterSpacing: -0.5, color: k.color }}>{k.value}</div>
                        <div style={{ fontSize: '0.7rem', color: 'var(--text-muted)', marginTop: 2, letterSpacing: '0.02em' }}>{k.label}</div>
                      </div>
                      <Sparkline data={k.sparkline} color={k.color}/>
                    </div>
                    <div style={{ marginTop: 7, fontSize: '0.7rem', fontFamily: 'var(--font-mono)', color: k.up ? 'var(--green)' : 'var(--red)' }}>
                      {k.up ? '↑' : '↓'} {k.change}
                    </div>
                  </div>
                ))}
              </div>

              {/* Visual analytics row */}
              <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr', gap: 12 }}>
                {/* Usage breakdown */}
                <div className="chart-container">
                  <div className="chart-title">Generation Types</div>
                  <div className="chart-subtitle">This month</div>
                  <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
                    <MiniDonut vals={[60, 25, 15]} colors={['var(--blue)', 'var(--green)', 'var(--purple)']}/>
                    <div style={{ flex: 1, display: 'flex', flexDirection: 'column', gap: 5 }}>
                      {[['Excel', '60%', 'var(--blue)'], ['Documents', '25%', 'var(--green)'], ['Charts', '15%', 'var(--purple)']].map(([l, v, c]) => (
                        <div key={l as string} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                          <div style={{ display: 'flex', alignItems: 'center', gap: 5 }}>
                            <div className="legend-dot" style={{ background: c as string }}/>
                            <span style={{ fontSize: '0.72rem', color: 'var(--text-sec)' }}>{l as string}</span>
                          </div>
                          <span style={{ fontSize: '0.72rem', fontFamily: 'var(--font-mono)', color: c as string }}>{v as string}</span>
                        </div>
                      ))}
                    </div>
                  </div>
                </div>

                {/* Activity bars */}
                <div className="chart-container">
                  <div className="chart-title">Daily Activity</div>
                  <div className="chart-subtitle">Files generated per day</div>
                  <div style={{ display: 'flex', flexDirection: 'column', gap: 3 }}>
                    {[['Mon', 8], ['Tue', 14], ['Wed', 6], ['Thu', 18], ['Fri', 12], ['Sat', 3], ['Sun', 5]].map(([day, val]) => (
                      <div key={day as string} style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                        <span style={{ fontSize: '0.65rem', color: 'var(--text-muted)', fontFamily: 'var(--font-mono)', width: 24, flexShrink: 0 }}>{day as string}</span>
                        <div style={{ flex: 1, height: 6, background: 'var(--border-2)', borderRadius: 3, overflow: 'hidden' }}>
                          <div style={{ height: '100%', width: `${((val as number) / 18) * 100}%`, background: 'var(--blue)', borderRadius: 3, transition: 'width 0.5s ease' }}/>
                        </div>
                        <span style={{ fontSize: '0.65rem', color: 'var(--text-muted)', fontFamily: 'var(--font-mono)', width: 16 }}>{val as number}</span>
                      </div>
                    ))}
                  </div>
                </div>

                {/* Storage usage */}
                <div className="chart-container">
                  <div className="chart-title">Storage Breakdown</div>
                  <div className="chart-subtitle">3.4 GB of 5 GB used</div>
                  <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
                    {[
                      { label: 'Excel Files', size: '1.8 GB', pct: 36, color: 'var(--green)' },
                      { label: 'Documents', size: '0.9 GB', pct: 18, color: 'var(--blue)' },
                      { label: 'PDFs', size: '0.5 GB', pct: 10, color: 'var(--red)' },
                      { label: 'Other', size: '0.2 GB', pct: 4, color: 'var(--yellow)' },
                    ].map(s => (
                      <div key={s.label}>
                        <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: 3, fontSize: '0.7rem' }}>
                          <span style={{ color: 'var(--text-sec)' }}>{s.label}</span>
                          <span style={{ color: 'var(--text-muted)', fontFamily: 'var(--font-mono)' }}>{s.size}</span>
                        </div>
                        <div className="progress-track" style={{ height: 3 }}><div className="progress-fill" style={{ width: `${s.pct}%`, background: s.color }}/></div>
                      </div>
                    ))}
                  </div>
                </div>
              </div>

              {/* Files */}
              <div style={{ background: 'var(--surface)', border: '1px solid var(--border)', borderRadius: 10, overflow: 'hidden', flex: 1 }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: 8, padding: '10px 14px', borderBottom: '1px solid var(--border)' }}>
                  <div className="tabs" style={{ flex: 1 }}>
                    {(['recent', 'starred', 'shared'] as const).map(tab => (
                      <div key={tab} className={`tab${activeTab === tab ? ' active' : ''}`} onClick={() => setActiveTab(tab)} style={{ textTransform: 'capitalize' }}>{tab}</div>
                    ))}
                  </div>
                  <div className="search-bar" style={{ width: 180 }}>
                    <svg width={12} height={12} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth={2} style={{ color: 'var(--text-muted)', flexShrink: 0 }}><circle cx={11} cy={11} r={8}/><path d="m21 21-4.35-4.35"/></svg>
                    <input placeholder="Search files..." value={search} onChange={e => setSearch(e.target.value)}/>
                  </div>
                  <button className="btn btn-primary btn-sm" onClick={() => nav('/excel')}>+ New</button>
                </div>
                <div style={{ overflowY: 'auto', maxHeight: 280 }}>
                  {filtered.length === 0 ? (
                    <div className="empty-state"><h3>No files found</h3><p>Try adjusting your search or create a new file</p><button className="btn btn-primary btn-sm mt-3" onClick={() => nav('/excel')}>Create file →</button></div>
                  ) : filtered.map(f => {
                    const ti = typeIcon(f.type)
                    return (
                      <div key={f.id} className="file-row" style={{ display: 'flex', alignItems: 'center', gap: 10, padding: '8px 14px', borderBottom: '1px solid var(--border)', cursor: 'pointer', transition: 'background var(--tr)', background: selected === f.id ? 'var(--blue-dim)' : 'transparent' }}
                        onClick={() => setSelected(f.id === selected ? null : f.id)} onDoubleClick={() => nav('/excel')}
                      >
                        <div style={{ width: 32, height: 32, borderRadius: 6, background: ti.bg, border: `1px solid ${ti.border}`, display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: '0.6rem', fontFamily: 'var(--font-mono)', fontWeight: 700, color: ti.color, flexShrink: 0 }}>{ti.icon}</div>
                        <div style={{ flex: 1, minWidth: 0 }}>
                          <div style={{ fontSize: '0.8125rem', fontWeight: 500, color: 'var(--text)', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{f.name}</div>
                          <div style={{ fontSize: '0.7rem', color: 'var(--text-muted)', marginTop: 1 }}>{f.modified} · {f.size}</div>
                        </div>
                        <div className="file-actions">
                          <button className="btn btn-icon-sm btn-ghost" onClick={e => { e.stopPropagation(); toggleStar(f.id) }} style={{ color: f.starred ? 'var(--yellow)' : 'var(--text-muted)' }}>{f.starred ? '★' : '☆'}</button>
                          <button className="btn btn-icon-sm btn-ghost" onClick={e => { e.stopPropagation(); nav('/excel') }} style={{ color: 'var(--text-muted)' }}>↗</button>
                          <button className="btn btn-icon-sm btn-ghost" style={{ color: 'var(--text-muted)' }}>↓</button>
                          <button className="btn btn-icon-sm btn-ghost" style={{ color: 'var(--text-muted)' }}>…</button>
                        </div>
                      </div>
                    )
                  })}
                </div>
              </div>

              {/* Activity */}
              <div style={{ background: 'var(--surface)', border: '1px solid var(--border)', borderRadius: 10, overflow: 'hidden' }}>
                <div className="panel-header">
                  <span>Activity Feed</span>
                  <button className="btn btn-xs btn-ghost" style={{ marginLeft: 'auto' }}>View all</button>
                </div>
                <div style={{ padding: '6px 0' }}>
                  {ACTIVITIES.map(a => (
                    <div key={a.id} style={{ display: 'flex', alignItems: 'flex-start', gap: 10, padding: '7px 14px' }}>
                      <div style={{ width: 28, height: 28, flexShrink: 0, display: 'flex', alignItems: 'center', justifyContent: 'center', background: a.type === 'warning' ? 'var(--yellow-dim)' : a.type === 'success' ? 'var(--green-dim)' : 'var(--blue-dim)', color: a.type === 'warning' ? 'var(--yellow)' : a.type === 'success' ? 'var(--green)' : 'var(--blue)', fontSize: 14, border: `1px solid ${a.type === 'warning' ? 'var(--yellow-border)' : a.type === 'success' ? 'var(--green-border)' : 'var(--blue-border)'}`, borderRadius: 6 }}>{a.icon}</div>
                      <div style={{ flex: 1, minWidth: 0 }}>
                        <div style={{ fontSize: '0.8rem', color: 'var(--text)', lineHeight: 1.4 }}>{a.text}</div>
                        <div style={{ fontSize: '0.7rem', color: 'var(--text-muted)', marginTop: 2, fontFamily: 'var(--font-mono)' }}>{a.time}</div>
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            </div>

            {/* RIGHT */}
            <div style={{ width: 230, flexShrink: 0, display: 'flex', flexDirection: 'column', gap: 12 }}>

              {/* Quick actions */}
              <div style={{ background: 'var(--surface)', border: '1px solid var(--border)', borderRadius: 10, overflow: 'hidden' }}>
                <div className="panel-header">Quick Actions</div>
                <div style={{ padding: '8px 10px', display: 'flex', flexDirection: 'column', gap: 6 }}>
                  {[
                    { icon: '⊞', label: 'Excel from prompt', desc: 'Generate .xlsx from text', color: 'var(--green)' },
                    { icon: '◱', label: 'Write a document', desc: 'Reports, proposals, CVs', color: 'var(--blue)' },
                    { icon: '◬', label: 'Build a chart', desc: 'From data or prompt', color: 'var(--yellow)' },
                    { icon: '◻', label: 'Create slide deck', desc: 'AI-generated PowerPoint', color: 'var(--purple)' },
                    { icon: '◎', label: 'Voice command', desc: 'Speak your request', color: 'var(--teal)' },
                  ].map(a => (
                    <button key={a.label} onClick={() => nav('/excel')} style={{ display: 'flex', alignItems: 'flex-start', gap: 10, padding: '9px 11px', background: 'var(--surface-2)', border: '1px solid var(--border)', borderRadius: 7, cursor: 'pointer', textAlign: 'left', transition: 'all var(--tr)', width: '100%' }}
                      onMouseEnter={e => { (e.currentTarget as HTMLElement).style.borderColor = 'var(--blue-border)'; (e.currentTarget as HTMLElement).style.background = 'var(--blue-dim)' }}
                      onMouseLeave={e => { (e.currentTarget as HTMLElement).style.borderColor = 'var(--border)'; (e.currentTarget as HTMLElement).style.background = 'var(--surface-2)' }}
                    >
                      <span style={{ fontSize: 16, lineHeight: 1, color: a.color, flexShrink: 0, marginTop: 1 }}>{a.icon}</span>
                      <div>
                        <div style={{ fontSize: '0.78rem', fontWeight: 600, color: 'var(--text)', marginBottom: 1 }}>{a.label}</div>
                        <div style={{ fontSize: '0.68rem', color: 'var(--text-muted)', lineHeight: 1.3 }}>{a.desc}</div>
                      </div>
                    </button>
                  ))}
                </div>
              </div>

              {/* Usage */}
              <div style={{ background: 'var(--surface)', border: '1px solid var(--border)', borderRadius: 10, overflow: 'hidden' }}>
                <div className="panel-header">
                  <span>Usage</span>
                  <span className="badge badge-yellow" style={{ marginLeft: 'auto' }}>Free Plan</span>
                </div>
                <div style={{ padding: '12px 14px' }}>
                  {[
                    { label: 'AI Generations', used: 27, total: 50, pct: 54, color: 'var(--blue)' },
                    { label: 'Storage', used: 3.4, total: 5, pct: 68, color: 'var(--green)', suffix: ' GB' },
                    { label: 'Voice Minutes', used: 8, total: 30, pct: 27, color: 'var(--purple)' },
                    { label: 'Workflows', used: 2, total: 3, pct: 67, color: 'var(--orange)' },
                  ].map(u => (
                    <div key={u.label} style={{ marginBottom: 10 }}>
                      <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: 4 }}>
                        <span style={{ fontSize: '0.72rem', color: 'var(--text-sec)' }}>{u.label}</span>
                        <span style={{ fontSize: '0.72rem', fontFamily: 'var(--font-mono)', color: 'var(--text-muted)' }}>{u.used}{u.suffix||''}/{u.total}{u.suffix||''}</span>
                      </div>
                      <div className="progress-track"><div className="progress-fill" style={{ width: `${u.pct}%`, background: u.color }}/></div>
                    </div>
                  ))}
                  <button className="btn btn-blue btn-sm w-full" style={{ marginTop: 6 }} onClick={() => setShowUpgrade(true)}>Upgrade to Pro →</button>
                </div>
              </div>

              {/* Recent prompts */}
              <div style={{ background: 'var(--surface)', border: '1px solid var(--border)', borderRadius: 10, overflow: 'hidden' }}>
                <div className="panel-header">Recent Prompts</div>
                {['Monthly budget tracker with charts', 'Sales pipeline Q4 forecast', 'Employee attendance sheet', 'Revenue comparison 2024 vs 2025', 'KPI dashboard with sparklines'].map((p, i) => (
                  <div key={i} style={{ padding: '7px 13px', fontSize: '0.73rem', color: 'var(--text-sec)', cursor: 'pointer', borderBottom: '1px solid var(--border)', transition: 'background var(--tr)', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}
                    onMouseEnter={e => (e.currentTarget.style.background = 'var(--accent-dim)')}
                    onMouseLeave={e => (e.currentTarget.style.background = 'transparent')}
                    onClick={() => { setPrompt(p); promptRef.current?.focus() }}
                  >{p}</div>
                ))}
              </div>
            </div>
          </div>
        </main>
        <AIChatPanel />
      </div>

      {/* Upgrade modal */}
      {showUpgrade && (
        <div className="modal-backdrop" onClick={() => setShowUpgrade(false)}>
          <div className="modal-box" style={{ width: 460 }} onClick={e => e.stopPropagation()}>
            <div className="modal-title">Upgrade to Pro</div>
            <div style={{ textAlign: 'center', padding: '8px 0 16px' }}>
              <div style={{ fontFamily: 'var(--font-display)', fontSize: '2.5rem', fontWeight: 800, color: 'var(--blue)', marginBottom: 4 }}>$12<span style={{ fontSize: '1rem', color: 'var(--text-muted)' }}>/mo</span></div>
              <div style={{ color: 'var(--text-sec)', fontSize: '0.85rem' }}>Everything you need. No limits.</div>
            </div>
            <div style={{ display: 'flex', flexDirection: 'column', gap: 8, marginBottom: 20 }}>
              {['Unlimited AI generations', 'Unlimited storage (50 GB)', 'All voice features (Whisper)', 'Advanced automation workflows', 'Priority customer support', 'Plugin marketplace access', 'Custom AI model selection'].map(f => (
                <div key={f} style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: '0.8rem', color: 'var(--text-sec)' }}>
                  <span style={{ color: 'var(--green)', flexShrink: 0 }}>✓</span>{f}
                </div>
              ))}
            </div>
            <div style={{ display: 'flex', gap: 10 }}>
              <button className="btn btn-primary w-full" style={{ justifyContent: 'center' }}>Start Pro Trial →</button>
              <button className="btn btn-outline" onClick={() => setShowUpgrade(false)}>Cancel</button>
            </div>
          </div>
        </div>
      )}
    </div>
  )
}
