import { useState, useEffect, useRef } from 'react'
import { useNavigate } from 'react-router-dom'
import { useSelector } from 'react-redux'
import { RootState } from '../index'
import Sidebar from '../components/Sidebar'
import Header from '../components/Header'
import { api } from '../api'

type ViewMode = 'list' | 'grid'
type FileTab = 'uploaded' | 'browse' | 'search'
type FileKind = 'excel' | 'document' | 'presentation' | 'pdf' | 'image' | 'text' | 'folder' | 'other'

const FILE_MANAGER_STATE_KEY = 'file_manager_state'
const SERVER_BASE = 'http://localhost:3001'

const inferFileKind = (file: any): FileKind => {
  const type = String(file?.type || '').toLowerCase()
  const name = String(file?.name || '').toLowerCase()
  if (type === 'folder') return 'folder'
  if (['xlsx', 'xls', 'csv'].includes(type)) return 'excel'
  if (['doc', 'docx', 'md'].includes(type) || /report|document|notice|proposal|minutes|certificate|circular/.test(name)) return 'document'
  if (['ppt', 'pptx'].includes(type) || /presentation|deck|slides|slide/.test(name)) return 'presentation'
  if (type === 'pdf') return 'pdf'
  if (['png', 'jpg', 'jpeg', 'gif', 'webp', 'svg'].includes(type)) return 'image'
  if (['txt', 'json'].includes(type)) return 'text'
  return 'other'
}

const inferDocType = (name: string) => {
  const lower = name.toLowerCase()
  if (lower.includes('proposal')) return 'proposal'
  if (lower.includes('notice')) return 'notice'
  if (lower.includes('minutes')) return 'minutes'
  if (lower.includes('certificate')) return 'certificate'
  if (lower.includes('circular')) return 'circular'
  return 'report'
}

const FILE_THEME: Record<FileKind, { label: string; short: string; bg: string; text: string; glow: string }> = {
  excel: { label: 'Excel', short: 'XL', bg: 'linear-gradient(145deg, #107c41, #0f5132)', text: '#dcfce7', glow: 'rgba(16,124,65,0.22)' },
  document: { label: 'Word', short: 'W', bg: 'linear-gradient(145deg, #185abd, #143a7b)', text: '#dbeafe', glow: 'rgba(24,90,189,0.22)' },
  presentation: { label: 'PPT', short: 'P', bg: 'linear-gradient(145deg, #d24726, #8a2f19)', text: '#ffedd5', glow: 'rgba(210,71,38,0.22)' },
  pdf: { label: 'PDF', short: 'PDF', bg: 'linear-gradient(145deg, #dc2626, #7f1d1d)', text: '#fee2e2', glow: 'rgba(220,38,38,0.22)' },
  image: { label: 'Image', short: 'IMG', bg: 'linear-gradient(145deg, #7c3aed, #4c1d95)', text: '#ede9fe', glow: 'rgba(124,58,237,0.22)' },
  text: { label: 'Text', short: 'TXT', bg: 'linear-gradient(145deg, #475569, #1e293b)', text: '#e2e8f0', glow: 'rgba(71,85,105,0.22)' },
  folder: { label: 'Folder', short: 'DIR', bg: 'linear-gradient(145deg, #f59e0b, #b45309)', text: '#fef3c7', glow: 'rgba(245,158,11,0.22)' },
  other: { label: 'File', short: 'FILE', bg: 'linear-gradient(145deg, #6b7280, #374151)', text: '#f3f4f6', glow: 'rgba(107,114,128,0.22)' }
}

function FileMark({ file, compact = false }: { file: any; compact?: boolean }) {
  const kind = inferFileKind(file)
  const theme = FILE_THEME[kind]
  const size = compact ? 32 : 42
  const imageUrl = kind === 'image' && file?.url ? `${SERVER_BASE}${file.url}` : ''
  return (
    <div style={{ width: size, height: size, borderRadius: 12, background: theme.bg, color: theme.text, display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', boxShadow: `0 10px 24px ${theme.glow}`, border: '1px solid rgba(255,255,255,0.12)', position: 'relative', overflow: 'hidden', flexShrink: 0 }}>
      {imageUrl ? (
        <>
          <img src={imageUrl} alt={file?.name || 'uploaded image'} style={{ width: '100%', height: '100%', objectFit: 'cover' }} />
          <div style={{ position: 'absolute', left: 4, right: 4, bottom: 4, fontSize: compact ? '0.42rem' : '0.5rem', fontWeight: 800, letterSpacing: '0.08em', textAlign: 'center', color: '#fff', background: 'rgba(15, 23, 42, 0.72)', borderRadius: 999, padding: '2px 4px', backdropFilter: 'blur(4px)' }}>
            LOGO
          </div>
        </>
      ) : (
        <>
          <div style={{ position: 'absolute', inset: 0, background: 'linear-gradient(180deg, rgba(255,255,255,0.18), transparent 55%)' }} />
          <div style={{ position: 'absolute', top: 0, right: 0, width: compact ? 10 : 14, height: compact ? 10 : 14, background: 'rgba(255,255,255,0.18)', clipPath: 'polygon(0 0, 100% 0, 100% 100%)' }} />
          <div style={{ fontSize: compact ? '0.68rem' : '0.82rem', fontWeight: 800, letterSpacing: '0.04em', position: 'relative' }}>{theme.short}</div>
          {!compact && <div style={{ fontSize: '0.46rem', textTransform: 'uppercase', letterSpacing: '0.08em', opacity: 0.85, position: 'relative' }}>{theme.label}</div>}
        </>
      )}
    </div>
  )
}

export default function FileManager() {
  const navigate = useNavigate()
  const user = useSelector((s: RootState) => s.app.user)
  const userDepartment = (user as any)?.department as string | undefined
  const [sidebarOpen, setSidebarOpen] = useState(true)
  const [tab, setTab] = useState<FileTab>(() => {
    try { return (JSON.parse(localStorage.getItem(FILE_MANAGER_STATE_KEY) || '{}').tab as FileTab) || 'uploaded' } catch { return 'uploaded' }
  })
  const [viewMode, setViewMode] = useState<ViewMode>(() => {
    try { return (JSON.parse(localStorage.getItem(FILE_MANAGER_STATE_KEY) || '{}').viewMode as ViewMode) || 'list' } catch { return 'list' }
  })
  const [files, setFiles] = useState<any[]>([])
  const [loading, setLoading] = useState(false)
  const [search, setSearch] = useState(() => {
    try { return JSON.parse(localStorage.getItem(FILE_MANAGER_STATE_KEY) || '{}').search || '' } catch { return '' }
  })
  const [filterType, setFilterType] = useState(() => {
    try { return JSON.parse(localStorage.getItem(FILE_MANAGER_STATE_KEY) || '{}').filterType || '' } catch { return '' }
  })
  const [filterDept, setFilterDept] = useState(() => {
    try { return JSON.parse(localStorage.getItem(FILE_MANAGER_STATE_KEY) || '{}').filterDept || '' } catch { return '' }
  })
  const [browseDir, setBrowseDir] = useState('')
  const [browseItems, setBrowseItems] = useState<any[]>([])
  const [browsePath, setBrowsePath] = useState('')
  const [browseParent, setBrowseParent] = useState('')
  const [diskSearch, setDiskSearch] = useState('')
  const [diskResults, setDiskResults] = useState<any[]>([])
  const [diskSearching, setDiskSearching] = useState(false)
  const [newDirName, setNewDirName] = useState('')
  const [showNewDir, setShowNewDir] = useState(false)
  const [uploading, setUploading] = useState(false)
  const [selected, setSelected] = useState<Set<string>>(new Set())
  const [dragOver, setDragOver] = useState(false)
  const fileRef = useRef<HTMLInputElement>(null)

  useEffect(() => { loadFiles() }, [filterType, filterDept])

  useEffect(() => {
    localStorage.setItem(FILE_MANAGER_STATE_KEY, JSON.stringify({ tab, viewMode, search, filterType, filterDept }))
  }, [tab, viewMode, search, filterType, filterDept])

  const loadFiles = async () => {
    setLoading(true)
    try {
      const res = await api.getFiles({ type: filterType || undefined, department: filterDept || undefined, search: search || undefined })
      setFiles(res)
    } catch { setFiles([]) }
    setLoading(false)
  }

  const browseDirectory = async (path?: string) => {
    try {
      const res = await api.browseDir(path || browseDir || '')
      setBrowseItems(res.items)
      setBrowsePath(res.path)
      setBrowseParent(res.parent)
    } catch (e: any) { setBrowseItems([]) }
  }

  const handleUpload = async (fileList: FileList | null) => {
    if (!fileList?.length) return
    setUploading(true)
    try {
      const files = Array.from(fileList)
      await api.uploadFiles(files, filterDept || userDepartment || 'general', user?.email || 'user')
      await loadFiles()
    } catch { alert('Upload failed — is backend running?') }
    setUploading(false)
  }

  const handleDelete = async (id: string) => {
    if (!confirm('Delete this file?')) return
    try { await api.deleteFile(id); await loadFiles() }
    catch { alert('Delete failed') }
  }

  const handleDiskSearch = async () => {
    if (!diskSearch.trim()) return
    setDiskSearching(true)
    try {
      const res = await api.searchDisk(diskSearch)
      setDiskResults(res)
    } catch { setDiskResults([]) }
    setDiskSearching(false)
  }

  const handleMkdir = async () => {
    if (!newDirName.trim() || !browsePath) return
    const fullPath = `${browsePath}/${newDirName}`
    try { await api.mkdir(fullPath); browseDirectory(browsePath); setNewDirName(''); setShowNewDir(false) }
    catch { alert('Failed to create directory') }
  }

  const openManagedFile = async (file: any) => {
    const kind = inferFileKind(file)
    const type = String(file?.type || '').toLowerCase()
    if (kind === 'excel') {
      localStorage.setItem('excel_autoload_file_id', file.id)
      navigate('/excel')
      return
    }
    if (['doc', 'docx', 'ppt', 'pptx', 'pdf'].includes(type)) {
      alert('Inline restore currently works for Excel files and text-based AI exports. Use download or the source app for this file type.')
      return
    }
    if (!file.url) return
    try {
      const response = await fetch(`${SERVER_BASE}${file.url}`)
      if (!response.ok) throw new Error(`HTTP ${response.status}`)
      const raw = await response.text()
      if (kind === 'document' || (kind === 'text' && /report|document|notice|proposal|minutes|certificate|circular/i.test(file.name))) {
        const doc = {
          id: `file_${file.id}`,
          title: file.name.replace(/\.[^.]+$/, ''),
          type: inferDocType(file.name),
          content: raw,
          createdAt: file.uploadedAt || new Date().toISOString()
        }
        const existing = JSON.parse(localStorage.getItem('ai_documents') || '[]')
        localStorage.setItem('ai_documents', JSON.stringify([doc, ...existing.filter((item: any) => item.id !== doc.id)]))
        localStorage.setItem('ai_documents_last_open', doc.id)
        navigate('/documents')
        return
      }
      if (kind === 'presentation' || (kind === 'text' && /presentation|deck|slides|slide/i.test(file.name))) {
        let parsed: any = null
        try { parsed = JSON.parse(raw) } catch {}
        const slides = Array.isArray(parsed?.slides)
          ? parsed.slides.map((slide: any, index: number) => ({
              id: index + 1,
              title: String(slide.title || `Slide ${index + 1}`),
              content: String(slide.content || ''),
              chart: slide.chart ? String(slide.chart) : '',
              notes: slide.notes ? String(slide.notes) : ''
            }))
          : [{ id: 1, title: file.name.replace(/\.[^.]+$/, ''), content: raw, chart: '', notes: '' }]
        localStorage.setItem('ai_ppt_slides', JSON.stringify(slides))
        localStorage.setItem('ai_ppt_draft_meta', JSON.stringify({
          prompt: parsed?.deck_title || file.name.replace(/\.[^.]+$/, ''),
          theme: 'Executive Blue',
          audience: 'Management',
          goal: parsed?.subtitle || 'Present workbook insights'
        }))
        localStorage.setItem('ai_ppt_context', `Loaded from File Manager\nSource file: ${file.name}\nDepartment: ${file.department || 'general'}`)
        navigate('/powerpoint')
      }
    } catch (error) {
      console.error('Open file failed:', error)
      alert('Could not open this generated file.')
    }
  }

  const toggleSelect = (id: string) => {
    setSelected(prev => { const n = new Set(prev); n.has(id) ? n.delete(id) : n.add(id); return n })
  }

  const FILE_ICONS: Record<string, { icon: string; color: string }> = {
    xlsx: { icon: 'XL', color: '#22c55e' }, xls: { icon: 'XL', color: '#22c55e' },
    csv: { icon: 'CSV', color: '#14b8a6' }, pdf: { icon: 'PDF', color: '#ef4444' },
    docx: { icon: 'W', color: '#3b82f6' }, doc: { icon: 'W', color: '#3b82f6' },
    pptx: { icon: 'PPT', color: '#f97316' }, ppt: { icon: 'PPT', color: '#f97316' },
    png: { icon: 'PNG', color: '#a855f7' }, jpg: { icon: 'IMG', color: '#a855f7' },
    txt: { icon: 'TXT', color: '#9ca3af' }, folder: { icon: '📁', color: '#fbbf24' },
  }
  const getIcon = (type: string) => FILE_ICONS[type.toLowerCase()] || { icon: type.toUpperCase().slice(0,3), color: '#9ca3af' }

  const fmtSize = (bytes: number) => bytes < 1024 ? `${bytes} B` : bytes < 1024*1024 ? `${(bytes/1024).toFixed(1)} KB` : `${(bytes/1024/1024).toFixed(1)} MB`
  const fmtDate = (d: string | Date) => d ? new Date(d).toLocaleDateString() : '—'

  const departments = [...new Set(files.map(f => f.department).filter(Boolean))]
  const types = [...new Set(files.map(f => f.type).filter(Boolean))]
  const filteredFiles = files.filter(f =>
    (!search || f.name.toLowerCase().includes(search.toLowerCase()))
  )
  const kindCounts = filteredFiles.reduce((acc, file) => {
    const kind = inferFileKind(file)
    acc[kind] = (acc[kind] || 0) + 1
    return acc
  }, {} as Record<string, number>)

  return (
    <div style={{ display: 'flex', flexDirection: 'column', height: '100vh', background: 'var(--bg)' }}>
      <Header toggleSidebar={() => setSidebarOpen(p => !p)} />
      <div style={{ display: 'flex', flex: 1, overflow: 'hidden' }}>
        <Sidebar isOpen={sidebarOpen} />
        <main style={{ flex: 1, display: 'flex', flexDirection: 'column', overflow: 'hidden', background: 'var(--bg-2)' }}>

          {/* Header */}
          <div style={{ padding: '16px 20px', borderBottom: '1px solid var(--border)', background: 'var(--bg)', display: 'flex', alignItems: 'center', gap: 12, flexShrink: 0 }}>
            <div>
              <h1 style={{ fontSize: '1.2rem', letterSpacing: -0.5 }}>File Manager</h1>
              <p style={{ fontSize: '0.75rem', color: 'var(--text-secondary)' }}>Upload, browse, and manage files locally</p>
            </div>
            <div style={{ marginLeft: 'auto', display: 'flex', gap: 8, alignItems: 'center' }}>
              <button className="btn btn-icon-sm btn-ghost" onClick={() => setViewMode(v => v === 'list' ? 'grid' : 'list')}>
                {viewMode === 'list' ? '⊞' : '≡'}
              </button>
              <button className="btn btn-outline btn-sm" onClick={() => fileRef.current?.click()} disabled={uploading}>
                {uploading ? '⟳ Uploading...' : '+ Upload'}
              </button>
              <input ref={fileRef} type="file" multiple accept="*/*" style={{ display: 'none' }} onChange={e => handleUpload(e.target.files)} />
            </div>
          </div>

          {/* Tabs */}
          <div style={{ display: 'flex', gap: 0, borderBottom: '1px solid var(--border)', background: 'var(--bg)', flexShrink: 0 }}>
            {[{ id: 'uploaded' as FileTab, label: `Uploaded (${files.length})` }, { id: 'browse' as FileTab, label: 'Browse Disk' }, { id: 'search' as FileTab, label: 'Search Disk' }].map(t => (
              <button key={t.id} onClick={() => { setTab(t.id); if (t.id === 'browse' && !browsePath) browseDirectory() }} style={{ padding: '8px 16px', background: 'none', border: 'none', borderBottom: tab === t.id ? '2px solid var(--accent)' : '2px solid transparent', cursor: 'pointer', fontSize: '0.78rem', fontWeight: tab === t.id ? 600 : 400, color: tab === t.id ? 'var(--accent)' : 'var(--text-secondary)' }}>
                {t.label}
              </button>
            ))}
          </div>

          {/* ── UPLOADED TAB ── */}
          {tab === 'uploaded' && (
            <div style={{ flex: 1, display: 'flex', flexDirection: 'column', overflow: 'hidden' }}>
              <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4, minmax(0, 1fr))', gap: 10, padding: '12px 16px', borderBottom: '1px solid var(--border)', background: 'var(--bg)' }}>
                {[
                  { label: 'Excel', value: kindCounts.excel || 0, color: '#107c41' },
                  { label: 'Documents', value: kindCounts.document || 0, color: '#185abd' },
                  { label: 'Presentations', value: kindCounts.presentation || 0, color: '#d24726' },
                  { label: 'Selected', value: selected.size, color: '#7c3aed' }
                ].map(card => (
                  <div key={card.label} style={{ padding: '12px 14px', borderRadius: 14, background: 'var(--surface)', border: '1px solid var(--border)' }}>
                    <div style={{ fontSize: '0.68rem', color: 'var(--text-muted)', textTransform: 'uppercase', letterSpacing: '0.08em' }}>{card.label}</div>
                    <div style={{ fontSize: '1.2rem', fontWeight: 800, color: card.color, marginTop: 6 }}>{card.value}</div>
                  </div>
                ))}
              </div>

              {/* Filters */}
              <div style={{ display: 'flex', gap: 8, padding: '10px 16px', borderBottom: '1px solid var(--border)', background: 'var(--bg)', flexShrink: 0, alignItems: 'center', flexWrap: 'wrap' }}>
                <div className="search-bar" style={{ width: 220 }}>
                  <svg width={12} height={12} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth={2}><circle cx={11} cy={11} r={8}/><path d="m21 21-4.35-4.35"/></svg>
                  <input placeholder="Search files..." value={search} onChange={e => setSearch(e.target.value)} onKeyDown={e => e.key === 'Enter' && loadFiles()} />
                </div>
                <select className="input" style={{ width: 130, height: 32, padding: '0 8px', fontSize: '0.78rem' }} value={filterDept} onChange={e => setFilterDept(e.target.value)}>
                  <option value="">All Departments</option>
                  {departments.map(d => <option key={d} value={d}>{d}</option>)}
                </select>
                <select className="input" style={{ width: 110, height: 32, padding: '0 8px', fontSize: '0.78rem' }} value={filterType} onChange={e => setFilterType(e.target.value)}>
                  <option value="">All Types</option>
                  {types.map(t => <option key={t} value={t}>.{t}</option>)}
                </select>
                <button className="btn btn-ghost btn-sm" onClick={loadFiles}>↻ Refresh</button>
                {selected.size > 0 && (
                  <button className="btn btn-sm btn-outline" style={{ color: 'var(--red)', borderColor: 'var(--red)' }} onClick={async () => {
                    if (!confirm(`Delete ${selected.size} files?`)) return
                    await Promise.all([...selected].map(id => api.deleteFile(id).catch(()=>{})))
                    setSelected(new Set()); loadFiles()
                  }}>Delete {selected.size} selected</button>
                )}
              </div>

              {/* Drop zone */}
              <div
                onDragOver={e => { e.preventDefault(); setDragOver(true) }}
                onDragLeave={() => setDragOver(false)}
                onDrop={async e => { e.preventDefault(); setDragOver(false); await handleUpload(e.dataTransfer.files) }}
                style={{ flex: 1, overflowY: 'auto', background: dragOver ? 'rgba(59,130,246,0.05)' : 'transparent', transition: 'background 0.2s', outline: dragOver ? '2px dashed var(--accent)' : 'none', outlineOffset: -4 }}
              >
                {loading ? (
                  <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', height: 200 }}>
                    <div className="spinner" style={{ width: 24, height: 24 }} />
                  </div>
                ) : filteredFiles.length === 0 ? (
                  <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', height: 300, gap: 12, color: 'var(--text-muted)' }}>
                    <div style={{ fontSize: 48, opacity: 0.3 }}>📁</div>
                    <p>No files yet. Drop files here or click Upload.</p>
                    <button className="btn btn-primary btn-sm" onClick={() => fileRef.current?.click()}>Upload Files</button>
                  </div>
                ) : viewMode === 'list' ? (
                  <div>
                    <div style={{ display: 'grid', gridTemplateColumns: '32px 52px 1fr 120px 100px 130px 140px', padding: '8px 16px', borderBottom: '1px solid var(--border)', fontSize: '0.68rem', color: 'var(--text-muted)', fontWeight: 600, textTransform: 'uppercase' }}>
                      <div/>
                      <div/>
                      <div>Name</div>
                      <div>Department</div>
                      <div>Size</div>
                      <div>Uploaded</div>
                      <div>Actions</div>
                    </div>
                    {filteredFiles.map(f => {
                      const kind = inferFileKind(f)
                      const isSelected = selected.has(f.id)
                      return (
                        <div key={f.id} style={{ display: 'grid', gridTemplateColumns: '32px 52px 1fr 120px 100px 130px 140px', alignItems: 'center', padding: '8px 16px', borderBottom: '1px solid var(--border)', cursor: 'pointer', background: isSelected ? 'var(--accent-dim)' : 'transparent', transition: 'background var(--tr)' }}
                          onMouseEnter={e => { if (!isSelected) (e.currentTarget as HTMLElement).style.background = 'var(--surface-2)' }}
                          onMouseLeave={e => { if (!isSelected) (e.currentTarget as HTMLElement).style.background = 'transparent' }}
                        >
                          <input type="checkbox" checked={isSelected} onChange={() => toggleSelect(f.id)} style={{ cursor: 'pointer' }} />
                          <FileMark file={f} compact />
                          <div>
                            <div style={{ fontSize: '0.8rem', fontWeight: 500, color: 'var(--text)', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', maxWidth: 300 }}>{f.name}</div>
                            <div style={{ fontSize: '0.68rem', color: 'var(--text-muted)' }}>By {f.uploadedBy} · {FILE_THEME[kind].label}</div>
                          </div>
                          <div style={{ fontSize: '0.75rem', color: 'var(--text-secondary)' }}>{f.department}</div>
                          <div style={{ fontSize: '0.75rem', color: 'var(--text-muted)' }}>{fmtSize(f.size)}</div>
                          <div style={{ fontSize: '0.72rem', color: 'var(--text-muted)' }}>{fmtDate(f.uploadedAt)}</div>
                          <div style={{ display: 'flex', gap: 4 }}>
                            {['excel', 'document', 'presentation'].includes(kind) && (
                              <button className="btn btn-xs btn-outline" onClick={e => { e.stopPropagation(); openManagedFile(f) }}>{kind === 'excel' ? 'Open' : 'Restore'}</button>
                            )}
                            <button className="btn btn-xs btn-ghost" style={{ color: 'var(--red)' }} onClick={e => { e.stopPropagation(); handleDelete(f.id) }}>✕</button>
                          </div>
                        </div>
                      )
                    })}
                  </div>
                ) : (
                  <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(150px, 1fr))', gap: 10, padding: 16 }}>
                    {filteredFiles.map(f => {
                      const kind = inferFileKind(f)
                      const theme = FILE_THEME[kind]
                      return (
                        <div key={f.id} style={{ background: 'var(--surface)', border: '1px solid var(--border)', borderRadius: 10, padding: 14, cursor: 'pointer', transition: 'all var(--tr)', textAlign: 'center' }}
                          onMouseEnter={e => { (e.currentTarget as HTMLElement).style.borderColor = theme.glow.replace('0.22', '0.55'); (e.currentTarget as HTMLElement).style.background = 'var(--surface-2)' }}
                          onMouseLeave={e => { (e.currentTarget as HTMLElement).style.borderColor = 'var(--border)'; (e.currentTarget as HTMLElement).style.background = 'var(--surface)' }}
                        >
                          <div style={{ display: 'flex', justifyContent: 'center', marginBottom: 8 }}>
                            <FileMark file={f} />
                          </div>
                          <div style={{ fontSize: '0.75rem', fontWeight: 500, color: 'var(--text)', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{f.name}</div>
                          <div style={{ fontSize: '0.65rem', color: 'var(--text-muted)', marginTop: 4 }}>{FILE_THEME[kind].label} · {fmtSize(f.size)}</div>
                          <div style={{ display: 'flex', gap: 4, marginTop: 8, justifyContent: 'center' }}>
                            {['excel', 'document', 'presentation'].includes(kind) && <button className="btn btn-xs btn-outline" onClick={() => openManagedFile(f)}>{kind === 'excel' ? 'Open' : 'Restore'}</button>}
                            <button className="btn btn-xs btn-ghost" style={{ color: 'var(--red)' }} onClick={() => handleDelete(f.id)}>✕</button>
                          </div>
                        </div>
                      )
                    })}
                  </div>
                )}
              </div>
            </div>
          )}

          {/* ── BROWSE TAB ── */}
          {tab === 'browse' && (
            <div style={{ flex: 1, display: 'flex', flexDirection: 'column', overflow: 'hidden' }}>
              <div style={{ display: 'flex', gap: 8, padding: '10px 16px', borderBottom: '1px solid var(--border)', background: 'var(--bg)', flexShrink: 0, alignItems: 'center' }}>
                <div className="search-bar" style={{ flex: 1 }}>
                  <span style={{ fontSize: 11, color: 'var(--text-muted)' }}>📁</span>
                  <input placeholder="Directory path..." value={browseDir} onChange={e => setBrowseDir(e.target.value)} onKeyDown={e => e.key === 'Enter' && browseDirectory(browseDir)} />
                </div>
                <button className="btn btn-outline btn-sm" onClick={() => browseDirectory(browseDir)}>Go</button>
                {browsePath && browsePath !== browseParent && <button className="btn btn-ghost btn-sm" onClick={() => browseDirectory(browseParent)}>↑ Up</button>}
                <button className="btn btn-ghost btn-sm" onClick={() => setShowNewDir(p => !p)}>+ New Folder</button>
              </div>
              {showNewDir && (
                <div style={{ display: 'flex', gap: 8, padding: '8px 16px', borderBottom: '1px solid var(--border)', background: 'var(--surface)', flexShrink: 0 }}>
                  <input className="input" placeholder="Folder name..." value={newDirName} onChange={e => setNewDirName(e.target.value)} style={{ width: 220, height: 32, padding: '0 10px', fontSize: '0.8rem' }} onKeyDown={e => e.key === 'Enter' && handleMkdir()} />
                  <button className="btn btn-primary btn-sm" onClick={handleMkdir}>Create</button>
                  <button className="btn btn-ghost btn-sm" onClick={() => setShowNewDir(false)}>Cancel</button>
                </div>
              )}
              {browsePath && <div style={{ padding: '6px 16px', background: 'var(--surface-2)', borderBottom: '1px solid var(--border)', fontSize: '0.7rem', fontFamily: 'var(--font-mono)', color: 'var(--text-muted)', flexShrink: 0 }}>{browsePath}</div>}
              <div style={{ flex: 1, overflowY: 'auto' }}>
                {browseItems.length === 0 ? (
                  <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', height: 300, gap: 8, color: 'var(--text-muted)' }}>
                    <div style={{ fontSize: 36, opacity: 0.3 }}>📂</div>
                    <p>Enter a directory path above and click Go.</p>
                  </div>
                ) : (
                  browseItems.map((item, i) => {
                    const fileLike = { type: item.isDir ? 'folder' : item.type, name: item.name }
                    return (
                      <div key={i} style={{ display: 'grid', gridTemplateColumns: '40px 1fr 100px 140px', alignItems: 'center', padding: '8px 16px', borderBottom: '1px solid var(--border)', cursor: item.isDir ? 'pointer' : 'default', transition: 'background var(--tr)' }}
                        onMouseEnter={e => (e.currentTarget.style.background = 'var(--surface-2)')}
                        onMouseLeave={e => (e.currentTarget.style.background = 'transparent')}
                        onClick={() => item.isDir && browseDirectory(item.path)}
                      >
                        <FileMark file={fileLike} compact />
                        <div>
                          <div style={{ fontSize: '0.8rem', color: 'var(--text)', fontWeight: item.isDir ? 600 : 400 }}>{item.name}</div>
                          <div style={{ fontSize: '0.68rem', color: 'var(--text-muted)', fontFamily: 'var(--font-mono)' }}>{item.path}</div>
                        </div>
                        <div style={{ fontSize: '0.72rem', color: 'var(--text-muted)' }}>{item.isDir ? '—' : fmtSize(item.size)}</div>
                        <div style={{ fontSize: '0.72rem', color: 'var(--text-muted)' }}>{fmtDate(item.modified)}</div>
                      </div>
                    )
                  })
                )}
              </div>
            </div>
          )}

          {/* ── SEARCH TAB ── */}
          {tab === 'search' && (
            <div style={{ flex: 1, display: 'flex', flexDirection: 'column', overflow: 'hidden' }}>
              <div style={{ display: 'flex', gap: 8, padding: '12px 16px', borderBottom: '1px solid var(--border)', background: 'var(--bg)', flexShrink: 0, alignItems: 'center' }}>
                <div className="search-bar" style={{ flex: 1 }}>
                  <svg width={12} height={12} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth={2}><circle cx={11} cy={11} r={8}/><path d="m21 21-4.35-4.35"/></svg>
                  <input placeholder="Search filename across all uploaded files..." value={diskSearch} onChange={e => setDiskSearch(e.target.value)} onKeyDown={e => e.key === 'Enter' && handleDiskSearch()} />
                </div>
                <button className="btn btn-primary btn-sm" onClick={handleDiskSearch} disabled={diskSearching}>{diskSearching ? '⟳ Searching...' : 'Search'}</button>
              </div>
              <div style={{ flex: 1, overflowY: 'auto' }}>
                {diskResults.length === 0 ? (
                  <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', height: 300, gap: 8, color: 'var(--text-muted)' }}>
                    <div style={{ fontSize: 36, opacity: 0.3 }}>🔍</div>
                    <p>Type a filename and press Search. Searches all uploaded files.</p>
                    <p style={{ fontSize: '0.72rem' }}>Supports partial name matching.</p>
                  </div>
                ) : (
                  <>
                    <div style={{ padding: '8px 16px', borderBottom: '1px solid var(--border)', fontSize: '0.72rem', color: 'var(--text-muted)' }}>
                      Found {diskResults.length} files
                    </div>
                    {diskResults.map((item, i) => {
                      const fileLike = { type: item.type || '', name: item.name }
                      return (
                        <div key={i} style={{ display: 'grid', gridTemplateColumns: '40px 1fr 100px 140px', alignItems: 'center', padding: '8px 16px', borderBottom: '1px solid var(--border)', transition: 'background var(--tr)' }}
                          onMouseEnter={e => (e.currentTarget.style.background = 'var(--surface-2)')}
                          onMouseLeave={e => (e.currentTarget.style.background = 'transparent')}
                        >
                          <FileMark file={fileLike} compact />
                          <div>
                            <div style={{ fontSize: '0.8rem', color: 'var(--text)', fontWeight: 500 }}>{item.name}</div>
                            <div style={{ fontSize: '0.68rem', color: 'var(--text-muted)', fontFamily: 'var(--font-mono)' }}>{item.path}</div>
                          </div>
                          <div style={{ fontSize: '0.72rem', color: 'var(--text-muted)' }}>{fmtSize(item.size)}</div>
                          <div style={{ fontSize: '0.72rem', color: 'var(--text-muted)' }}>{fmtDate(item.modified)}</div>
                        </div>
                      )
                    })}
                  </>
                )}
              </div>
            </div>
          )}
        </main>
      </div>
    </div>
  )
}
