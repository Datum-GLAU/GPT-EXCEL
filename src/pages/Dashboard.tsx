import { useState, useEffect, useMemo, useRef } from 'react'
import { useNavigate } from 'react-router-dom'
import { useSelector } from 'react-redux'
import { RootState } from '../index'
import Sidebar from '../components/Sidebar'
import Header from '../components/Header'
import AIChatPanel from '../components/AIChatPanel'
import { api } from '../api'

const SERVER_BASE = 'http://localhost:3001'

const Sparkline = ({ data, color = 'var(--accent)' }: { data: number[]; color?: string }) => {
  const max = Math.max(...data), min = Math.min(...data)
  const norm = (v: number) => 1 - (v - min) / (max - min || 1)
  const w = 64, h = 28
  const pts = data.map((v, i) => `${(i / (data.length - 1)) * w},${norm(v) * h}`).join(' ')
  return (
    <svg width={w} height={h} style={{ overflow: 'visible' }}>
      <polyline points={pts} fill="none" stroke={color} strokeWidth={1.5} strokeLinejoin="round" strokeLinecap="round" />
      <circle cx={(data.length - 1) / (data.length - 1) * w} cy={norm(data[data.length - 1]) * h} r={2.5} fill={color} />
    </svg>
  )
}

type FileKind = 'excel' | 'document' | 'presentation' | 'pdf' | 'image' | 'text' | 'folder' | 'other'

const FILE_META: Record<FileKind, { label: string; short: string; bg: string; text: string; glow: string }> = {
  excel: { label: 'Excel', short: 'XL', bg: 'linear-gradient(145deg, #107c41, #0f5132)', text: '#dcfce7', glow: 'rgba(16,124,65,0.22)' },
  document: { label: 'Document', short: 'DOC', bg: 'linear-gradient(145deg, #185abd, #143a7b)', text: '#dbeafe', glow: 'rgba(24,90,189,0.22)' },
  presentation: { label: 'Presentation', short: 'PPT', bg: 'linear-gradient(145deg, #d24726, #8a2f19)', text: '#ffedd5', glow: 'rgba(210,71,38,0.22)' },
  pdf: { label: 'PDF', short: 'PDF', bg: 'linear-gradient(145deg, #dc2626, #7f1d1d)', text: '#fee2e2', glow: 'rgba(220,38,38,0.22)' },
  image: { label: 'Image', short: 'IMG', bg: 'linear-gradient(145deg, #7c3aed, #4c1d95)', text: '#ede9fe', glow: 'rgba(124,58,237,0.22)' },
  text: { label: 'Text', short: 'TXT', bg: 'linear-gradient(145deg, #475569, #1e293b)', text: '#e2e8f0', glow: 'rgba(71,85,105,0.22)' },
  folder: { label: 'Folder', short: 'DIR', bg: 'linear-gradient(145deg, #f59e0b, #b45309)', text: '#fef3c7', glow: 'rgba(245,158,11,0.22)' },
  other: { label: 'File', short: 'FILE', bg: 'linear-gradient(145deg, #6b7280, #374151)', text: '#f3f4f6', glow: 'rgba(107,114,128,0.22)' },
}

const KIND_SAMPLE_TYPE: Record<FileKind, string> = {
  excel: 'xlsx',
  document: 'docx',
  presentation: 'pptx',
  pdf: 'pdf',
  image: 'png',
  text: 'txt',
  folder: 'folder',
  other: 'file',
}

const getFileExtension = (file: any) => {
  const explicit = String(file?.type || '').toLowerCase()
  if (explicit) return explicit
  const match = String(file?.name || '').toLowerCase().match(/\.([a-z0-9]+)$/)
  return match?.[1] || ''
}

const inferFileKind = (file: any): FileKind => {
  const type = getFileExtension(file)
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

const getFileVisual = (file: any) => {
  const kind = inferFileKind(file)
  const ext = getFileExtension(file)
  const meta = FILE_META[kind]
  return {
    ...meta,
    kind,
    ext,
    short: kind === 'other' && ext ? ext.slice(0, 4).toUpperCase() : meta.short,
  }
}

const formatFileSize = (bytes: number) => {
  const value = Number(bytes || 0)
  if (value < 1024) return `${value} B`
  if (value < 1024 * 1024) return `${(value / 1024).toFixed(1)} KB`
  return `${(value / 1024 / 1024).toFixed(1)} MB`
}

const formatDateTime = (value?: string) => {
  if (!value) return '-'
  return new Date(value).toLocaleString()
}

const buildLogSheetRows = (files: any[]) => {
  const sorted = [...files].sort((left, right) => {
    const kindDiff = FILE_META[inferFileKind(left)].label.localeCompare(FILE_META[inferFileKind(right)].label)
    if (kindDiff) return kindDiff
    const extDiff = getFileExtension(left).localeCompare(getFileExtension(right))
    if (extDiff) return extDiff
    return new Date(right.uploadedAt || 0).getTime() - new Date(left.uploadedAt || 0).getTime()
  })

  const summary = new Map<string, { ext: string; kind: FileKind; count: number; bytes: number }>()
  sorted.forEach(file => {
    const kind = inferFileKind(file)
    const ext = getFileExtension(file) || '-'
    const key = `${kind}:${ext}`
    const bucket = summary.get(key) || { ext, kind, count: 0, bytes: 0 }
    bucket.count += 1
    bucket.bytes += Number(file.size || 0)
    summary.set(key, bucket)
  })

  return [
    ['File Type Log'],
    ['Generated At', new Date().toLocaleString()],
    ['Total Files', files.length],
    ['Type Groups', summary.size],
    [],
    ['Type Summary'],
    ['Extension', 'Category', 'Count', 'Total Size (KB)'],
    ...Array.from(summary.values())
      .sort((left, right) => right.count - left.count || FILE_META[left.kind].label.localeCompare(FILE_META[right.kind].label))
      .map(item => [item.ext === '-' ? '-' : `.${item.ext}`, FILE_META[item.kind].label, item.count, Number((item.bytes / 1024).toFixed(1))]),
    [],
    ['Detailed Log'],
    ['Extension', 'Category', 'File Name', 'Department', 'Size (KB)', 'Uploaded By', 'Uploaded At'],
    ...sorted.map(file => {
      const visual = getFileVisual(file)
      return [
        visual.ext ? `.${visual.ext}` : '-',
        visual.label,
        file.name || 'Untitled',
        file.department || 'general',
        Number((Number(file.size || 0) / 1024).toFixed(1)),
        file.uploadedBy || 'user',
        formatDateTime(file.uploadedAt),
      ]
    }),
  ]
}

function FileLogo({ file, compact = false }: { file: any; compact?: boolean }) {
  const visual = getFileVisual(file)
  const size = compact ? 38 : 46
  const imageUrl = visual.kind === 'image' && file?.url ? `${SERVER_BASE}${file.url}` : ''

  return (
    <div style={{ width: size, height: size, borderRadius: 12, background: visual.bg, color: visual.text, display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', boxShadow: `0 10px 24px ${visual.glow}`, border: '1px solid rgba(255,255,255,0.12)', position: 'relative', overflow: 'hidden', flexShrink: 0 }}>
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
          <div style={{ position: 'absolute', top: 0, right: 0, width: compact ? 12 : 16, height: compact ? 12 : 16, background: 'rgba(255,255,255,0.18)', clipPath: 'polygon(0 0, 100% 0, 100% 100%)' }} />
          <div style={{ fontSize: compact ? '0.72rem' : '0.88rem', fontWeight: 800, letterSpacing: '0.04em', position: 'relative' }}>{visual.short}</div>
          {!compact && <div style={{ fontSize: '0.46rem', textTransform: 'uppercase', letterSpacing: '0.08em', opacity: 0.85, position: 'relative' }}>{visual.label}</div>}
        </>
      )}
    </div>
  )
}

export default function Dashboard() {
  const navigate = useNavigate()
  const user = useSelector((s: RootState) => s.app.user)
  const userDepartment = (user as any)?.department as string | undefined
  const [sidebarOpen, setSidebarOpen] = useState(true)
  const [stats, setStats] = useState<any>(null)
  const [allFiles, setAllFiles] = useState<any[]>([])
  const [recentFiles, setRecentFiles] = useState<any[]>([])
  const [prompt, setPrompt] = useState('')
  const [generating, setGenerating] = useState(false)
  const [greeting, setGreeting] = useState('')
  const [backendOnline, setBackendOnline] = useState(false)
  const [dragOver, setDragOver] = useState(false)
  const [logSheetBusy, setLogSheetBusy] = useState(false)
  const [statusNote, setStatusNote] = useState('')
  const promptRef = useRef<HTMLTextAreaElement>(null)

  useEffect(() => {
    const h = new Date().getHours()
    setGreeting(h < 12 ? 'Good morning' : h < 17 ? 'Good afternoon' : 'Good evening')
    loadData()
  }, [])

  const loadData = async () => {
    try {
      await api.health()
      setBackendOnline(true)
      const [dashStats, files] = await Promise.all([api.getDashboardStats(), api.getFiles()])
      const nextFiles = Array.isArray(files) ? files : []
      setStats(dashStats)
      setAllFiles(nextFiles)
      setRecentFiles([...nextFiles].sort((left, right) => new Date(right.uploadedAt || 0).getTime() - new Date(left.uploadedAt || 0).getTime()).slice(0, 8))
      if (!nextFiles.length) {
        setStatusNote('Upload a file to start tracking file activity by type.')
      } else {
        const departments = new Set(nextFiles.map(file => file.department).filter(Boolean)).size
        setStatusNote(`${nextFiles.length} files synced across ${departments || 1} department${departments === 1 ? '' : 's'}.`)
      }
    } catch {
      setBackendOnline(false)
      setAllFiles([])
      setRecentFiles([])
      setStatusNote('Backend is offline. Start the server to load dashboard data.')
    }
  }

  const handleDrop = async (e: React.DragEvent) => {
    e.preventDefault(); setDragOver(false)
    const files = Array.from(e.dataTransfer.files).filter(f => !!f.name)
    if (!files.length) return
    try {
      await api.uploadFiles(files, userDepartment || 'general', user?.email || 'user')
      await loadData()
      setStatusNote(`Uploaded ${files.length} file${files.length === 1 ? '' : 's'} successfully.`)
      if (files.every(f => /\.(xlsx|xls|csv)$/i.test(f.name))) navigate('/excel')
    } catch {
      setStatusNote('Upload failed. Check that the backend is running and try again.')
    }
  }

  const handlePrompt = async () => {
    if (!prompt.trim()) return
    setGenerating(true)
    await new Promise(r => setTimeout(r, 600))
    setGenerating(false)
    navigate('/excel')
  }

  const fileTypeSummary = useMemo(() => {
    const buckets = new Map<FileKind, { kind: FileKind; count: number; bytes: number }>()
    allFiles.forEach(file => {
      const kind = inferFileKind(file)
      const bucket = buckets.get(kind) || { kind, count: 0, bytes: 0 }
      bucket.count += 1
      bucket.bytes += Number(file.size || 0)
      buckets.set(kind, bucket)
    })
    return Array.from(buckets.values()).sort((left, right) => right.count - left.count || FILE_META[left.kind].label.localeCompare(FILE_META[right.kind].label))
  }, [allFiles])

  const logPreviewRows = useMemo(() => {
    return [...allFiles]
      .sort((left, right) => {
        const kindDiff = FILE_META[inferFileKind(left)].label.localeCompare(FILE_META[inferFileKind(right)].label)
        if (kindDiff) return kindDiff
        return new Date(right.uploadedAt || 0).getTime() - new Date(left.uploadedAt || 0).getTime()
      })
      .slice(0, 8)
  }, [allFiles])

  const aiFileCount = useMemo(() => allFiles.filter(file => String(file.uploadedBy || '').toLowerCase() === 'ai').length, [allFiles])

  const openRecentFile = (file: any) => {
    if (inferFileKind(file) === 'excel') {
      localStorage.setItem('excel_autoload_file_id', file.id)
      navigate('/excel')
      return
    }
    navigate('/file-manager')
  }

  const handleCreateLogSheet = async () => {
    if (!allFiles.length) {
      setStatusNote('No files available yet. Upload something first.')
      return
    }
    setLogSheetBusy(true)
    try {
      const rows = buildLogSheetRows(allFiles)
      const filename = `file_type_log_${new Date().toISOString().slice(0, 10)}.xlsx`
      const response = await api.generateExcel(rows, filename, 'File Type Log', userDepartment || 'general')
      if (response?.file?.id) {
        localStorage.setItem('excel_autoload_file_id', response.file.id)
        setStatusNote('Created a file log workbook and opened it in Excel Sheet.')
        navigate('/excel')
      } else {
        setStatusNote('Created a file log workbook.')
      }
    } catch {
      setStatusNote('Could not create the file log workbook right now.')
    }
    setLogSheetBusy(false)
  }

  const METRICS = [
    { label: 'Total Files', value: stats?.totalFiles ?? allFiles.length ?? '-', change: 'tracked', up: true, sparkline: [2,4,3,6,5,7,8,7,9,10,11,Math.max(allFiles.length, 1)], color: 'var(--blue)' },
    { label: 'Known Types', value: fileTypeSummary.length || '-', change: 'active groups', up: true, sparkline: [1,1,2,2,3,3,4,4,5,6,6,Math.max(fileTypeSummary.length, 1)], color: 'var(--green)' },
    { label: 'AI Files', value: aiFileCount, change: 'generated', up: true, sparkline: [0,1,1,2,2,3,4,4,5,5,6,Math.max(aiFileCount, 1)], color: 'var(--purple)' },
    { label: 'Departments', value: stats?.departments?.length ?? '-', change: 'connected', up: true, sparkline: [1,2,2,3,3,4,4,5,5,6,6,stats?.departments?.length||1], color: 'var(--orange)' },
  ]

  const quickActions = [
    { icon: '⊞', label: 'Analyze Excel', color: 'var(--green)', path: '/excel' },
    { icon: '◱', label: 'File Manager', color: 'var(--blue)', path: '/file-manager' },
    { icon: '◬', label: 'Compare Files', color: 'var(--purple)', path: '/excel' },
    { icon: '◱', label: 'Generate Doc', color: 'var(--orange)', path: '/documents' },
    { icon: '⌘', label: 'Workflows', color: 'var(--teal)', path: '/workflow' },
    { icon: '⚙', label: 'Settings', color: 'var(--text-muted)', path: '/settings' },
  ]

  return (
    <div style={{ display: 'flex', flexDirection: 'column', height: '100vh', background: 'var(--bg)' }}>
      <Header toggleSidebar={() => setSidebarOpen(p => !p)} />
      <div style={{ display: 'flex', flex: 1, overflow: 'hidden' }}>
        <Sidebar isOpen={sidebarOpen} />
        <main style={{ flex: 1, overflowY: 'auto', background: 'var(--bg-2)', padding: '24px 28px' }}>

          {/* Backend status */}
          <div style={{ display: 'flex', alignItems: 'center', gap: 6, marginBottom: 16, fontSize: '0.7rem', fontFamily: 'var(--font-mono)' }}>
            <div style={{ width: 6, height: 6, borderRadius: '50%', background: backendOnline ? 'var(--green)' : 'var(--red)' }} />
            <span style={{ color: 'var(--text-muted)' }}>{backendOnline ? 'Backend online · localhost:3001' : 'Backend offline — run: cd server && npm start'}</span>
            {statusNote && <span style={{ color: 'var(--text-muted)' }}>· {statusNote}</span>}
          </div>

          {/* Greeting */}
          <div style={{ marginBottom: 20 }}>
            <div style={{ fontSize: '0.7rem', fontFamily: 'var(--font-mono)', color: 'var(--text-muted)', textTransform: 'uppercase', marginBottom: 4 }}>
              {new Date().toLocaleDateString('en-US', { weekday: 'long', month: 'long', day: 'numeric', year: 'numeric' })}
            </div>
            <h1 style={{ fontSize: '1.5rem', letterSpacing: -0.5, marginBottom: 4 }}>
              {greeting}, {user?.name?.split(' ')[0] || 'there'}.
            </h1>
            <p style={{ fontSize: '0.8rem', color: 'var(--text-secondary)' }}>
              {userDepartment && <span style={{ color: 'var(--accent)', fontWeight: 600 }}>{userDepartment}</span>}
              {userDepartment && ' · '}
              Analyze student data, compare sections, generate reports.
            </p>
          </div>

          {/* Drop zone prompt bar */}
          <div
            onDragOver={e => { e.preventDefault(); setDragOver(true) }}
            onDragLeave={() => setDragOver(false)}
            onDrop={handleDrop}
            style={{
              background: dragOver ? 'var(--accent-dim)' : 'var(--surface)',
              border: `1px ${dragOver ? 'solid' : 'solid'} ${dragOver ? 'var(--accent)' : 'var(--border)'}`,
              borderRadius: 10, overflow: 'hidden', marginBottom: 20,
              transition: 'all 0.2s'
            }}
          >
            <div style={{ display: 'flex', gap: 10, padding: '12px 14px' }}>
              <div style={{ width: 28, height: 28, background: 'var(--blue)', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 14, borderRadius: 6, flexShrink: 0, marginTop: 1 }}>⊞</div>
              <textarea
                ref={promptRef}
                className="input"
                value={prompt}
                onChange={e => setPrompt(e.target.value)}
                onKeyDown={e => { if (e.key === 'Enter' && (e.metaKey || e.ctrlKey)) handlePrompt() }}
                placeholder={dragOver ? 'Drop Excel files here to analyze...' : 'Ask anything or drag & drop Excel files... "Analyze Q4 results" or "Compare section A and B"'}
                style={{ flex: 1, border: 'none', background: 'transparent', resize: 'none', minHeight: 22, maxHeight: 100, boxShadow: 'none', fontSize: '0.875rem', padding: 0, lineHeight: 1.6 }}
                rows={1}
                onInput={e => { const el = e.currentTarget; el.style.height = 'auto'; el.style.height = Math.min(el.scrollHeight, 90) + 'px' }}
              />
            </div>
            <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '8px 14px', borderTop: '1px solid var(--border)', background: 'var(--bg-3)' }}>
              <div style={{ display: 'flex', gap: 6 }}>
                {['Analyze Excel', 'Top Students', 'Section Compare', 'Below 75%'].map(t => (
                  <button key={t} className="btn btn-xs btn-outline" onClick={() => { setPrompt(t); promptRef.current?.focus() }}>{t}</button>
                ))}
              </div>
              <button className="btn btn-primary btn-sm" onClick={handlePrompt} disabled={!prompt.trim() || generating} style={{ minWidth: 90 }}>
                {generating ? <><span className="spinner" style={{ width: 12, height: 12 }} /> Running</> : 'Go →'}
              </button>
            </div>
          </div>

          {/* Metrics */}
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4,1fr)', gap: 12, marginBottom: 20 }}>
            {METRICS.map((m, i) => (
              <div key={i} style={{ background: 'var(--surface)', border: '1px solid var(--border)', borderRadius: 8, padding: '12px 14px', transition: 'border-color 0.2s' }}
                onMouseEnter={e => (e.currentTarget.style.borderColor = m.color)}
                onMouseLeave={e => (e.currentTarget.style.borderColor = 'var(--border)')}
              >
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start' }}>
                  <div>
                    <div style={{ fontFamily: 'var(--font-display)', fontSize: '1.6rem', fontWeight: 800, letterSpacing: -0.5, color: m.color }}>{m.value}</div>
                    <div style={{ fontSize: '0.7rem', color: 'var(--text-muted)', marginTop: 2 }}>{m.label}</div>
                  </div>
                  <Sparkline data={m.sparkline} color={m.color} />
                </div>
                <div style={{ marginTop: 8, fontSize: '0.7rem', fontFamily: 'var(--font-mono)', color: m.color }}>● {m.change}</div>
              </div>
            ))}
          </div>

          {/* Quick actions + recent files */}
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 280px', gap: 16, marginBottom: 16 }}>
            {/* Recent files */}
            <div style={{ background: 'var(--surface)', border: '1px solid var(--border)', borderRadius: 10, overflow: 'hidden' }}>
              <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '10px 14px', borderBottom: '1px solid var(--border)' }}>
                <span style={{ fontSize: '0.8rem', fontWeight: 600, color: 'var(--text)' }}>Recent Files</span>
                <button className="btn btn-primary btn-sm" onClick={() => navigate('/file-manager')}>View All</button>
              </div>
              {recentFiles.length === 0 ? (
                <div style={{ padding: '32px', textAlign: 'center', color: 'var(--text-muted)', fontSize: '0.8rem' }}>
                  <div style={{ fontSize: 32, marginBottom: 8, opacity: 0.4 }}>⊞</div>
                  No files yet. Upload Excel files to get started.
                  <br /><button className="btn btn-primary btn-sm" style={{ marginTop: 12 }} onClick={() => navigate('/file-manager')}>Upload Files</button>
                </div>
              ) : recentFiles.map(f => {
                const visual = getFileVisual(f)
                return (
                  <div key={f.id} style={{ display: 'flex', alignItems: 'center', gap: 10, padding: '9px 14px', borderBottom: '1px solid var(--border)', cursor: 'pointer', transition: 'background var(--tr)' }}
                    onMouseEnter={e => (e.currentTarget.style.background = 'var(--surface-2)')}
                    onMouseLeave={e => (e.currentTarget.style.background = 'transparent')}
                    onClick={() => openRecentFile(f)}
                  >
                    <FileLogo file={f} compact />
                    <div style={{ flex: 1, minWidth: 0 }}>
                      <div style={{ fontSize: '0.8rem', fontWeight: 500, color: 'var(--text)', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{f.name}</div>
                      <div style={{ fontSize: '0.7rem', color: 'var(--text-muted)' }}>{visual.label} · {f.department || 'general'} · {formatFileSize(f.size)}</div>
                      <div style={{ fontSize: '0.66rem', color: 'var(--text-muted)', marginTop: 2 }}>Uploaded by {f.uploadedBy || 'user'} · {formatDateTime(f.uploadedAt)}</div>
                    </div>
                    <button className="btn btn-xs btn-outline" onClick={e => { e.stopPropagation(); openRecentFile(f) }}>{visual.kind === 'excel' ? 'Open' : 'Manage'}</button>
                  </div>
                )
              })}
            </div>

            {/* Quick actions */}
            <div style={{ background: 'var(--surface)', border: '1px solid var(--border)', borderRadius: 10, overflow: 'hidden' }}>
              <div style={{ padding: '10px 14px', borderBottom: '1px solid var(--border)', fontSize: '0.8rem', fontWeight: 600, color: 'var(--text)' }}>Quick Actions</div>
              <div style={{ padding: 10, display: 'flex', flexDirection: 'column', gap: 6 }}>
                {quickActions.map(a => (
                  <button key={a.label} onClick={() => navigate(a.path)} style={{ display: 'flex', alignItems: 'center', gap: 10, padding: '9px 12px', background: 'var(--surface-2)', border: '1px solid var(--border)', borderRadius: 8, cursor: 'pointer', transition: 'all 0.15s', textAlign: 'left' }}
                    onMouseEnter={e => { const el = e.currentTarget as HTMLElement; el.style.borderColor = a.color; el.style.background = `${a.color}10` }}
                    onMouseLeave={e => { const el = e.currentTarget as HTMLElement; el.style.borderColor = 'var(--border)'; el.style.background = 'var(--surface-2)' }}
                  >
                    <span style={{ color: a.color, fontSize: 16, width: 20, textAlign: 'center' }}>{a.icon}</span>
                    <span style={{ fontSize: '0.8rem', color: 'var(--text)', fontWeight: 500 }}>{a.label}</span>
                    <span style={{ marginLeft: 'auto', color: 'var(--text-muted)', fontSize: 12 }}>→</span>
                  </button>
                ))}
              </div>
            </div>
          </div>

          <div style={{ display: 'grid', gridTemplateColumns: '1fr 320px', gap: 16, marginBottom: 16 }}>
            <div style={{ background: 'var(--surface)', border: '1px solid var(--border)', borderRadius: 10, overflow: 'hidden' }}>
              <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '10px 14px', borderBottom: '1px solid var(--border)' }}>
                <div>
                  <div style={{ fontSize: '0.82rem', fontWeight: 600, color: 'var(--text)' }}>File Log Sheet Preview</div>
                  <div style={{ fontSize: '0.7rem', color: 'var(--text-muted)', marginTop: 2 }}>Preview of the rows that will be grouped into an Excel workbook by file type.</div>
                </div>
                <button className="btn btn-primary btn-sm" onClick={handleCreateLogSheet} disabled={logSheetBusy || !allFiles.length}>
                  {logSheetBusy ? 'Creating...' : 'Open as Excel'}
                </button>
              </div>
              {logPreviewRows.length === 0 ? (
                <div style={{ padding: '24px 14px', color: 'var(--text-muted)', fontSize: '0.78rem' }}>No file activity yet. Upload files to build the log sheet.</div>
              ) : (
                <div style={{ overflowX: 'auto' }}>
                  <table style={{ width: '100%', borderCollapse: 'collapse', minWidth: 640, fontSize: '0.74rem' }}>
                    <thead>
                      <tr style={{ background: 'var(--bg-3)' }}>
                        <th style={{ padding: '9px 12px', textAlign: 'left', color: 'var(--text-muted)', borderBottom: '1px solid var(--border)' }}>Type</th>
                        <th style={{ padding: '9px 12px', textAlign: 'left', color: 'var(--text-muted)', borderBottom: '1px solid var(--border)' }}>Category</th>
                        <th style={{ padding: '9px 12px', textAlign: 'left', color: 'var(--text-muted)', borderBottom: '1px solid var(--border)' }}>File Name</th>
                        <th style={{ padding: '9px 12px', textAlign: 'left', color: 'var(--text-muted)', borderBottom: '1px solid var(--border)' }}>Department</th>
                        <th style={{ padding: '9px 12px', textAlign: 'left', color: 'var(--text-muted)', borderBottom: '1px solid var(--border)' }}>Size</th>
                        <th style={{ padding: '9px 12px', textAlign: 'left', color: 'var(--text-muted)', borderBottom: '1px solid var(--border)' }}>Uploaded By</th>
                      </tr>
                    </thead>
                    <tbody>
                      {logPreviewRows.map(file => {
                        const visual = getFileVisual(file)
                        return (
                          <tr key={file.id}>
                            <td style={{ padding: '9px 12px', borderBottom: '1px solid var(--border)' }}>{visual.ext ? `.${visual.ext}` : '-'}</td>
                            <td style={{ padding: '9px 12px', borderBottom: '1px solid var(--border)' }}>{visual.label}</td>
                            <td style={{ padding: '9px 12px', borderBottom: '1px solid var(--border)', color: 'var(--text)' }}>{file.name}</td>
                            <td style={{ padding: '9px 12px', borderBottom: '1px solid var(--border)' }}>{file.department || 'general'}</td>
                            <td style={{ padding: '9px 12px', borderBottom: '1px solid var(--border)' }}>{formatFileSize(file.size)}</td>
                            <td style={{ padding: '9px 12px', borderBottom: '1px solid var(--border)' }}>{file.uploadedBy || 'user'}</td>
                          </tr>
                        )
                      })}
                    </tbody>
                  </table>
                </div>
              )}
            </div>

            <div style={{ display: 'grid', gap: 16 }}>
              <div style={{ background: 'var(--surface)', border: '1px solid var(--border)', borderRadius: 10, overflow: 'hidden' }}>
                <div style={{ padding: '10px 14px', borderBottom: '1px solid var(--border)', fontSize: '0.82rem', fontWeight: 600, color: 'var(--text)' }}>File Types</div>
                <div style={{ padding: 10, display: 'flex', flexDirection: 'column', gap: 8 }}>
                  {fileTypeSummary.length === 0 ? (
                    <div style={{ fontSize: '0.76rem', color: 'var(--text-muted)' }}>No file types available yet.</div>
                  ) : (
                    fileTypeSummary.slice(0, 6).map(item => {
                      const sample = { type: KIND_SAMPLE_TYPE[item.kind], name: item.kind }
                      const visual = getFileVisual(sample)
                      return (
                        <div key={item.kind} style={{ display: 'grid', gridTemplateColumns: '40px 1fr auto', gap: 10, alignItems: 'center', padding: '9px 10px', borderRadius: 8, background: 'var(--surface-2)', border: '1px solid var(--border)' }}>
                          <FileLogo file={sample} compact />
                          <div>
                            <div style={{ fontSize: '0.78rem', color: 'var(--text)', fontWeight: 600 }}>{visual.label}</div>
                            <div style={{ fontSize: '0.68rem', color: 'var(--text-muted)', marginTop: 2 }}>{formatFileSize(item.bytes)} total</div>
                          </div>
                          <div style={{ fontSize: '0.8rem', color: 'var(--text)', fontWeight: 700 }}>{item.count}</div>
                        </div>
                      )
                    })
                  )}
                </div>
              </div>

              <div style={{ background: 'var(--surface)', border: '1px solid var(--border)', borderRadius: 10, padding: 14 }}>
                <div style={{ fontSize: '0.82rem', fontWeight: 600, color: 'var(--text)', marginBottom: 6 }}>Workbook Shortcut</div>
                <div style={{ fontSize: '0.74rem', color: 'var(--text-muted)', lineHeight: 1.6, marginBottom: 12 }}>
                  Create a workbook that summarizes uploaded files by extension and category, then open it directly in Excel Sheet.
                </div>
                <button className="btn btn-primary btn-sm" onClick={handleCreateLogSheet} disabled={logSheetBusy || !allFiles.length} style={{ width: '100%', justifyContent: 'center' }}>
                  {logSheetBusy ? 'Building workbook...' : 'Create File Log Workbook'}
                </button>
              </div>
            </div>
          </div>

          {/* Department overview */}
          {stats?.departments && stats.departments.length > 0 && (
            <div style={{ background: 'var(--surface)', border: '1px solid var(--border)', borderRadius: 10, padding: '14px', marginBottom: 16 }}>
              <div style={{ fontSize: '0.8rem', fontWeight: 600, color: 'var(--text)', marginBottom: 12 }}>Departments</div>
              <div style={{ display: 'flex', flexWrap: 'wrap', gap: 8 }}>
                {stats.departments.map((dept: string, i: number) => {
                  const colors = ['var(--blue)', 'var(--green)', 'var(--purple)', 'var(--orange)', 'var(--teal)']
                  const c = colors[i % colors.length]
                  return (
                    <div key={dept} style={{ padding: '6px 14px', background: `${c}12`, border: `1px solid ${c}30`, borderRadius: 20, fontSize: '0.78rem', color: c, fontWeight: 500 }}>{dept}</div>
                  )
                })}
              </div>
            </div>
          )}
        </main>
        <AIChatPanel initialCollapsed />
      </div>
    </div>
  )
}
