import { useState, useRef, useEffect, useMemo } from 'react'
import { useSelector } from 'react-redux'
import { RootState } from '../index'
import AIChatPanel from '../components/AIChatPanel'
import Sidebar from '../components/Sidebar'
import Header from '../components/Header'
import { api } from '../api'

type MainTab = 'files' | 'sheet' | 'analyze' | 'charts' | 'create'

interface CellState {
  value: string
  editing: boolean
  bold: boolean
  align: 'left' | 'center' | 'right'
}

interface AiPreviewState {
  kind: 'rows' | 'chart'
  title: string
  description: string
  rows?: string[][]
  chartType?: string
  chartDataKey?: string
  createdAt: number
}

interface AiHistoryState {
  kind: 'rows' | 'chart'
  grid: CellState[][]
  gridCols: number
  colWidths: number[]
  activeTab: MainTab
  chartType: string
  chartDataKey: string
}

interface WorkspaceIntent {
  tab: 'analyze' | 'charts' | 'create'
  prompt: string
  questions: string[]
}

const COL_LETTERS = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.split('')

function FileBadge({ file }: { file: any }) {
  return (
    <div style={{ width: 36, height: 36, borderRadius: 10, background: 'linear-gradient(145deg, #107c41, #0f5132)', color: '#dcfce7', display: 'flex', alignItems: 'center', justifyContent: 'center', boxShadow: '0 10px 24px rgba(16,124,65,0.22)', border: '1px solid rgba(255,255,255,0.12)', position: 'relative', overflow: 'hidden', flexShrink: 0 }}>
      <div style={{ position: 'absolute', inset: 0, background: 'linear-gradient(180deg, rgba(255,255,255,0.18), transparent 55%)' }} />
      <div style={{ position: 'absolute', top: 0, right: 0, width: 12, height: 12, background: 'rgba(255,255,255,0.18)', clipPath: 'polygon(0 0, 100% 0, 100% 100%)' }} />
      <div style={{ fontSize: '0.62rem', fontWeight: 800, letterSpacing: '0.04em', position: 'relative' }}>XL</div>
    </div>
  )
}

export default function ExcelSheet() {
  const user = useSelector((s: RootState) => s.app.user)
  const [sidebarOpen, setSidebarOpen] = useState(true)
  const [activeTab, setActiveTab] = useState<MainTab>('files')
  const [backendOnline, setBackendOnline] = useState(false)

  // Files
  const [files, setFiles] = useState<any[]>([])
  const [filesLoading, setFilesLoading] = useState(false)
  const [uploading, setUploading] = useState(false)
  const [fileSearch, setFileSearch] = useState('')
  const [savingSheet, setSavingSheet] = useState(false)
  const fileRef = useRef<HTMLInputElement>(null)

  // Sheet state
  const [currentFile, setCurrentFile] = useState<any>(null)
  const [sheetNames, setSheetNames] = useState<string[]>([])
  const [activeSheetIdx, setActiveSheetIdx] = useState(0)
  const [grid, setGrid] = useState<CellState[][]>([])
  const [gridCols, setGridCols] = useState(0)
  const [selectedCell, setSelectedCell] = useState<{ r: number; c: number } | null>(null)
  const [formulaBarVal, setFormulaBarVal] = useState('')
  const [isEditing, setIsEditing] = useState(false)
  const [sheetLoading, setSheetLoading] = useState(false)
  const [colWidths, setColWidths] = useState<number[]>([])
  const editInputRef = useRef<HTMLInputElement>(null)

  // Analysis
  const [stats, setStats] = useState<any>(null)
  const [cmdInput, setCmdInput] = useState('')
  const [cmdLoading, setCmdLoading] = useState(false)
  const [cmdResult, setCmdResult] = useState<any>(null)
  const [dataSearch, setDataSearch] = useState('')
  const [filteredStudents, setFilteredStudents] = useState<any[]>([])

  // Charts
  const [chartType, setChartType] = useState('bar')
  const [chartDataKey, setChartDataKey] = useState('section')
  const [aiPreview, setAiPreview] = useState<AiPreviewState | null>(null)
  const [aiHistory, setAiHistory] = useState<AiHistoryState | null>(null)
  const [previewNote, setPreviewNote] = useState('')
  const [aiStatus, setAiStatus] = useState<{ tone: 'processing' | 'ready' | 'applied'; text: string } | null>(null)
  const [workspaceIntent, setWorkspaceIntent] = useState<WorkspaceIntent | null>(null)

  // Create
  const [createPrompt, setCreatePrompt] = useState('')
  const [creating, setCreating] = useState(false)

  // Toast
  const [toast, setToast] = useState<{ msg: string; type: 'ok' | 'err' } | null>(null)
  const toastTimeoutRef = useRef<ReturnType<typeof setTimeout> | null>(null)

  const showToast = (msg: string, type: 'ok' | 'err' = 'ok') => {
    if (toastTimeoutRef.current) clearTimeout(toastTimeoutRef.current)
    setToast({ msg, type })
    toastTimeoutRef.current = setTimeout(() => {
      setToast(null)
      toastTimeoutRef.current = null
    }, 3000)
  }

  useEffect(() => {
    const checkBackend = async () => {
      try {
        await api.health()
        setBackendOnline(true)
      } catch {
        setBackendOnline(false)
      }
    }
    checkBackend()
    loadFiles()
    return () => {
      if (toastTimeoutRef.current) clearTimeout(toastTimeoutRef.current)
    }
  }, [])

  useEffect(() => {
    const pendingId = localStorage.getItem('excel_autoload_file_id')
    if (!pendingId || !files.length || currentFile?.id === pendingId) return
    const target = files.find(file => file.id === pendingId)
    if (!target) return
    localStorage.removeItem('excel_autoload_file_id')
    openFile(target)
  }, [files, currentFile])

  const loadFiles = async () => {
    setFilesLoading(true)
    try {
      const params: any = {}
      if (fileSearch && fileSearch.trim()) params.search = fileSearch.trim()
      const res = await api.getFiles(params)
      setFiles((res || []).filter((file: any) => ['xlsx', 'xls', 'csv'].includes(String(file?.type || '').toLowerCase())))
      console.log('Files loaded:', res)
    } catch (e) { 
      console.error('Load files error:', e)
      setFiles([]) 
    }
    setFilesLoading(false)
  }

  const handleUpload = async (fileList: File[]) => {
    if (!fileList.length) return
    setUploading(true)
    try {
      await api.uploadFiles(fileList, user?.department || 'general', user?.email || 'guest@demo.com')
      await loadFiles()
      showToast(`Uploaded ${fileList.length} file(s)`)
    } catch (e: any) { 
      console.error('Upload error:', e)
      showToast(e.message || 'Upload failed', 'err') 
    }
    setUploading(false)
  }

  const handleDeleteFile = async (id: string) => {
    if (!confirm('Delete this file?')) return
    try {
      await api.deleteFile(id)
      await loadFiles()
      if (currentFile?.id === id) { 
        setCurrentFile(null); setGrid([]); setStats(null) 
      }
      showToast('File deleted')
    } catch { 
      showToast('Delete failed', 'err') 
    }
  }

  const handleOpenGeneratedFile = async (file: any, rows?: any[][]) => {
    await loadFiles()
    setCurrentFile(file)
    if (rows?.length) buildGrid(rows)
    setSheetNames(['Sheet1'])
    setStats(null)
    setActiveSheetIdx(0)
    setActiveTab('sheet')
    setAiStatus({ tone: 'ready', text: `Opened ${file.name} from AI chat. You can edit cells, save the workbook, or ask for more changes.` })
    showToast(`Opened ${file.name}`)
    try {
      await openFile(file)
    } catch {}
  }

  const openFile = async (file: any) => {
    setSheetLoading(true)
    try {
      const res = await api.analyzeFile(file.id)
      console.log('File data:', res)
      setCurrentFile(file)
      setSheetNames(res.sheetNames || [])
      setStats(res.sheets || {})
      
      // Build grid from rawData
      const sheetKey = res.sheetNames?.[0]
      if (sheetKey && res.rawData?.[sheetKey]) {
        const rawRows = res.rawData[sheetKey].rawRows || []
        buildGrid(rawRows)
      } else {
        // Try alternative approach
        const firstSheet = Object.values(res.rawData || {})[0] as any
        if (firstSheet?.rawRows) {
          buildGrid(firstSheet.rawRows)
        }
      }
      
      setActiveSheetIdx(0)
      setActiveTab('sheet')
      showToast(`Opened ${file.name}`)
    } catch (e: any) { 
      console.error('Open file error:', e)
      if (currentFile?.id === file.id) setCurrentFile(null)
      showToast(e.message || 'Failed to open', 'err') 
    }
    setSheetLoading(false)
  }

  const saveCurrentSheet = async () => {
    if (!currentFile || !grid.length) return
    setSavingSheet(true)
    try {
      const rows = grid.map(row => row.map(cell => ({ v: cell.value })))
      await api.saveExcel(currentFile.id, rows, sheetNames[activeSheetIdx] || undefined)
      showToast('Workbook saved')
    } catch (e: any) {
      showToast(e.message || 'Save failed', 'err')
    }
    setSavingSheet(false)
  }

  const buildGrid = (rawRows: any[][]) => {
    if (!rawRows?.length) { 
      setGrid([]); 
      setGridCols(0); 
      return 
    }
    
    const maxCols = Math.max(...rawRows.map(r => r.length))
    const normalizedRows = rawRows.map(r => {
      const padded = [...r]
      while (padded.length < maxCols) padded.push('')
      return padded
    })
    
    const g: CellState[][] = normalizedRows.map(row =>
      row.map(cell => ({
        value: cell === null || cell === undefined ? '' : String(cell),
        editing: false, 
        bold: false, 
        align: 'left'
      }))
    )
    
    setGrid(g)
    setGridCols(maxCols)
    
    // Auto column widths
    const widths = Array(maxCols).fill(0).map((_, ci) => {
      const maxLen = normalizedRows.reduce((m, r) => Math.max(m, String(r[ci] || '').length), 0)
      return Math.min(Math.max(maxLen * 7 + 20, 70), 200)
    })
    setColWidths(widths)
    setSelectedCell(null)
    setFormulaBarVal('')
    setCmdResult(null)
  }

  const handleAIGridUpdate = (rows: string[][], desc: string) => {
    if (!rows?.length) return
    setAiPreview({
      kind: 'rows',
      title: 'AI sheet preview ready',
      description: desc || 'AI prepared a result table for review before applying it to the sheet.',
      rows,
      createdAt: Date.now()
    })
    setAiStatus({ tone: 'ready', text: 'AI finished preparing a sheet preview. Review it, then apply or discard.' })
    setActiveTab('sheet')
    showToast('AI preview ready')
  }

  const handleAIChartRequest = ({ chartType: nextType, chartDataKey: nextKey, title }: { chartType: string, chartDataKey: string, title: string }) => {
    setAiStatus({ tone: 'processing', text: 'AI is preparing a chart preview from the current workbook.' })
    setAiPreview({
      kind: 'chart',
      title: title || 'AI chart preview ready',
      description: `Prepared a ${nextType} chart using ${nextKey} insights from the current sheet.`,
      chartType: nextType,
      chartDataKey: nextKey,
      createdAt: Date.now()
    })
    setAiStatus({ tone: 'ready', text: 'AI chart preview is ready. Apply it to open the chart workspace, or discard it.' })
    setActiveTab('charts')
    showToast(title || 'Chart preview ready')
  }

  const handleWorkspaceIntent = (intent: WorkspaceIntent) => {
    setWorkspaceIntent(intent)
    setActiveTab(intent.tab)
    if (intent.tab === 'analyze') {
      setCmdInput(intent.prompt)
      setAiStatus({ tone: 'ready', text: 'Analysis workspace opened from chat. You can run the prompt below or pick a follow-up question.' })
    } else if (intent.tab === 'create') {
      setCreatePrompt(intent.prompt)
      setAiStatus({ tone: 'ready', text: 'Create workspace opened from chat. Refine the instructions before generating the sheet.' })
    } else {
      setAiStatus({ tone: 'ready', text: 'Chart workspace opened from chat. Choose the visual style and data slice, then apply the preview.' })
    }
  }

  useEffect(() => {
    const saved = sessionStorage.getItem('excel_workspace_intent')
    if (!saved) return
    try {
      const parsed = JSON.parse(saved)
      if (parsed?.tab && parsed?.prompt) handleWorkspaceIntent(parsed)
    } catch {}
    sessionStorage.removeItem('excel_workspace_intent')
  }, [])

  const applyIntentQuestion = (question: string) => {
    if (!workspaceIntent) return
    if (workspaceIntent.tab === 'create') {
      setCreatePrompt(prev => prev ? `${prev}\n${question}` : question)
      return
    }
    if (workspaceIntent.tab === 'analyze') {
      setCmdInput(question)
      return
    }
    const lower = question.toLowerCase()
    if (lower.includes('line')) setChartType('line')
    else if (lower.includes('pie')) setChartType('pie')
    else if (lower.includes('bar')) setChartType('bar')
    if (lower.includes('subject')) setChartDataKey('subject')
    else if (lower.includes('grade')) setChartDataKey('grade')
    else if (lower.includes('score')) setChartDataKey('score')
    else if (lower.includes('section')) setChartDataKey('section')
  }

  const applyAiPreview = () => {
    if (!aiPreview) return
    setAiHistory({
      kind: aiPreview.kind,
      grid: grid.map(row => row.map(cell => ({ ...cell }))),
      gridCols,
      colWidths: [...colWidths],
      activeTab,
      chartType,
      chartDataKey
    })
    if (aiPreview.kind === 'rows' && aiPreview.rows?.length) {
      buildGrid(aiPreview.rows)
      setActiveTab('sheet')
      showToast(aiPreview.description || 'AI changes applied')
    } else if (aiPreview.kind === 'chart') {
      setChartType(aiPreview.chartType || 'bar')
      setChartDataKey(aiPreview.chartDataKey || 'section')
      setActiveTab('charts')
      showToast(aiPreview.title || 'Chart applied')
    }
    setAiStatus({ tone: 'applied', text: 'Preview applied. You can undo the last AI action if you want to roll back.' })
    setAiPreview(null)
    setPreviewNote('')
  }

  const discardAiPreview = () => {
    setAiPreview(null)
    setPreviewNote('')
    setAiStatus({ tone: 'ready', text: 'AI preview discarded. Ask for another version any time.' })
    showToast('AI preview discarded')
  }

  const undoAiApply = () => {
    if (!aiHistory) return
    setGrid(aiHistory.grid.map(row => row.map(cell => ({ ...cell }))))
    setGridCols(aiHistory.gridCols)
    setColWidths([...aiHistory.colWidths])
    setChartType(aiHistory.chartType)
    setChartDataKey(aiHistory.chartDataKey)
    setActiveTab(aiHistory.activeTab)
    setSelectedCell(null)
    setFormulaBarVal('')
    setAiHistory(null)
    setAiStatus({ tone: 'ready', text: 'Last AI-applied change was undone.' })
    showToast('Undid last AI change')
  }

  const getCellAddr = (r: number, c: number) => {
    let colStr = ''
    let col = c
    while (col >= 0) {
      colStr = String.fromCharCode(65 + (col % 26)) + colStr
      col = Math.floor(col / 26) - 1
    }
    return `${colStr}${r + 1}`
  }

  const selectCell = (r: number, c: number) => {
    setSelectedCell({ r, c })
    setFormulaBarVal(grid[r]?.[c]?.value || '')
  }

  const commitEdit = (r: number, c: number, val: string) => {
    const newGrid = grid.map(row => row.map(cell => ({ ...cell })))
    if (newGrid[r]?.[c]) {
      newGrid[r][c].value = val
      newGrid[r][c].editing = false
    }
    setGrid(newGrid)
    setIsEditing(false)
    setFormulaBarVal(val)
  }

  const startEdit = (r: number, c: number) => {
    setIsEditing(true)
    setSelectedCell({ r, c })
    setTimeout(() => editInputRef.current?.focus(), 10)
  }

  const handleKeyDown = (e: React.KeyboardEvent, r: number, c: number) => {
    if (e.key === 'Enter') { 
      commitEdit(r, c, (e.target as HTMLInputElement).value)
      if (r + 1 < grid.length) selectCell(r + 1, c)
      e.preventDefault() 
    } else if (e.key === 'Tab') {
      commitEdit(r, c, (e.target as HTMLInputElement).value)
      if (c + 1 < gridCols) selectCell(r, c + 1)
      e.preventDefault()
    } else if (e.key === 'Escape') { 
      setIsEditing(false) 
    }
  }

  const handleGridKeyDown = (e: React.KeyboardEvent) => {
    if (!selectedCell) return
    const { r, c } = selectedCell
    if (e.key === 'ArrowUp' && r > 0) { selectCell(r - 1, c); e.preventDefault() }
    else if (e.key === 'ArrowDown' && r + 1 < grid.length) { selectCell(r + 1, c); e.preventDefault() }
    else if (e.key === 'ArrowLeft' && c > 0) { selectCell(r, c - 1); e.preventDefault() }
    else if (e.key === 'ArrowRight' && c + 1 < gridCols) { selectCell(r, c + 1); e.preventDefault() }
    else if (e.key === 'Delete') { commitEdit(r, c, ''); e.preventDefault() }
    else if (e.key === 'Enter' || e.key === 'F2') { startEdit(r, c); e.preventDefault() }
  }

  const runCommand = async () => {
    if (!cmdInput.trim() || !currentFile) return
    setCmdLoading(true)
    setCmdResult(null)
    try {
      const res = await api.excelCommand(currentFile.id, cmdInput)
      setCmdResult(res)
    } catch (e: any) { 
      showToast(e.message || 'Command failed', 'err') 
    }
    setCmdLoading(false)
  }

  useEffect(() => {
    if (!dataSearch || !stats) { 
      setFilteredStudents([])
      return 
    }
    const sheetStats = stats[sheetNames[activeSheetIdx]]
    if (!sheetStats?.students) return
    const q = dataSearch.toLowerCase()
    setFilteredStudents(sheetStats.students.filter((s: any) =>
      s.name?.toLowerCase().includes(q) || 
      s.roll?.toLowerCase().includes(q) ||
      s.section?.toLowerCase().includes(q)
    ))
  }, [dataSearch, stats, activeSheetIdx, sheetNames])

  const getChartData = () => {
    const sh = stats?.[sheetNames[activeSheetIdx]]
    if (!sh) return []
    if (chartDataKey === 'section') return sh.sectionStats || []
    if (chartDataKey === 'subject') return sh.subjectStats || []
    if (chartDataKey === 'grade') return sh.gradeDistribution || []
    if (chartDataKey === 'score') return sh.scoreDistribution || []
    return sh.sectionStats || []
  }

  const chartRows = getChartData()
  const getChartLabel = (row: any) => row.section || row.subject || row.grade || row.range || row.name || 'Item'
  const getChartValue = (row: any) => {
    const value = row.avg ?? row.count ?? row.passRate ?? row.total ?? 0
    return Number.isFinite(Number(value)) ? Number(value) : 0
  }

  const handleCreate = async () => {
    if (!createPrompt.trim()) return
    setCreating(true)
    try {
      const template = [
        { Name: 'Student 1', Roll: '001', Section: 'A', Maths: 85, Physics: 78, Chemistry: 90, Attendance: 92 },
        { Name: 'Student 2', Roll: '002', Section: 'A', Maths: 72, Physics: 65, Chemistry: 70, Attendance: 88 },
        { Name: 'Student 3', Roll: '003', Section: 'B', Maths: 90, Physics: 88, Chemistry: 85, Attendance: 95 },
        { Name: 'Student 4', Roll: '004', Section: 'B', Maths: 55, Physics: 60, Chemistry: 58, Attendance: 70 },
        { Name: 'Student 5', Roll: '005', Section: 'A', Maths: 95, Physics: 92, Chemistry: 88, Attendance: 98 },
      ]
      const fname = `${createPrompt.slice(0, 30).replace(/[^a-zA-Z0-9]/g, '_')}.xlsx`
      await api.generateExcel(template, fname, 'Sheet1', user?.department || 'general')
      await loadFiles()
      setCreatePrompt('')
      setActiveTab('files')
      showToast('File created — open it from My Files')
    } catch (e: any) { 
      showToast(e.message || 'Create failed', 'err') 
    }
    setCreating(false)
  }

  const curStats = useMemo(() => stats?.[sheetNames[activeSheetIdx]], [stats, sheetNames, activeSheetIdx])
  const rawGridValues = useMemo(() => grid.map(row => row.map(cell => cell.value)), [grid])

  return (
    <div style={{ display: 'flex', flexDirection: 'column', height: '100vh', background: 'var(--bg)' }}>
      <Header toggleSidebar={() => setSidebarOpen(p => !p)} />

      {toast && (
        <div style={{ position: 'fixed', top: 60, right: 20, zIndex: 9999, padding: '10px 18px', borderRadius: 10, background: toast.type === 'ok' ? 'rgba(34,197,94,0.15)' : 'rgba(239,68,68,0.15)', border: `1px solid ${toast.type === 'ok' ? 'rgba(34,197,94,0.4)' : 'rgba(239,68,68,0.4)'}`, color: toast.type === 'ok' ? 'var(--green)' : 'var(--red)', fontSize: '0.8rem' }}>
          {toast.type === 'ok' ? '✓' : '⚠'} {toast.msg}
        </div>
      )}

      <div style={{ display: 'flex', flex: 1, overflow: 'hidden' }}>
        <Sidebar isOpen={sidebarOpen} />
        <main style={{ flex: 1, display: 'flex', flexDirection: 'column', overflow: 'hidden', background: 'var(--bg-2)' }}>

          {/* Status Bar */}
          <div style={{ display: 'flex', alignItems: 'center', gap: 6, padding: '0 12px', height: 36, borderBottom: '1px solid var(--border)', background: 'var(--bg)', fontSize: '0.7rem' }}>
            <div style={{ width: 8, height: 8, borderRadius: '50%', background: backendOnline ? '#22c55e' : '#ef4444' }} />
            <span style={{ color: 'var(--text-muted)' }}>{backendOnline ? 'Backend connected' : 'Backend offline'}</span>
            <div style={{ marginLeft: 'auto', display: 'flex', gap: 8 }}>
              {currentFile && (
                <button className="btn btn-outline btn-sm" onClick={saveCurrentSheet} disabled={savingSheet}>
                  {savingSheet ? 'Saving...' : 'Save Workbook'}
                </button>
              )}
              <button className="btn btn-primary btn-sm" onClick={() => fileRef.current?.click()} disabled={uploading}>
                {uploading ? '⟳ Uploading...' : '+ Upload Excel'}
              </button>
              <input ref={fileRef} type="file" multiple accept=".xlsx,.xls,.csv" style={{ display: 'none' }} 
                onChange={e => e.target.files?.length && handleUpload(Array.from(e.target.files))} 
              />
            </div>
          </div>

          {/* Tabs */}
          <div style={{ display: 'flex', borderBottom: '1px solid var(--border)', background: 'var(--bg)', paddingLeft: 12 }}>
            {[
              { id: 'files', label: '📁 My Files' },
              { id: 'sheet', label: '⊞ Sheet View', disabled: !currentFile },
              { id: 'analyze', label: '📊 Analysis', disabled: !currentFile },
              { id: 'charts', label: '📈 Charts', disabled: !currentFile },
              { id: 'create', label: '✨ Create New' },
            ].map(tab => (
              <button key={tab.id} onClick={() => !tab.disabled && setActiveTab(tab.id as MainTab)} disabled={tab.disabled}
                style={{ padding: '8px 16px', background: 'none', border: 'none', borderBottom: activeTab === tab.id ? '2px solid var(--accent)' : '2px solid transparent', cursor: tab.disabled ? 'not-allowed' : 'pointer', fontSize: '0.8rem', color: activeTab === tab.id ? 'var(--accent)' : tab.disabled ? 'var(--text-muted)' : 'var(--text-secondary)', opacity: tab.disabled ? 0.5 : 1 }}>
                {tab.label}
              </button>
            ))}
          </div>

          {/* FORMULA BAR */}
          {activeTab === 'sheet' && currentFile && selectedCell && (
            <div style={{ display: 'flex', alignItems: 'center', gap: 8, padding: '6px 12px', borderBottom: '1px solid var(--border)', background: 'var(--surface)', height: 36 }}>
              <span style={{ minWidth: 60, padding: '4px 8px', background: 'var(--surface-2)', border: '1px solid var(--border)', borderRadius: 4, fontSize: '0.7rem', fontFamily: 'monospace' }}>
                {getCellAddr(selectedCell.r, selectedCell.c)}
              </span>
              <span style={{ fontSize: '0.8rem', color: 'var(--text-muted)' }}>ƒx</span>
              <input value={formulaBarVal} onChange={e => setFormulaBarVal(e.target.value)} onKeyDown={e => e.key === 'Enter' && selectedCell && commitEdit(selectedCell.r, selectedCell.c, formulaBarVal)} style={{ flex: 1, border: 'none', background: 'transparent', outline: 'none', fontSize: '0.8rem', fontFamily: 'monospace' }} />
              <button className="btn btn-primary btn-sm" onClick={saveCurrentSheet} disabled={savingSheet} style={{ flexShrink: 0 }}>
                {savingSheet ? 'Saving...' : 'Save Workbook'}
              </button>
            </div>
          )}

          {/* TAB: FILES */}
          {activeTab === 'files' && (
            <div style={{ flex: 1, overflowY: 'auto', padding: 16 }}>
              <div style={{ display: 'flex', gap: 8, marginBottom: 16 }}>
                <input 
                  placeholder="Search files..." 
                  value={fileSearch} 
                  onChange={e => setFileSearch(e.target.value)}
                  onKeyDown={e => e.key === 'Enter' && loadFiles()}
                  style={{ flex: 1, padding: '8px 12px', background: 'var(--surface)', border: '1px solid var(--border)', borderRadius: 6, color: 'var(--text)' }}
                />
                <button className="btn btn-outline btn-sm" onClick={loadFiles}>↻ Refresh</button>
              </div>
              <div style={{ marginTop: -6, marginBottom: 14, fontSize: '0.72rem', color: 'var(--text-muted)' }}>
                Excel files only: `.xlsx`, `.xls`, and `.csv`.
              </div>

              {filesLoading ? (
                <div style={{ textAlign: 'center', padding: 40 }}>Loading...</div>
              ) : files.length === 0 ? (
                <div style={{ textAlign: 'center', padding: '60px 20px', color: 'var(--text-muted)' }}>
                  <div style={{ fontSize: 48, marginBottom: 12 }}>📂</div>
                  <p>No files uploaded yet.</p>
                  <button className="btn btn-primary" style={{ marginTop: 12 }} onClick={() => fileRef.current?.click()}>Upload Excel Files</button>
                </div>
              ) : (
                files.map(file => (
                  <div key={file.id} style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '12px 16px', background: 'var(--surface)', marginBottom: 8, borderRadius: 8, border: '1px solid var(--border)', cursor: 'pointer' }} onClick={() => openFile(file)}>
                    <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
                      <FileBadge file={file} />
                      <div>
                        <div style={{ fontWeight: 600, color: 'var(--text)' }}>{file.name}</div>
                        <div style={{ fontSize: '0.7rem', color: 'var(--text-muted)' }}>{(file.size / 1024).toFixed(1)} KB • {file.department || 'general'}</div>
                      </div>
                    </div>
                    <div style={{ display: 'flex', gap: 8 }}>
                      <button className="btn btn-primary btn-sm" onClick={(e) => { e.stopPropagation(); openFile(file) }}>Open</button>
                      <button className="btn btn-ghost btn-sm" style={{ color: 'var(--red)' }} onClick={(e) => { e.stopPropagation(); handleDeleteFile(file.id) }}>Delete</button>
                    </div>
                  </div>
                ))
              )}
            </div>
          )}

          {/* TAB: SHEET VIEW */}
          {activeTab === 'sheet' && (
            <div style={{ flex: 1, display: 'flex', flexDirection: 'column', overflow: 'hidden' }}>
              {sheetLoading ? (
                <div style={{ flex: 1, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>Loading sheet...</div>
              ) : !grid.length ? (
                <div style={{ flex: 1, display: 'flex', alignItems: 'center', justifyContent: 'center', flexDirection: 'column', gap: 12 }}>
                  <div style={{ fontSize: 48, opacity: 0.3 }}>⊞</div>
                  <p>Open a file from My Files tab</p>
                  <button className="btn btn-primary" onClick={() => setActiveTab('files')}>Go to My Files</button>
                </div>
              ) : (
                <div tabIndex={0} onKeyDown={handleGridKeyDown} style={{ flex: 1, overflow: 'auto', outline: 'none' }}>
                  <table style={{ borderCollapse: 'collapse', fontSize: '0.75rem', minWidth: '100%' }}>
                    <thead style={{ position: 'sticky', top: 0, zIndex: 10 }}>
                      <tr>
                        <th style={{ width: 45, height: 28, border: '1px solid var(--border)', background: 'var(--surface-2)', position: 'sticky', left: 0 }} />
                        {Array(gridCols).fill(0).map((_, ci) => (
                          <th key={ci} style={{ width: colWidths[ci] || 100, height: 28, border: '1px solid var(--border)', background: 'var(--surface-2)', textAlign: 'center', fontWeight: 600, color: 'var(--text-muted)' }}>
                            {ci < 26 ? COL_LETTERS[ci] : String.fromCharCode(65 + Math.floor(ci / 26) - 1) + COL_LETTERS[ci % 26]}
                          </th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {grid.map((row, ri) => (
                        <tr key={ri}>
                          <td style={{ width: 45, height: 26, border: '1px solid var(--border)', background: 'var(--surface-2)', textAlign: 'center', fontWeight: 600, color: 'var(--text-muted)', position: 'sticky', left: 0, zIndex: 5 }}>{ri + 1}</td>
                          {row.map((cell, ci) => {
                            const isSelected = selectedCell?.r === ri && selectedCell?.c === ci
                            return (
                              <td key={ci} style={{ width: colWidths[ci] || 100, height: 26, border: `1px solid ${isSelected ? 'var(--accent)' : 'var(--border)'}`, padding: 0, background: isSelected ? 'rgba(59,130,246,0.1)' : ri === 0 ? 'var(--surface-2)' : 'var(--bg)' }} onClick={() => selectCell(ri, ci)} onDoubleClick={() => startEdit(ri, ci)}>
                                {isEditing && isSelected ? (
                                  <input ref={editInputRef} defaultValue={cell.value} onBlur={e => commitEdit(ri, ci, e.target.value)} onKeyDown={e => handleKeyDown(e, ri, ci)} style={{ width: '100%', height: '100%', border: 'none', padding: '0 6px', outline: 'none', fontSize: '0.75rem', background: 'rgba(59,130,246,0.2)' }} autoFocus />
                                ) : (
                                  <div style={{ padding: '0 6px', height: '100%', display: 'flex', alignItems: 'center', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{cell.value}</div>
                                )}
                              </td>
                            )
                          })}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}

              {/* AI Command Bar */}
              

              {/* Command Result */}
              
            </div>
          )}

          {/* TAB: ANALYSIS */}
          {activeTab === 'analyze' && curStats && (
            <div style={{ flex: 1, overflowY: 'auto', padding: 16 }}>
              <div style={{ display: 'grid', gridTemplateColumns: '1.15fr 0.85fr', gap: 14, marginBottom: 18 }}>
                <div style={{ padding: '16px 18px', borderRadius: 16, background: 'linear-gradient(135deg, rgba(37,99,235,0.16), rgba(15,23,42,0.96))', border: '1px solid rgba(96,165,250,0.18)' }}>
                  <div style={{ fontSize: '0.72rem', color: '#bfdbfe', textTransform: 'uppercase', letterSpacing: '0.08em', marginBottom: 6 }}>AI Analysis Studio</div>
                  <div style={{ fontSize: '1.02rem', color: '#f8fafc', fontWeight: 700, marginBottom: 6 }}>Live insights linked to your workbook</div>
                  <div style={{ fontSize: '0.8rem', color: '#dbeafe', lineHeight: 1.65 }}>
                    Use chat or the controls here to ask for risks, comparisons, summaries, and recommended visuals. The analysis prompt stays connected to the current sheet.
                  </div>
                </div>
                <div style={{ padding: '16px 18px', borderRadius: 16, background: 'var(--surface)', border: '1px solid var(--border)' }}>
                  <div style={{ fontSize: '0.72rem', color: 'var(--text-muted)', textTransform: 'uppercase', letterSpacing: '0.08em', marginBottom: 8 }}>AI Prompt</div>
                  <input value={cmdInput} onChange={e => setCmdInput(e.target.value)} onKeyDown={e => e.key === 'Enter' && runCommand()} placeholder='Ask for insight, risk, ranking, summary, or a recommendation...' style={{ width: '100%', padding: '10px 12px', background: 'var(--surface-2)', border: '1px solid var(--border)', borderRadius: 10, color: 'var(--text)', marginBottom: 10 }} />
                  <button className="btn btn-primary btn-sm" onClick={runCommand} disabled={cmdLoading || !cmdInput.trim()} style={{ width: '100%', justifyContent: 'center' }}>
                    {cmdLoading ? 'Analyzing...' : 'Run AI Analysis'}
                  </button>
                </div>
              </div>

              {workspaceIntent?.tab === 'analyze' && workspaceIntent.questions.length > 0 && (
                <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap', marginBottom: 18 }}>
                  {workspaceIntent.questions.map(question => (
                    <button key={question} className="btn btn-outline btn-sm" onClick={() => applyIntentQuestion(question)}>
                      {question}
                    </button>
                  ))}
                </div>
              )}

              <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4,1fr)', gap: 12, marginBottom: 20 }}>
                <div style={{ background: 'var(--surface)', padding: 12, borderRadius: 8, textAlign: 'center', border: '1px solid var(--border)' }}><div style={{ fontSize: 24, fontWeight: 700, color: '#3b82f6' }}>{curStats.total}</div><div style={{ fontSize: 12, color: 'var(--text-muted)' }}>Total Students</div></div>
                <div style={{ background: 'var(--surface)', padding: 12, borderRadius: 8, textAlign: 'center', border: '1px solid var(--border)' }}><div style={{ fontSize: 24, fontWeight: 700, color: '#22c55e' }}>{curStats.passed}</div><div style={{ fontSize: 12, color: 'var(--text-muted)' }}>Passed</div></div>
                <div style={{ background: 'var(--surface)', padding: 12, borderRadius: 8, textAlign: 'center', border: '1px solid var(--border)' }}><div style={{ fontSize: 24, fontWeight: 700, color: '#ef4444' }}>{curStats.failed}</div><div style={{ fontSize: 12, color: 'var(--text-muted)' }}>Failed</div></div>
                <div style={{ background: 'var(--surface)', padding: 12, borderRadius: 8, textAlign: 'center', border: '1px solid var(--border)' }}><div style={{ fontSize: 24, fontWeight: 700, color: '#f97316' }}>{curStats.passRate}%</div><div style={{ fontSize: 12, color: 'var(--text-muted)' }}>Pass Rate</div></div>
              </div>

              {cmdResult && (
                <div style={{ background: 'var(--surface)', border: '1px solid var(--border)', borderRadius: 14, padding: 14 }}>
                  <div style={{ fontSize: '0.78rem', color: 'var(--green)', fontWeight: 700, marginBottom: 8 }}>{cmdResult.message || 'AI analysis result'}</div>
                  {Array.isArray(cmdResult.data) && cmdResult.data.length > 0 ? (
                    <div style={{ overflowX: 'auto' }}>
                      <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.72rem' }}>
                        <thead>
                          <tr>
                            {Object.keys(cmdResult.data[0]).slice(0, 5).map(key => (
                              <th key={key} style={{ padding: '6px 8px', textAlign: 'left', color: 'var(--text-muted)', borderBottom: '1px solid var(--border)' }}>{key}</th>
                            ))}
                          </tr>
                        </thead>
                        <tbody>
                          {cmdResult.data.slice(0, 10).map((row: any, idx: number) => (
                            <tr key={idx}>
                              {Object.values(row).slice(0, 5).map((value: any, cellIdx: number) => (
                                <td key={cellIdx} style={{ padding: '6px 8px', color: 'var(--text-secondary)', borderBottom: '1px solid var(--border)' }}>{String(value)}</td>
                              ))}
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  ) : (
                    <div style={{ fontSize: '0.8rem', color: 'var(--text-secondary)' }}>Run a prompt like "compare sections" or "show top risks" to get a live AI answer here.</div>
                  )}
                </div>
              )}
            </div>
          )}

          {(aiStatus || aiPreview || aiHistory) && (
            <div style={{ padding: '12px 14px', borderBottom: '1px solid var(--border)', background: 'linear-gradient(135deg, rgba(15,23,42,0.98), rgba(15,23,42,0.9))' }}>
              {aiStatus && (
                <div style={{ display: 'flex', alignItems: 'center', gap: 10, marginBottom: aiPreview || aiHistory ? 10 : 0 }}>
                  <div style={{ width: 10, height: 10, borderRadius: '50%', background: aiStatus.tone === 'processing' ? '#f59e0b' : aiStatus.tone === 'applied' ? '#22c55e' : '#3b82f6', boxShadow: `0 0 0 6px ${aiStatus.tone === 'processing' ? 'rgba(245,158,11,0.12)' : aiStatus.tone === 'applied' ? 'rgba(34,197,94,0.12)' : 'rgba(59,130,246,0.12)'}` }} />
                  <div>
                    <div style={{ fontSize: '0.72rem', color: '#cbd5e1', textTransform: 'uppercase', letterSpacing: '0.08em' }}>AI Workspace</div>
                    <div style={{ fontSize: '0.82rem', color: '#f8fafc', fontWeight: 600 }}>{aiStatus.text}</div>
                  </div>
                </div>
              )}

              {aiPreview && (
                <div style={{ display: 'grid', gridTemplateColumns: 'minmax(0, 1fr) auto', gap: 14, alignItems: 'start', padding: 14, background: 'rgba(30,41,59,0.82)', border: '1px solid rgba(148,163,184,0.18)', borderRadius: 14 }}>
                  <div>
                    <div style={{ display: 'flex', gap: 8, alignItems: 'center', marginBottom: 6 }}>
                      <span style={{ fontSize: '0.68rem', color: '#93c5fd', textTransform: 'uppercase', letterSpacing: '0.08em' }}>
                        {aiPreview.kind === 'rows' ? 'Sheet Preview' : 'Chart Preview'}
                      </span>
                      <span style={{ fontSize: '0.65rem', color: '#94a3b8' }}>{new Date(aiPreview.createdAt).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}</span>
                    </div>
                    <div style={{ fontSize: '0.95rem', color: '#f8fafc', fontWeight: 700, marginBottom: 6 }}>{aiPreview.title}</div>
                    <div style={{ fontSize: '0.76rem', color: '#cbd5e1', lineHeight: 1.6, marginBottom: 10 }}>{aiPreview.description}</div>

                    {aiPreview.kind === 'rows' && aiPreview.rows && (
                      <div style={{ border: '1px solid rgba(148,163,184,0.16)', borderRadius: 12, overflow: 'hidden', background: 'rgba(15,23,42,0.65)' }}>
                        <div style={{ padding: '8px 10px', fontSize: '0.68rem', color: '#94a3b8', borderBottom: '1px solid rgba(148,163,184,0.14)' }}>
                          Previewing {Math.max(aiPreview.rows.length - 1, 0)} rows and {aiPreview.rows[0]?.length || 0} columns before applying
                        </div>
                        <div style={{ overflowX: 'auto' }}>
                          <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.72rem' }}>
                            <thead>
                              <tr>
                                {(aiPreview.rows[0] || []).slice(0, 6).map((cell, idx) => (
                                  <th key={idx} style={{ padding: '7px 9px', textAlign: 'left', color: '#e2e8f0', borderBottom: '1px solid rgba(148,163,184,0.14)', background: 'rgba(30,41,59,0.85)' }}>{cell}</th>
                                ))}
                              </tr>
                            </thead>
                            <tbody>
                              {aiPreview.rows.slice(1, 6).map((row, rowIdx) => (
                                <tr key={rowIdx}>
                                  {row.slice(0, 6).map((cell, colIdx) => (
                                    <td key={colIdx} style={{ padding: '7px 9px', color: '#cbd5e1', borderBottom: '1px solid rgba(148,163,184,0.12)' }}>{cell}</td>
                                  ))}
                                </tr>
                              ))}
                            </tbody>
                          </table>
                        </div>
                      </div>
                    )}

                    {aiPreview.kind === 'chart' && (
                      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(160px, 1fr))', gap: 10 }}>
                        <div style={{ padding: '10px 12px', background: 'rgba(15,23,42,0.55)', border: '1px solid rgba(148,163,184,0.14)', borderRadius: 12 }}>
                          <div style={{ fontSize: '0.65rem', color: '#94a3b8', textTransform: 'uppercase' }}>Chart Type</div>
                          <div style={{ fontSize: '0.88rem', color: '#f8fafc', fontWeight: 600, marginTop: 4 }}>{aiPreview.chartType}</div>
                        </div>
                        <div style={{ padding: '10px 12px', background: 'rgba(15,23,42,0.55)', border: '1px solid rgba(148,163,184,0.14)', borderRadius: 12 }}>
                          <div style={{ fontSize: '0.65rem', color: '#94a3b8', textTransform: 'uppercase' }}>Data Source</div>
                          <div style={{ fontSize: '0.88rem', color: '#f8fafc', fontWeight: 600, marginTop: 4 }}>{aiPreview.chartDataKey}</div>
                        </div>
                      </div>
                    )}
                  </div>

                  <div style={{ width: 220, display: 'flex', flexDirection: 'column', gap: 10 }}>
                    <button className="btn btn-primary" onClick={applyAiPreview}>Apply Preview</button>
                    <button className="btn btn-outline" onClick={discardAiPreview}>Discard</button>
                    <button className="btn btn-ghost" disabled={!aiHistory} onClick={undoAiApply} style={{ opacity: aiHistory ? 1 : 0.5 }}>
                      Undo Last AI Change
                    </button>
                    <div style={{ padding: '10px 12px', background: 'rgba(15,23,42,0.55)', border: '1px solid rgba(148,163,184,0.14)', borderRadius: 12 }}>
                      <div style={{ fontSize: '0.64rem', color: '#94a3b8', textTransform: 'uppercase', marginBottom: 6 }}>Modify Hint</div>
                      <textarea
                        value={previewNote}
                        onChange={e => setPreviewNote(e.target.value)}
                        placeholder='Example: add monthly trend and sort highest first'
                        style={{ width: '100%', minHeight: 72, resize: 'vertical', borderRadius: 8, border: '1px solid rgba(148,163,184,0.14)', background: 'rgba(2,6,23,0.65)', color: '#e2e8f0', padding: '8px 10px', fontSize: '0.72rem', boxSizing: 'border-box' }}
                      />
                      <div style={{ fontSize: '0.67rem', color: '#cbd5e1', lineHeight: 1.5, marginTop: 8 }}>
                        Use this note in chat to ask for the next revision before you apply the preview.
                      </div>
                    </div>
                  </div>
                </div>
              )}
            </div>
          )}

          {activeTab === 'charts' && currentFile && (
            <div style={{ flex: 1, overflowY: 'auto', padding: 16 }}>
              <div style={{ display: 'grid', gridTemplateColumns: '1.1fr 0.9fr', gap: 14, marginBottom: 18 }}>
                <div style={{ padding: '16px 18px', borderRadius: 16, background: 'linear-gradient(135deg, rgba(14,165,233,0.16), rgba(15,23,42,0.96))', border: '1px solid rgba(56,189,248,0.18)' }}>
                  <div style={{ fontSize: '0.72rem', color: '#bae6fd', textTransform: 'uppercase', letterSpacing: '0.08em', marginBottom: 6 }}>AI Chart Studio</div>
                  <div style={{ fontSize: '1.02rem', color: '#f8fafc', fontWeight: 700, marginBottom: 6 }}>Build visuals from chat or refine them here</div>
                  <div style={{ fontSize: '0.8rem', color: '#e0f2fe', lineHeight: 1.65 }}>
                    Say "create chart" in chat and it will open this tab with a prepared direction. You can switch chart style, change the data slice, then apply the preview.
                  </div>
                </div>
                <div style={{ padding: '16px 18px', borderRadius: 16, background: 'var(--surface)', border: '1px solid var(--border)' }}>
                  <div style={{ fontSize: '0.72rem', color: 'var(--text-muted)', textTransform: 'uppercase', letterSpacing: '0.08em', marginBottom: 8 }}>Suggested Questions</div>
                  <div style={{ display: 'flex', flexWrap: 'wrap', gap: 8 }}>
                    {(workspaceIntent?.tab === 'charts' ? workspaceIntent.questions : ['Which metric should I visualize?', 'Compare sections with a bar chart', 'Show score distribution as line', 'Make a pie chart for grades']).map(question => (
                      <button key={question} className="btn btn-outline btn-sm" onClick={() => applyIntentQuestion(question)}>
                        {question}
                      </button>
                    ))}
                  </div>
                </div>
              </div>

              <div style={{ display: 'flex', gap: 10, marginBottom: 16, flexWrap: 'wrap' }}>
                <select value={chartType} onChange={e => setChartType(e.target.value)} style={{ padding: '8px 10px', background: 'var(--surface)', border: '1px solid var(--border)', borderRadius: 8, color: 'var(--text)' }}>
                  <option value="bar">Bar Chart</option>
                  <option value="line">Line Chart</option>
                  <option value="pie">Pie Style</option>
                </select>
                <select value={chartDataKey} onChange={e => setChartDataKey(e.target.value)} style={{ padding: '8px 10px', background: 'var(--surface)', border: '1px solid var(--border)', borderRadius: 8, color: 'var(--text)' }}>
                  <option value="section">Section</option>
                  <option value="subject">Subject</option>
                  <option value="grade">Grade</option>
                  <option value="score">Score Distribution</option>
                </select>
              </div>

              {!chartRows.length ? (
                <div style={{ padding: 24, background: 'var(--surface)', border: '1px solid var(--border)', borderRadius: 12, color: 'var(--text-muted)' }}>
                  No chartable summary found for this file. Try asking AI for a chart after opening a dataset with sections, grades, or score distribution.
                </div>
              ) : (
                <div style={{ background: 'var(--surface)', border: '1px solid var(--border)', borderRadius: 14, padding: 16 }}>
                  <div style={{ fontSize: '1rem', fontWeight: 700, marginBottom: 6, color: 'var(--text)' }}>{currentFile.name} Visualization</div>
                  <div style={{ fontSize: '0.78rem', color: 'var(--text-muted)', marginBottom: 18 }}>AI-ready chart for {chartDataKey} data using {chartType} mode.</div>
                  <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
                    {chartRows.slice(0, 12).map((row: any, idx: number) => {
                      const label = getChartLabel(row)
                      const value = getChartValue(row)
                      const max = Math.max(...chartRows.map((item: any) => getChartValue(item)), 1)
                      const width = `${Math.max((value / max) * 100, 6)}%`
                      return (
                        <div key={`${label}_${idx}`} style={{ display: 'grid', gridTemplateColumns: '140px 1fr 60px', gap: 10, alignItems: 'center' }}>
                          <div style={{ fontSize: '0.74rem', color: 'var(--text-secondary)', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{label}</div>
                          <div style={{ height: 28, background: 'var(--bg)', border: '1px solid var(--border)', borderRadius: 999, overflow: 'hidden', position: 'relative' }}>
                            <div style={{ width, height: '100%', background: chartType === 'pie' ? 'linear-gradient(90deg, #22c55e, #3b82f6)' : chartType === 'line' ? 'linear-gradient(90deg, #f59e0b, #ef4444)' : 'linear-gradient(90deg, #60a5fa, #2563eb)', borderRadius: 999 }} />
                          </div>
                          <div style={{ fontSize: '0.75rem', color: 'var(--text)', textAlign: 'right', fontWeight: 600 }}>{value}</div>
                        </div>
                      )
                    })}
                  </div>
                </div>
              )}
            </div>
          )}

          {/* TAB: CREATE NEW */}
          {activeTab === 'create' && (
            <div style={{ flex: 1, overflowY: 'auto', padding: 32, maxWidth: 600, margin: '0 auto', width: '100%' }}>
              <div style={{ padding: '16px 18px', borderRadius: 16, background: 'linear-gradient(135deg, rgba(168,85,247,0.16), rgba(15,23,42,0.95))', border: '1px solid rgba(196,181,253,0.16)', marginBottom: 18 }}>
                <div style={{ fontSize: '0.72rem', color: '#ddd6fe', textTransform: 'uppercase', letterSpacing: '0.08em', marginBottom: 6 }}>AI Creation Studio</div>
                <div style={{ fontSize: '1rem', color: '#f5f3ff', fontWeight: 700, marginBottom: 6 }}>Describe what to create and let AI ask the next smart questions</div>
                <div style={{ fontSize: '0.8rem', color: '#ede9fe', lineHeight: 1.65 }}>
                  Use this space for new sheets, transformed tables, ranked reports, or structured exports. Commands from chat can land here with context already filled in.
                </div>
              </div>
              {workspaceIntent?.tab === 'create' && workspaceIntent.questions.length > 0 && (
                <div style={{ display: 'flex', flexWrap: 'wrap', gap: 8, marginBottom: 16 }}>
                  {workspaceIntent.questions.map(question => (
                    <button key={question} className="btn btn-outline btn-sm" onClick={() => applyIntentQuestion(question)}>
                      {question}
                    </button>
                  ))}
                </div>
              )}
              <div style={{ textAlign: 'center', marginBottom: 24 }}>
                <div style={{ fontSize: 48, marginBottom: 12 }}>✨</div>
                <h2>Create New Spreadsheet</h2>
              </div>
              <textarea value={createPrompt} onChange={e => setCreatePrompt(e.target.value)} placeholder="Describe what you need..." style={{ width: '100%', minHeight: 100, padding: 12, background: 'var(--surface)', border: '1px solid var(--border)', borderRadius: 8, marginBottom: 12 }} />
              <button className="btn btn-primary" onClick={handleCreate} disabled={creating || !createPrompt.trim()} style={{ width: '100%', padding: 12 }}>{creating ? 'Creating...' : '✨ Generate'}</button>
            </div>
          )}

        </main>

        <AIChatPanel 
          currentFile={currentFile}
          rawRows={rawGridValues}
          stats={curStats}
          onGridUpdate={handleAIGridUpdate}
          onNewFile={handleOpenGeneratedFile}
          onShowChart={handleAIChartRequest}
          onWorkspaceIntent={handleWorkspaceIntent}
          onApplyPreview={applyAiPreview}
          onDiscardPreview={discardAiPreview}
          onUndoLastChange={undoAiApply}
        />
      </div>
    </div>
  )
}
