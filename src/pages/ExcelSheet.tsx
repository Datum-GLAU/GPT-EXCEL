import { useState, useRef, useEffect, useMemo } from 'react'
import { useSelector } from 'react-redux'
import { RootState } from '../index'
import AIChatPanel from '../components/AIChatPanel'
import Sidebar from '../components/Sidebar'
import Header from '../components/Header'
import { api } from '../api'

type MainTab = 'files' | 'sheet' | 'analyze' | 'charts' | 'create' | 'offline'

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

interface AnalysisWorkspaceState {
  kind: 'analysis' | 'chart' | 'document' | 'file'
  title: string
  description: string
  rows?: any[][]
  data?: any[]
  chartType?: string
  chartDataKey?: string
  documentType?: string
  documentLength?: 'short' | 'medium' | 'long'
  content?: string
  createdAt: string
}

type PythonAction = 'analyze' | 'read' | 'clean' | 'chart' | 'report' | 'excel' | 'process' | 'template'

interface PythonOutput {
  label: string
  path: string
}

type WorkMode = 'online' | 'offline'
const EXCEL_VIEW_STATE_KEY = 'excel_view_state'

interface OfflineCommand {
  command: string
  effect: string
}

const COL_LETTERS = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.split('')
const OFFLINE_COMMANDS: OfflineCommand[] = [
  { command: 'preview rows', effect: 'Show the first records from the current workbook.' },
  { command: 'analyze data', effect: 'Return rows, columns, null counts, and basic statistics.' },
  { command: 'clean data', effect: 'Remove duplicates and fill missing values.' },
  { command: 'create chart', effect: 'Generate a chart image from the workbook.' },
  { command: 'generate report', effect: 'Create a text summary report.' },
  { command: 'advanced excel', effect: 'Build a formatted workbook with summary sheet.' },
  { command: 'file info', effect: 'Show file size, column names, and numeric summary.' },
  { command: 'split rows', effect: 'Split a large workbook into chunks.' },
  { command: 'segment by region', effect: 'Create one file per unique value in a column.' },
  { command: 'template', effect: 'Create a starter Excel template offline.' },
  { command: 'top numeric columns', effect: 'Inspect numeric fields quickly.' },
  { command: 'show null values', effect: 'See missing values per column.' },
  { command: 'show duplicates', effect: 'Identify duplicate row counts.' },
  { command: 'summary statistics', effect: 'Return descriptive stats for each column.' },
  { command: 'read first 10 rows', effect: 'Preview a compact sample from the workbook.' },
  { command: 'bar chart', effect: 'Build a bar chart from available columns.' },
  { command: 'line chart', effect: 'Build a line chart for trend-like data.' },
  { command: 'pie chart', effect: 'Build a pie chart for category totals.' },
  { command: 'scatter chart', effect: 'Build a scatter chart for numeric comparisons.' },
  { command: 'clean column names', effect: 'Normalize spaces and casing in headers.' },
  { command: 'fill numeric nulls', effect: 'Replace numeric blanks with medians.' },
  { command: 'fill text nulls', effect: 'Replace text blanks with fallback labels.' },
  { command: 'save cleaned excel', effect: 'Generate a cleaned workbook file.' },
  { command: 'generate summary sheet', effect: 'Create a summary worksheet offline.' },
  { command: 'merge files', effect: 'Combine compatible Excel files into one output.' },
  { command: 'segment by month', effect: 'Split data by date period when a date column exists.' },
  { command: 'segment by quarter', effect: 'Split rows by quarter from a date column.' },
  { command: 'segment by year', effect: 'Split rows into yearly workbooks.' },
  { command: 'chunk 500 rows', effect: 'Break the workbook into 500-row parts.' },
  { command: 'chunk 1000 rows', effect: 'Break the workbook into 1000-row parts.' },
  { command: 'show numeric summary', effect: 'Summarize mean, std, min, and max values.' },
  { command: 'show text columns', effect: 'List text-style columns in the workbook.' },
  { command: 'show numeric columns', effect: 'List numeric columns in the workbook.' },
  { command: 'count rows', effect: 'Return total workbook row count.' },
  { command: 'count columns', effect: 'Return total workbook column count.' },
  { command: 'show column types', effect: 'Display inferred data types per column.' },
  { command: 'find outliers', effect: 'Use numeric stats as a quick outlier guide.' },
  { command: 'check empty columns', effect: 'Surface columns with only missing values.' },
  { command: 'report duplicate rows', effect: 'Return duplicate row totals in analysis.' },
  { command: 'compare max and min', effect: 'Show value spread for numeric fields.' },
  { command: 'sheet template sales', effect: 'Create a basic sales sheet template.' },
  { command: 'sheet template attendance', effect: 'Create an attendance-style template.' },
  { command: 'sheet template budget', effect: 'Create a budget-friendly starter sheet.' },
  { command: 'sheet template inventory', effect: 'Create a basic inventory starter file.' },
  { command: 'offline pipeline', effect: 'Run info, analysis, clean, chart, report, and advanced Excel together.' },
  { command: 'preview 25 rows', effect: 'Read a larger preview sample.' },
  { command: 'show first headers', effect: 'Display header names quickly.' },
  { command: 'chart amount by category', effect: 'Try a category summary chart.' },
  { command: 'chart revenue trend', effect: 'Try a line chart for trend columns.' },
  { command: 'analyze attendance', effect: 'Analyze attendance-like numeric columns.' },
  { command: 'analyze marks', effect: 'Analyze marks or score columns.' },
  { command: 'analyze salary data', effect: 'Profile salary-like numeric sheets.' },
  { command: 'analyze expenses', effect: 'Review expense sheets offline.' },
  { command: 'analyze inventory', effect: 'Review stock or inventory counts.' },
  { command: 'null percentage', effect: 'Return missing-value percentages.' },
  { command: 'median values', effect: 'Review median values for numeric fields.' },
  { command: 'variance values', effect: 'Inspect variance for numeric fields.' },
  { command: 'quartile summary', effect: 'Show Q1 and Q3 statistics.' },
  { command: 'range summary', effect: 'Show max-min range by numeric column.' },
  { command: 'status report', effect: 'Create a plain-text report you can open locally.' },
  { command: 'offline chart export', effect: 'Export chart PNG for the active workbook.' },
  { command: 'offline excel export', effect: 'Export the advanced Excel output.' },
  { command: 'preview workbook', effect: 'Read workbook rows without editing them.' },
  { command: 'normalize headers', effect: 'Trim spaces and convert headers to snake_case.' },
  { command: 'replace blanks', effect: 'Fill empty cells during the clean step.' },
  { command: 'remove repeats', effect: 'Drop duplicate rows offline.' },
  { command: 'check file size', effect: 'Return workbook size in KB.' },
  { command: 'check date column', effect: 'Use date parsing in segmentation tasks.' },
  { command: 'split by column', effect: 'Split data into files by a chosen column.' },
  { command: 'split by rows', effect: 'Split data into row-count chunks.' },
  { command: 'split by date', effect: 'Split data by month, week, quarter, or year.' },
  { command: 'monthly segments', effect: 'Create monthly segmented files.' },
  { command: 'weekly segments', effect: 'Create weekly segmented files.' },
  { command: 'daily segments', effect: 'Create daily segmented files.' },
  { command: 'yearly segments', effect: 'Create yearly segmented files.' },
  { command: 'chart top categories', effect: 'Show the most frequent categories in a chart.' },
  { command: 'show output paths', effect: 'Return generated file locations.' },
  { command: 'download cleaned file', effect: 'Open the cleaned Excel output path.' },
  { command: 'download chart', effect: 'Open the generated chart output path.' },
  { command: 'download report', effect: 'Open the generated report output path.' },
  { command: 'offline workbook summary', effect: 'Get a concise workbook summary without AI.' },
  { command: 'offline data profile', effect: 'Profile the current workbook locally.' },
  { command: 'local excel stats', effect: 'Compute workbook stats with Python only.' },
  { command: 'local chart build', effect: 'Build a chart without cloud AI.' },
  { command: 'local cleanup', effect: 'Clean workbook values without internet access.' },
  { command: 'local export', effect: 'Export processed workbook artifacts offline.' },
  { command: 'sheet health check', effect: 'Inspect row count, types, nulls, and duplicates.' },
  { command: 'dataset overview', effect: 'See a dataset overview block.' },
  { command: 'quick audit', effect: 'Get a quick workbook audit offline.' },
  { command: 'operations summary', effect: 'Summarize available offline operations.' },
  { command: 'offline help', effect: 'Use this command list as a quick reference.' },
  { command: 'analyze column types', effect: 'Review inferred types by header.' },
  { command: 'check distribution', effect: 'Use basic stats to inspect distributions.' },
  { command: 'check score spread', effect: 'Review range and quartiles for score columns.' },
  { command: 'check sales spread', effect: 'Review spread for sales fields.' },
  { command: 'count unique values', effect: 'Useful before splitting by a category column.' },
  { command: 'prepare cleaned export', effect: 'Generate a cleaned workbook output.' },
  { command: 'prepare summary export', effect: 'Generate a workbook with summary sheet.' },
  { command: 'prepare chart export', effect: 'Generate a local PNG chart.' },
  { command: 'prepare report export', effect: 'Generate a local TXT report.' },
  { command: 'run offline mode', effect: 'Use the local Python engine when network AI is unavailable.' },
  { command: 'return to online mode', effect: 'Switch back to AI-powered online workflow.' },
]

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
  const [activeTab, setActiveTab] = useState<MainTab>(() => {
    try {
      const saved = JSON.parse(localStorage.getItem(EXCEL_VIEW_STATE_KEY) || '{}')
      return (saved.activeTab as MainTab) || 'files'
    } catch {
      return 'files'
    }
  })
  const [networkAvailable, setNetworkAvailable] = useState<boolean>(() => typeof navigator === 'undefined' ? true : navigator.onLine)
  const [backendOnline, setBackendOnline] = useState(false)
  const [pythonOnline, setPythonOnline] = useState(false)
  const [workMode, setWorkMode] = useState<WorkMode>(() => typeof navigator !== 'undefined' && !navigator.onLine ? 'offline' : 'online')
  const [modeSwitching, setModeSwitching] = useState(false)
  const [aiPanelWidth, setAiPanelWidth] = useState(380)
  const [resizingPanel, setResizingPanel] = useState(false)

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
  const [analysisWorkspace, setAnalysisWorkspace] = useState<AnalysisWorkspaceState | null>(null)

  // Create
  const [createPrompt, setCreatePrompt] = useState('')
  const [creating, setCreating] = useState(false)
  const [pythonBusy, setPythonBusy] = useState<PythonAction | null>(null)
  const [pythonResult, setPythonResult] = useState<any>(null)
  const [pythonOutputs, setPythonOutputs] = useState<PythonOutput[]>([])
  const [pythonPrompt, setPythonPrompt] = useState('')

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

  const effectiveOffline = false
  const offlineReason = !networkAvailable
    ? 'No network detected. Switching to local offline tools.'
    : !backendOnline
      ? 'AI backend is unavailable. Local offline tools are ready instead.'
      : 'Offline mode is active. Local commands and Python tools are available.'
  const visibleTabs: Array<{ id: MainTab; label: string; disabled?: boolean }> = [
    { id: 'files', label: 'My Files' },
    { id: 'sheet', label: 'Sheet View', disabled: !currentFile },
    { id: 'analyze', label: 'Analysis', disabled: !currentFile },
    { id: 'charts', label: 'Charts', disabled: !currentFile },
  ]

  useEffect(() => {
    const checkBackend = async () => {
      try {
        await api.health()
        setBackendOnline(true)
      } catch {
        setBackendOnline(false)
      }
      try {
        await api.pythonHealth()
        setPythonOnline(true)
      } catch {
        setPythonOnline(false)
      }
    }
    checkBackend()
    loadFiles()
    return () => {
      if (toastTimeoutRef.current) clearTimeout(toastTimeoutRef.current)
    }
  }, [])

  useEffect(() => {
    const handleOnline = () => setNetworkAvailable(true)
    const handleOffline = () => setNetworkAvailable(false)
    window.addEventListener('online', handleOnline)
    window.addEventListener('offline', handleOffline)
    return () => {
      window.removeEventListener('online', handleOnline)
      window.removeEventListener('offline', handleOffline)
    }
  }, [])

  useEffect(() => {
    if (!effectiveOffline && activeTab === 'offline') setActiveTab('files')
  }, [effectiveOffline, activeTab])

  useEffect(() => {
    if (!resizingPanel) return
    const handleMove = (event: MouseEvent) => {
      const nextWidth = window.innerWidth - event.clientX
      setAiPanelWidth(Math.min(620, Math.max(300, nextWidth)))
    }
    const handleUp = () => setResizingPanel(false)
    window.addEventListener('mousemove', handleMove)
    window.addEventListener('mouseup', handleUp)
    return () => {
      window.removeEventListener('mousemove', handleMove)
      window.removeEventListener('mouseup', handleUp)
    }
  }, [resizingPanel])

  useEffect(() => {
    if (!networkAvailable) {
      setWorkMode('offline')
      return
    }
    if (workMode === 'offline' && backendOnline) return
    if (backendOnline) setWorkMode('online')
  }, [networkAvailable, backendOnline, workMode])

  useEffect(() => {
    const pendingId = localStorage.getItem('excel_autoload_file_id')
    const savedState = (() => {
      try { return JSON.parse(localStorage.getItem(EXCEL_VIEW_STATE_KEY) || '{}') } catch { return {} }
    })()
    const restoreId = pendingId || savedState.currentFileId
    if (!restoreId || !files.length || currentFile?.id === restoreId) return
    const target = files.find(file => file.id === restoreId)
    if (!target) return
    if (pendingId) localStorage.removeItem('excel_autoload_file_id')
    openFile(target)
  }, [files, currentFile])

  useEffect(() => {
    try {
      localStorage.setItem(EXCEL_VIEW_STATE_KEY, JSON.stringify({
        activeTab,
        currentFileId: currentFile?.id || null
      }))
    } catch {}
  }, [activeTab, currentFile?.id])

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
        try {
          const saved = JSON.parse(localStorage.getItem(EXCEL_VIEW_STATE_KEY) || '{}')
          localStorage.setItem(EXCEL_VIEW_STATE_KEY, JSON.stringify({ ...saved, currentFileId: null }))
        } catch {}
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
    setAnalysisWorkspace({
      kind: 'file',
      title: file.name,
      description: 'Created from chat. Review the preview here, then switch to Sheet View if you want to edit cells manually.',
      rows: rows || [],
      data: [],
      createdAt: new Date().toISOString()
    })
    setActiveTab('analyze')
    setAiStatus({ tone: 'ready', text: `Opened ${file.name} from AI chat. You can edit cells, save the workbook, or ask for more changes.` })
    showToast(`Opened ${file.name}`)
    try {
      await openFile(file, { preserveActiveTab: true })
    } catch {}
  }

  const openFile = async (file: any, options?: { preserveActiveTab?: boolean }) => {
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
      if (!options?.preserveActiveTab) setActiveTab('sheet')
      try {
        const saved = JSON.parse(localStorage.getItem(EXCEL_VIEW_STATE_KEY) || '{}')
        localStorage.setItem(EXCEL_VIEW_STATE_KEY, JSON.stringify({
          ...saved,
          currentFileId: file.id,
          activeTab: options?.preserveActiveTab ? saved.activeTab || activeTab : 'sheet'
        }))
      } catch {}
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
    setAnalysisWorkspace({
      kind: 'chart',
      title: title || 'AI chart preview ready',
      description: `Prepared a ${nextType} chart using ${nextKey} insights from the current sheet.`,
      chartType: nextType,
      chartDataKey: nextKey,
      data: [],
      createdAt: new Date().toISOString()
    })
    setAiStatus({ tone: 'ready', text: 'AI chart preview is ready in Excel. Apply it to keep the chart suggestion, or discard it.' })
    setActiveTab('analyze')
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

  const handleChatAnalysisResult = (result: any) => {
    setAnalysisWorkspace({
      kind: result.kind || 'analysis',
      title: result.title || 'AI analysis result',
      description: result.description || 'Chat generated a result for the current workbook.',
      rows: result.rows || [],
      data: result.data || [],
      chartType: result.chartType,
      chartDataKey: result.chartDataKey,
      createdAt: result.createdAt || new Date().toISOString()
    })
    setActiveTab('analyze')
  }

  const handleChatDocumentResult = (doc: { id: string; title: string; type: string; content: string; createdAt: string; length: 'short' | 'medium' | 'long' }) => {
    setAnalysisWorkspace({
      kind: 'document',
      title: doc.title,
      description: `Generated a ${doc.length} ${doc.type}. You can edit the draft here or continue refining it in chat.`,
      content: doc.content,
      documentType: doc.type,
      documentLength: doc.length,
      data: [],
      createdAt: doc.createdAt
    })
    setActiveTab('analyze')
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

  const applyGridMutation = (nextGrid: CellState[][], nextCols = Math.max(...nextGrid.map(row => row.length), 0)) => {
    const normalizedGrid = nextGrid.map(row => {
      const clone = row.map(cell => ({ ...cell }))
      while (clone.length < nextCols) clone.push({ value: '', editing: false, bold: false, align: 'left' })
      return clone
    })
    setGrid(normalizedGrid)
    setGridCols(nextCols)
    setColWidths(prev => {
      const next = [...prev]
      while (next.length < nextCols) next.push(110)
      return next.slice(0, nextCols)
    })
  }

  const addRow = () => {
    const nextCols = Math.max(gridCols, 1)
    const newRow = Array.from({ length: nextCols }, () => ({ value: '', editing: false, bold: false, align: 'left' as const }))
    const nextGrid = grid.length ? [...grid.map(row => row.map(cell => ({ ...cell }))), newRow] : [newRow]
    applyGridMutation(nextGrid, nextCols)
    setSelectedCell({ r: nextGrid.length - 1, c: 0 })
    setFormulaBarVal('')
  }

  const addColumn = () => {
    const nextCols = Math.max(gridCols + 1, 1)
    const nextGrid = (grid.length ? grid : [[{ value: '', editing: false, bold: false, align: 'left' as const }]]).map(row => {
      const clone = row.map(cell => ({ ...cell }))
      clone.push({ value: '', editing: false, bold: false, align: 'left' })
      return clone
    })
    applyGridMutation(nextGrid, nextCols)
    setSelectedCell(prev => ({ r: prev?.r ?? 0, c: nextCols - 1 }))
    setFormulaBarVal('')
  }

  const removeSelectedRow = () => {
    if (!selectedCell || grid.length <= 1) return
    const nextGrid = grid.map(row => row.map(cell => ({ ...cell })))
    nextGrid.splice(selectedCell.r, 1)
    applyGridMutation(nextGrid, gridCols)
    setSelectedCell({ r: Math.max(0, selectedCell.r - 1), c: Math.min(selectedCell.c, Math.max(gridCols - 1, 0)) })
    setFormulaBarVal('')
  }

  const removeSelectedColumn = () => {
    if (!selectedCell || gridCols <= 1) return
    const nextGrid = grid.map(row => row.filter((_, idx) => idx !== selectedCell.c).map(cell => ({ ...cell })))
    applyGridMutation(nextGrid, gridCols - 1)
    setSelectedCell({ r: selectedCell.r, c: Math.max(0, selectedCell.c - 1) })
    setFormulaBarVal('')
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
      setAnalysisWorkspace({
        kind: 'analysis',
        title: 'AI analysis result',
        description: res?.message || 'Command result is ready.',
        rows: Array.isArray(res?.data) && res.data.length
          ? [Object.keys(res.data[0]), ...res.data.slice(0, 12).map((row: any) => Object.values(row))]
          : [],
        data: res?.data || [],
        createdAt: new Date().toISOString()
      })
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

  const downloadAnalysisDocument = () => {
    if (analysisWorkspace?.kind !== 'document' || !analysisWorkspace.content) return
    const blob = new Blob([analysisWorkspace.content], { type: 'text/plain' })
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url
    a.download = `${analysisWorkspace.title || 'document'}.txt`
    a.click()
    URL.revokeObjectURL(url)
  }

  const getServerFileUrl = (file: any) => {
    const rawUrl = String(file?.url || '')
    if (!rawUrl) return ''
    return rawUrl.startsWith('http') ? rawUrl : `http://localhost:3001${rawUrl}`
  }

  const getCurrentWorkbookFile = async () => {
    if (!currentFile) throw new Error('Open a workbook in My Files first.')
    const sourceUrl = getServerFileUrl(currentFile)
    if (!sourceUrl) throw new Error('This workbook does not have a downloadable server URL yet.')
    const response = await fetch(sourceUrl)
    if (!response.ok) throw new Error('Could not load the current workbook from the app server.')
    const blob = await response.blob()
    const fileName = currentFile.name || 'workbook.xlsx'
    return new File([blob], fileName, { type: blob.type || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
  }

  const collectPythonOutputs = (payload: any): PythonOutput[] => {
    const outputs: PythonOutput[] = []
    const pairs = [
      ['Chart', payload?.chart_path],
      ['Cleaned Excel', payload?.output_path && String(payload.output_path).endsWith('.xlsx') ? payload.output_path : ''],
      ['Report', payload?.output_path && String(payload.output_path).endsWith('.txt') ? payload.output_path : ''],
      ['Template', payload?.output_path && String(payload.output_path).endsWith('.xlsx') ? payload.output_path : ''],
    ] as const

    pairs.forEach(([label, path]) => {
      if (path) outputs.push({ label, path: String(path) })
    })

    return outputs
  }

  const switchWorkMode = async (nextMode: WorkMode) => {
    if (nextMode === workMode) return
    setModeSwitching(true)
    showToast(nextMode === 'offline' ? 'Going to offline mode...' : 'Returning to online AI mode...')
    await new Promise(resolve => setTimeout(resolve, 900))
    if (nextMode === 'online' && (!networkAvailable || !backendOnline)) {
      showToast('Online AI mode is not available right now.', 'err')
      setModeSwitching(false)
      return
    }
    setWorkMode(nextMode)
    setModeSwitching(false)
  }

  const runPythonAction = async (action: PythonAction, promptOverride?: string) => {
    if (!pythonOnline) {
      showToast('Python engine is offline. Run: cd python_engine && uvicorn main:app --reload --port 8001', 'err')
      return
    }

    if (action !== 'template' && !currentFile) {
      showToast('Open a workbook first, then run Python actions on it.', 'err')
      return
    }

    setPythonBusy(action)
    setPythonResult(null)
    setPythonOutputs([])

    try {
      let result: any
      if (action === 'template') {
        result = await api.pythonTemplate(10, true)
      } else {
        const workbook = await getCurrentWorkbookFile()
        if (action === 'analyze') result = await api.pythonAnalyze(workbook)
        else if (action === 'read') result = await api.pythonRead(workbook)
        else if (action === 'clean') result = await api.pythonClean(workbook)
        else if (action === 'chart') result = await api.pythonChart(workbook, chartType === 'pie' ? 'pie' : chartType === 'line' ? 'line' : 'auto')
        else if (action === 'report') result = await api.pythonReport(workbook)
        else if (action === 'excel') result = await api.pythonAdvancedExcel(workbook)
        else result = await api.pythonProcess(workbook, promptOverride || pythonPrompt.trim() || createPrompt.trim() || 'analyze data')
      }

      setPythonResult(result)
      setPythonOutputs(collectPythonOutputs(result))
      showToast(`Python ${action} completed`)
    } catch (e: any) {
      showToast(e.message || `Python ${action} failed`, 'err')
    }

    setPythonBusy(null)
  }

  const handleOfflineCommand = async (command: string) => {
    const nextCommand = command.trim()
    if (!nextCommand) return
    setCmdResult(null)
    setCmdInput(nextCommand)
    setPythonPrompt(nextCommand)
    setCreatePrompt(nextCommand)
    await runPythonAction('process', nextCommand)
  }

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
          <div style={{ display: 'flex', alignItems: 'center', gap: 6, padding: '6px 12px', minHeight: 44, borderBottom: '1px solid var(--border)', background: 'var(--bg)', fontSize: '0.7rem', flexWrap: 'wrap' }}>
            <div style={{ width: 8, height: 8, borderRadius: '50%', background: networkAvailable ? '#22c55e' : '#f59e0b' }} />
            <span style={{ color: 'var(--text-muted)' }}>{networkAvailable ? 'Network available' : 'No network'}</span>
            <div style={{ width: 8, height: 8, borderRadius: '50%', background: backendOnline ? '#22c55e' : '#ef4444' }} />
            <span style={{ color: 'var(--text-muted)' }}>{backendOnline ? 'Backend connected' : 'Backend offline'}</span>
            <div style={{ width: 8, height: 8, borderRadius: '50%', background: pythonOnline ? '#22c55e' : '#f59e0b', marginLeft: 10 }} />
            <span style={{ color: 'var(--text-muted)' }}>{pythonOnline ? 'Local tools ready' : 'Local tools unavailable'}</span>
            <div style={{ marginLeft: 10, padding: '4px 10px', borderRadius: 999, background: 'rgba(59,130,246,0.12)', color: '#2563eb', fontWeight: 700 }}>
              Online AI workspace
            </div>
            <div style={{ marginLeft: 'auto', display: 'flex', gap: 8, alignItems: 'center' }}>
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
          <div style={{ display: 'flex', borderBottom: '1px solid var(--border)', background: 'var(--bg)', paddingLeft: 12, flexWrap: 'wrap' }}>
            {visibleTabs.map(tab => (
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
              <button className="btn btn-outline btn-sm" onClick={addRow} style={{ flexShrink: 0 }}>+ Row</button>
              <button className="btn btn-outline btn-sm" onClick={addColumn} style={{ flexShrink: 0 }}>+ Column</button>
              <button className="btn btn-outline btn-sm" onClick={removeSelectedRow} disabled={!selectedCell || grid.length <= 1} style={{ flexShrink: 0 }}>Delete Row</button>
              <button className="btn btn-outline btn-sm" onClick={removeSelectedColumn} disabled={!selectedCell || gridCols <= 1} style={{ flexShrink: 0 }}>Delete Column</button>
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
          {activeTab === 'analyze' && currentFile && (
            <div style={{ flex: 1, overflowY: 'auto', padding: 16 }}>
              <div style={{ display: 'grid', gridTemplateColumns: '1.15fr 0.85fr', gap: 14, marginBottom: 18 }}>
                <div style={{ padding: '16px 18px', borderRadius: 16, background: 'linear-gradient(135deg, rgba(37,99,235,0.16), rgba(15,23,42,0.96))', border: '1px solid rgba(96,165,250,0.18)' }}>
                  <div style={{ fontSize: '0.72rem', color: '#bfdbfe', textTransform: 'uppercase', letterSpacing: '0.08em', marginBottom: 6 }}>Analysis Hub</div>
                  <div style={{ fontSize: '1.02rem', color: '#f8fafc', fontWeight: 700, marginBottom: 6 }}>Chat results appear here</div>
                  <div style={{ fontSize: '0.8rem', color: '#dbeafe', lineHeight: 1.65 }}>
                    Create files, ask for analysis, generate charts, or create documents from chat. This area stays empty until a command produces something.
                  </div>
                </div>
                <div style={{ padding: '16px 18px', borderRadius: 16, background: 'var(--surface)', border: '1px solid var(--border)' }}>
                  <div style={{ fontSize: '0.72rem', color: 'var(--text-muted)', textTransform: 'uppercase', letterSpacing: '0.08em', marginBottom: 8 }}>Quick Analysis</div>
                  <input value={cmdInput} onChange={e => setCmdInput(e.target.value)} onKeyDown={e => e.key === 'Enter' && runCommand()} placeholder='Try "compare sections" or "show top risks"' style={{ width: '100%', padding: '10px 12px', background: 'var(--surface-2)', border: '1px solid var(--border)', borderRadius: 10, color: 'var(--text)', marginBottom: 10 }} />
                  <button className="btn btn-primary btn-sm" onClick={runCommand} disabled={cmdLoading || !cmdInput.trim()} style={{ width: '100%', justifyContent: 'center' }}>
                    {cmdLoading ? 'Analyzing...' : 'Run Analysis'}
                  </button>
                </div>
              </div>

              {!analysisWorkspace ? (
                <div style={{ background: 'var(--surface)', border: '1px solid var(--border)', borderRadius: 16, padding: '48px 24px', textAlign: 'center' }}>
                  <div style={{ fontSize: 38, marginBottom: 12 }}>AI</div>
                  <div style={{ fontSize: '1rem', color: 'var(--text)', fontWeight: 700, marginBottom: 8 }}>Write commands in chat to see results here</div>
                  <div style={{ fontSize: '0.8rem', color: 'var(--text-secondary)', maxWidth: 560, margin: '0 auto', lineHeight: 1.7 }}>
                    Example: create a new sheet, analyze this file, make a bar chart, or create a short report.
                  </div>
                </div>
              ) : (
                <div style={{ display: 'grid', gap: 16 }}>
                  <div style={{ background: 'var(--surface)', border: '1px solid var(--border)', borderRadius: 16, padding: 16 }}>
                    <div style={{ display: 'flex', justifyContent: 'space-between', gap: 12, alignItems: 'flex-start', marginBottom: 10 }}>
                      <div>
                        <div style={{ fontSize: '0.72rem', color: 'var(--text-muted)', textTransform: 'uppercase', letterSpacing: '0.08em' }}>{analysisWorkspace.kind}</div>
                        <div style={{ fontSize: '1.05rem', color: 'var(--text)', fontWeight: 700, marginTop: 4 }}>{analysisWorkspace.title}</div>
                      </div>
                      <div style={{ fontSize: '0.72rem', color: 'var(--text-muted)' }}>{new Date(analysisWorkspace.createdAt).toLocaleString()}</div>
                    </div>
                    <div style={{ fontSize: '0.82rem', color: 'var(--text-secondary)', lineHeight: 1.65 }}>{analysisWorkspace.description}</div>
                  </div>

                  {analysisWorkspace.kind === 'chart' && (
                    <div style={{ background: 'var(--surface)', border: '1px solid var(--border)', borderRadius: 16, padding: 16 }}>
                      <div style={{ fontSize: '0.82rem', color: 'var(--text-muted)', marginBottom: 14 }}>
                        Chart type: <strong>{analysisWorkspace.chartType}</strong> | Data source: <strong>{analysisWorkspace.chartDataKey}</strong>
                      </div>
                      {!chartRows.length ? (
                        <div style={{ color: 'var(--text-secondary)', fontSize: '0.8rem' }}>No chartable summary is available for this workbook yet.</div>
                      ) : (
                        <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
                          {chartRows.slice(0, 12).map((row: any, idx: number) => {
                            const label = getChartLabel(row)
                            const value = getChartValue(row)
                            const max = Math.max(...chartRows.map((item: any) => getChartValue(item)), 1)
                            const width = `${Math.max((value / max) * 100, 6)}%`
                            return (
                              <div key={`${label}_${idx}`} style={{ display: 'grid', gridTemplateColumns: '140px 1fr 60px', gap: 10, alignItems: 'center' }}>
                                <div style={{ fontSize: '0.74rem', color: 'var(--text-secondary)', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{label}</div>
                                <div style={{ height: 28, background: 'var(--bg)', border: '1px solid var(--border)', borderRadius: 999, overflow: 'hidden' }}>
                                  <div style={{ width, height: '100%', background: analysisWorkspace.chartType === 'pie' ? 'linear-gradient(90deg, #22c55e, #3b82f6)' : analysisWorkspace.chartType === 'line' ? 'linear-gradient(90deg, #f59e0b, #ef4444)' : 'linear-gradient(90deg, #60a5fa, #2563eb)', borderRadius: 999 }} />
                                </div>
                                <div style={{ fontSize: '0.75rem', color: 'var(--text)', textAlign: 'right', fontWeight: 600 }}>{value}</div>
                              </div>
                            )
                          })}
                        </div>
                      )}
                    </div>
                  )}

                  {analysisWorkspace.kind === 'document' && (
                    <div style={{ background: 'var(--surface)', border: '1px solid var(--border)', borderRadius: 16, padding: 16 }}>
                      <div style={{ display: 'flex', gap: 8, alignItems: 'center', marginBottom: 12, flexWrap: 'wrap' }}>
                        <div style={{ fontSize: '0.74rem', color: 'var(--text-muted)' }}>Type: <strong>{analysisWorkspace.documentType}</strong></div>
                        <div style={{ fontSize: '0.74rem', color: 'var(--text-muted)' }}>Length: <strong>{analysisWorkspace.documentLength}</strong></div>
                        <button className="btn btn-outline btn-sm" onClick={downloadAnalysisDocument}>Download .txt</button>
                      </div>
                      <textarea
                        value={analysisWorkspace.content || ''}
                        onChange={e => setAnalysisWorkspace(prev => prev && prev.kind === 'document' ? { ...prev, content: e.target.value } : prev)}
                        style={{ width: '100%', minHeight: 340, background: 'var(--bg)', border: '1px solid var(--border)', borderRadius: 12, color: 'var(--text)', padding: 14, resize: 'vertical' }}
                      />
                      <div style={{ fontSize: '0.75rem', color: 'var(--text-muted)', marginTop: 10 }}>
                        You can edit this draft here, or ask chat to rewrite it as short, medium, long, formal, simpler, or more detailed.
                      </div>
                    </div>
                  )}

                  {(analysisWorkspace.kind === 'analysis' || analysisWorkspace.kind === 'file') && (
                    <div style={{ background: 'var(--surface)', border: '1px solid var(--border)', borderRadius: 16, padding: 16 }}>
                      {analysisWorkspace.rows && analysisWorkspace.rows.length > 0 ? (
                        <div style={{ overflowX: 'auto' }}>
                          <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.74rem' }}>
                            <thead>
                              <tr>
                                {(analysisWorkspace.rows[0] || []).slice(0, 8).map((cell, idx) => (
                                  <th key={idx} style={{ padding: '8px 10px', textAlign: 'left', color: 'var(--text-muted)', borderBottom: '1px solid var(--border)' }}>{String(cell)}</th>
                                ))}
                              </tr>
                            </thead>
                            <tbody>
                              {analysisWorkspace.rows.slice(1, 13).map((row, rowIdx) => (
                                <tr key={rowIdx}>
                                  {row.slice(0, 8).map((cell: any, colIdx: number) => (
                                    <td key={colIdx} style={{ padding: '8px 10px', color: 'var(--text-secondary)', borderBottom: '1px solid var(--border)' }}>{String(cell ?? '')}</td>
                                  ))}
                                </tr>
                              ))}
                            </tbody>
                          </table>
                        </div>
                      ) : Array.isArray(analysisWorkspace.data) && analysisWorkspace.data.length > 0 ? (
                        <pre style={{ margin: 0, whiteSpace: 'pre-wrap', wordBreak: 'break-word', fontSize: '0.74rem', lineHeight: 1.6, color: 'var(--text-secondary)', fontFamily: 'var(--font-mono, Consolas, monospace)' }}>
                          {JSON.stringify(analysisWorkspace.data, null, 2)}
                        </pre>
                      ) : (
                        <div style={{ color: 'var(--text-secondary)', fontSize: '0.8rem' }}>No preview rows were returned for this result.</div>
                      )}
                    </div>
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
                  <div style={{ fontSize: '0.72rem', color: '#bae6fd', textTransform: 'uppercase', letterSpacing: '0.08em', marginBottom: 6 }}>{effectiveOffline ? 'Offline Chart Workspace' : 'AI Chart Studio'}</div>
                  <div style={{ fontSize: '1.02rem', color: '#f8fafc', fontWeight: 700, marginBottom: 6 }}>{effectiveOffline ? 'Build workbook visuals without cloud AI' : 'Build visuals from chat or refine them here'}</div>
                  <div style={{ fontSize: '0.8rem', color: '#e0f2fe', lineHeight: 1.65 }}>
                    {effectiveOffline ? 'Use local chart commands here or in chat. The selected workbook stays available for offline chart exports and quick visuals.' : 'Say "create chart" in chat and it will open this tab with a prepared direction. You can switch chart style, change the data slice, then apply the preview.'}
                  </div>
                </div>
                <div style={{ padding: '16px 18px', borderRadius: 16, background: 'var(--surface)', border: '1px solid var(--border)' }}>
                  <div style={{ fontSize: '0.72rem', color: 'var(--text-muted)', textTransform: 'uppercase', letterSpacing: '0.08em', marginBottom: 8 }}>{effectiveOffline ? 'Offline Chart Commands' : 'Suggested Questions'}</div>
                  <div style={{ display: 'flex', flexWrap: 'wrap', gap: 8 }}>
                    {(workspaceIntent?.tab === 'charts' ? workspaceIntent.questions : effectiveOffline ? ['create chart', 'bar chart', 'line chart', 'chart top categories'] : ['Which metric should I visualize?', 'Compare sections with a bar chart', 'Show score distribution as line', 'Make a pie chart for grades']).map(question => (
                      <button key={question} className="btn btn-outline btn-sm" onClick={() => effectiveOffline ? void handleOfflineCommand(question) : applyIntentQuestion(question)}>
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
                {effectiveOffline && (
                  <button className="btn btn-primary btn-sm" onClick={() => void runPythonAction('chart')} disabled={!!pythonBusy || !pythonOnline || !currentFile}>
                    {pythonBusy === 'chart' ? 'Building Chart...' : 'Run Offline Chart'}
                  </button>
                )}
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

              {effectiveOffline && (
                <div style={{ marginTop: 16, display: 'grid', gridTemplateColumns: '1fr', gap: 12 }}>
                  {pythonOutputs.length > 0 && (
                    <div style={{ background: 'var(--surface)', border: '1px solid var(--border)', borderRadius: 14, padding: 14 }}>
                      <div style={{ fontSize: '0.72rem', color: 'var(--text-muted)', textTransform: 'uppercase', letterSpacing: '0.08em', marginBottom: 8 }}>Offline Exports</div>
                      <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
                        {pythonOutputs.map(output => (
                          <a key={output.path} href={api.pythonDownloadUrl(output.path)} target="_blank" rel="noreferrer" className="btn btn-outline btn-sm" style={{ justifyContent: 'center' }}>
                            Open {output.label}
                          </a>
                        ))}
                      </div>
                    </div>
                  )}
                  <div style={{ background: 'var(--surface)', border: '1px solid var(--border)', borderRadius: 14, padding: 14 }}>
                    <div style={{ fontSize: '0.78rem', color: 'var(--green)', fontWeight: 700, marginBottom: 8 }}>Offline chart result</div>
                    <pre style={{ margin: 0, whiteSpace: 'pre-wrap', wordBreak: 'break-word', fontSize: '0.72rem', lineHeight: 1.55, color: 'var(--text-secondary)', fontFamily: 'var(--font-mono, Consolas, monospace)' }}>
                      {pythonResult ? JSON.stringify(pythonResult, null, 2) : 'Run the offline chart action to generate a local chart file and workbook summary.'}
                    </pre>
                  </div>
                </div>
              )}
            </div>
          )}

          {/* TAB: CREATE NEW */}
          {activeTab === 'create' && (
            <div style={{ flex: 1, overflowY: 'auto', padding: 24, maxWidth: 940, margin: '0 auto', width: '100%' }}>
              <div style={{ padding: '16px 18px', borderRadius: 16, background: effectiveOffline ? 'linear-gradient(135deg, rgba(59,130,246,0.1), rgba(15,23,42,0.92))' : 'linear-gradient(135deg, rgba(168,85,247,0.16), rgba(15,23,42,0.95))', border: effectiveOffline ? '1px solid rgba(96,165,250,0.18)' : '1px solid rgba(196,181,253,0.16)', marginBottom: 18 }}>
                <div style={{ fontSize: '0.72rem', color: effectiveOffline ? '#bfdbfe' : '#ddd6fe', textTransform: 'uppercase', letterSpacing: '0.08em', marginBottom: 6 }}>{effectiveOffline ? 'AI Workspace Paused' : 'AI Creation Studio'}</div>
                <div style={{ fontSize: '1rem', color: '#f5f3ff', fontWeight: 700, marginBottom: 6 }}>{effectiveOffline ? 'Use Offline Lab for local commands and exports' : 'Describe what to create and let AI ask the next smart questions'}</div>
                <div style={{ fontSize: '0.8rem', color: '#ede9fe', lineHeight: 1.65 }}>
                  {effectiveOffline ? 'Cloud generation is paused in offline mode. Your workbook stays available, and the Offline Lab tab beside this one is ready for local analysis, cleanup, charting, reports, and exports.' : 'Use this space for new sheets, transformed tables, ranked reports, or structured exports. Commands from chat can land here with context already filled in.'}
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
              <textarea value={createPrompt} onChange={e => setCreatePrompt(e.target.value)} placeholder={effectiveOffline ? 'AI create is paused while offline. Open Offline Lab for local workbook work.' : 'Describe what you need...'} disabled={effectiveOffline} style={{ width: '100%', minHeight: 100, padding: 12, background: 'var(--surface)', border: '1px solid var(--border)', borderRadius: 8, marginBottom: 12, color: 'var(--text)', opacity: effectiveOffline ? 0.7 : 1 }} />
              <div style={{ background: 'var(--surface)', border: '1px solid var(--border)', borderRadius: 16, padding: 18, marginTop: 18, opacity: effectiveOffline ? 0.78 : 1 }}>
                <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 12, marginBottom: 8 }}>
                  <div>
                    <div style={{ fontSize: '0.72rem', color: 'var(--text-muted)', textTransform: 'uppercase', letterSpacing: '0.08em' }}>{effectiveOffline ? 'Offline Lab Available' : 'Local Workbook Tools'}</div>
                    <div style={{ fontSize: '1rem', color: 'var(--text)', fontWeight: 700, marginTop: 4 }}>{effectiveOffline ? 'Use the dedicated offline tab beside this one' : 'Keep local workbook tools ready'}</div>
                  </div>
                  <div style={{ padding: '6px 10px', borderRadius: 999, background: pythonOnline ? 'rgba(34,197,94,0.14)' : 'rgba(245,158,11,0.14)', color: pythonOnline ? 'var(--green)' : '#f59e0b', fontSize: '0.72rem', fontWeight: 700 }}>
                    {pythonOnline ? 'Ready' : 'Unavailable'}
                  </div>
                </div>
                <div style={{ fontSize: '0.8rem', color: 'var(--text-secondary)', lineHeight: 1.65, marginBottom: 14 }}>
                  {effectiveOffline
                    ? 'Offline mode now has its own workspace. Open Offline Lab to run local commands, cleanup, charting, exports, and workbook analysis.'
                    : 'Online AI mode is active. If the network drops, this panel can still help with local workbook operations through the offline engine.'}
                </div>
                <div style={{ padding: '10px 12px', borderRadius: 12, background: 'var(--bg)', border: '1px solid var(--border)', marginBottom: 14 }}>
                  <div style={{ fontSize: '0.68rem', color: 'var(--text-muted)', textTransform: 'uppercase', marginBottom: 4 }}>Active Workbook</div>
                  <div style={{ fontSize: '0.84rem', color: 'var(--text)', fontWeight: 600 }}>{currentFile?.name || 'Open a workbook first from My Files'}</div>
                </div>
                <div style={{ display: 'grid', gridTemplateColumns: 'repeat(2, minmax(0, 1fr))', gap: 8, marginBottom: 14 }}>
                  {[
                    { id: 'analyze' as PythonAction, label: 'Analyze' },
                    { id: 'read' as PythonAction, label: 'Preview Rows' },
                    { id: 'clean' as PythonAction, label: 'Clean File' },
                    { id: 'chart' as PythonAction, label: 'Chart' },
                    { id: 'report' as PythonAction, label: 'TXT Report' },
                    { id: 'excel' as PythonAction, label: 'Advanced Excel' },
                    { id: 'template' as PythonAction, label: 'Template' },
                  ].map(action => (
                    <button
                      key={action.id}
                      className={action.id === 'template' ? 'btn btn-outline btn-sm' : 'btn btn-secondary btn-sm'}
                      onClick={() => void runPythonAction(action.id)}
                      disabled={effectiveOffline || !!pythonBusy || (!currentFile && action.id !== 'template') || !pythonOnline}
                      style={{ justifyContent: 'center' }}
                    >
                      {pythonBusy === action.id ? 'Working...' : action.label}
                    </button>
                  ))}
                </div>
                <div style={{ marginBottom: 14 }}>
                  <div style={{ fontSize: '0.72rem', color: 'var(--text-muted)', textTransform: 'uppercase', letterSpacing: '0.08em', marginBottom: 6 }}>{effectiveOffline ? 'Offline Command' : 'Prompt-Based Action'}</div>
                  <textarea
                    value={pythonPrompt}
                    onChange={e => {
                      setPythonPrompt(e.target.value)
                      if (effectiveOffline) setCreatePrompt(e.target.value)
                    }}
                    placeholder="Example: create chart, file info, or segment by Region"
                    disabled={effectiveOffline}
                    style={{ width: '100%', minHeight: 78, padding: 10, background: 'var(--bg)', border: '1px solid var(--border)', borderRadius: 10, color: 'var(--text)', marginBottom: 8, opacity: effectiveOffline ? 0.65 : 1 }}
                  />
                  <button className="btn btn-primary btn-sm" onClick={() => void runPythonAction('process')} disabled={effectiveOffline || !!pythonBusy || !currentFile || !pythonOnline || !(pythonPrompt.trim() || createPrompt.trim())} style={{ width: '100%', justifyContent: 'center' }}>
                    {pythonBusy === 'process' ? 'Running Command...' : 'Run Local Prompt'}
                  </button>
                </div>
                {pythonOutputs.length > 0 && (
                  <div style={{ marginBottom: 14 }}>
                    <div style={{ fontSize: '0.72rem', color: 'var(--text-muted)', textTransform: 'uppercase', letterSpacing: '0.08em', marginBottom: 8 }}>Downloads</div>
                    <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
                      {pythonOutputs.map(output => (
                        <a key={output.path} href={api.pythonDownloadUrl(output.path)} target="_blank" rel="noreferrer" className="btn btn-outline btn-sm" style={{ justifyContent: 'center' }}>
                          Open {output.label}
                        </a>
                      ))}
                    </div>
                  </div>
                )}
                <div style={{ padding: '12px 14px', borderRadius: 12, background: 'var(--bg)', border: '1px solid var(--border)' }}>
                  <div style={{ fontSize: '0.72rem', color: 'var(--text-muted)', textTransform: 'uppercase', letterSpacing: '0.08em', marginBottom: 8 }}>{effectiveOffline ? 'Latest Offline Result' : 'Latest Local Result'}</div>
                  <pre style={{ margin: 0, whiteSpace: 'pre-wrap', wordBreak: 'break-word', fontSize: '0.72rem', lineHeight: 1.55, color: 'var(--text-secondary)', fontFamily: 'var(--font-mono, Consolas, monospace)' }}>
                    {pythonResult ? JSON.stringify(pythonResult, null, 2) : 'Run a local workbook action here to see the result inside React.'}
                  </pre>
                </div>
              </div>
              <button
                className={effectiveOffline ? 'btn btn-outline' : 'btn btn-primary'}
                onClick={effectiveOffline ? () => setActiveTab('offline') : handleCreate}
                disabled={creating || !!pythonBusy || (!effectiveOffline && !createPrompt.trim())}
                style={{ width: '100%', padding: 12 }}
              >
                {effectiveOffline
                  ? 'Open Offline Lab'
                  : (creating ? 'Creating...' : 'AI Generate')}
              </button>
            </div>
          )}

          {activeTab === 'offline' && effectiveOffline && (
            <div style={{ flex: 1, overflowY: 'auto', padding: 24, maxWidth: 1040, margin: '0 auto', width: '100%' }}>
              <div style={{ padding: '18px 20px', borderRadius: 20, background: 'linear-gradient(135deg, color-mix(in srgb, var(--surface-3) 84%, var(--orange) 16%), color-mix(in srgb, var(--surface-3) 58%, var(--orange) 42%))', border: '1px solid color-mix(in srgb, var(--border-hi) 76%, var(--orange) 24%)', marginBottom: 18, boxShadow: '0 16px 32px var(--accent-dim2)' }}>
                <div style={{ fontSize: '0.72rem', color: 'var(--orange)', textTransform: 'uppercase', letterSpacing: '0.08em', marginBottom: 6 }}>Offline Lab</div>
                <div style={{ fontSize: '1.08rem', color: 'var(--text)', fontWeight: 700, marginBottom: 8 }}>Work locally while AI mode is unavailable</div>
                <div style={{ fontSize: '0.82rem', color: 'var(--text-secondary)', lineHeight: 1.7 }}>
                  {offlineReason} Use this workspace for chart exports, cleanup, analysis, templates, and command-based workbook processing.
                </div>
              </div>

              <div style={{ display: 'grid', gridTemplateColumns: '1.15fr 0.85fr', gap: 16, marginBottom: 18 }}>
                <div style={{ background: 'var(--surface-2)', border: '1px solid var(--border-hi)', borderRadius: 18, padding: 18, boxShadow: '0 12px 28px var(--accent-dim2)' }}>
                  <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 12, marginBottom: 10 }}>
                    <div>
                      <div style={{ fontSize: '0.72rem', color: 'var(--text-muted)', textTransform: 'uppercase', letterSpacing: '0.08em' }}>Mode Status</div>
                      <div style={{ fontSize: '1rem', color: 'var(--text)', fontWeight: 700, marginTop: 4 }}>Offline tools are {pythonOnline ? 'ready' : 'not connected'}</div>
                    </div>
                    <div style={{ padding: '7px 12px', borderRadius: 999, background: pythonOnline ? 'rgba(34,197,94,0.14)' : 'rgba(239,68,68,0.12)', color: pythonOnline ? 'var(--green)' : 'var(--red)', fontSize: '0.72rem', fontWeight: 700 }}>
                      {pythonOnline ? 'Local Engine Ready' : 'Engine Offline'}
                    </div>
                  </div>
                  <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, minmax(0, 1fr))', gap: 10, marginBottom: 12 }}>
                    <div style={{ padding: 12, borderRadius: 14, background: 'var(--surface-3)', border: '1px solid var(--border-hi)' }}>
                      <div style={{ fontSize: '0.68rem', color: 'var(--text-muted)', textTransform: 'uppercase' }}>Current Mode</div>
                      <div style={{ fontSize: '0.92rem', color: '#b45309', fontWeight: 700, marginTop: 6 }}>Offline</div>
                    </div>
                    <div style={{ padding: 12, borderRadius: 14, background: 'var(--surface-3)', border: '1px solid var(--border-hi)' }}>
                      <div style={{ fontSize: '0.68rem', color: 'var(--text-muted)', textTransform: 'uppercase' }}>Workbook</div>
                      <div style={{ fontSize: '0.92rem', color: 'var(--text)', fontWeight: 700, marginTop: 6 }}>{currentFile?.name || 'No workbook open'}</div>
                    </div>
                    <div style={{ padding: 12, borderRadius: 14, background: 'var(--surface-3)', border: '1px solid var(--border-hi)' }}>
                      <div style={{ fontSize: '0.68rem', color: 'var(--text-muted)', textTransform: 'uppercase' }}>Assistant Panel</div>
                      <div style={{ fontSize: '0.92rem', color: 'var(--text)', fontWeight: 700, marginTop: 6 }}>Offline active</div>
                    </div>
                  </div>
                  <textarea
                    value={pythonPrompt}
                    onChange={e => {
                      setPythonPrompt(e.target.value)
                      setCreatePrompt(e.target.value)
                    }}
                    placeholder="Type an offline command like analyze data, create chart, clean data, or advanced excel"
                    style={{ width: '100%', minHeight: 96, padding: 12, background: 'var(--bg-3)', border: '1px solid var(--border-hi)', borderRadius: 12, color: 'var(--text)', marginBottom: 10 }}
                  />
                  <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap' }}>
                    <button className="btn btn-primary btn-sm" onClick={() => void runPythonAction('process')} disabled={!!pythonBusy || !currentFile || !pythonOnline || !pythonPrompt.trim()}>
                      {pythonBusy === 'process' ? 'Running...' : 'Run Offline Command'}
                    </button>
                    <button className="btn btn-outline btn-sm" onClick={() => void runPythonAction('analyze')} disabled={!!pythonBusy || !currentFile || !pythonOnline}>
                      {pythonBusy === 'analyze' ? 'Analyzing...' : 'Quick Analyze'}
                    </button>
                    <button className="btn btn-outline btn-sm" onClick={() => void runPythonAction('chart')} disabled={!!pythonBusy || !currentFile || !pythonOnline}>
                      {pythonBusy === 'chart' ? 'Building...' : 'Quick Chart'}
                    </button>
                    <button className="btn btn-outline btn-sm" onClick={() => void runPythonAction('clean')} disabled={!!pythonBusy || !currentFile || !pythonOnline}>
                      {pythonBusy === 'clean' ? 'Cleaning...' : 'Quick Clean'}
                    </button>
                  </div>
                </div>

                <div style={{ background: 'var(--surface-2)', border: '1px solid var(--border-hi)', borderRadius: 18, padding: 18, boxShadow: '0 12px 28px var(--accent-dim2)' }}>
                  <div style={{ fontSize: '0.72rem', color: 'var(--text-muted)', textTransform: 'uppercase', letterSpacing: '0.08em', marginBottom: 10 }}>Recommended Commands</div>
                  <div style={{ display: 'grid', gap: 8, maxHeight: 420, overflowY: 'auto' }}>
                    {OFFLINE_COMMANDS.slice(0, 18).map(item => (
                      <button
                        key={item.command}
                        onClick={() => {
                          setPythonPrompt(item.command)
                          setCreatePrompt(item.command)
                        }}
                        style={{ textAlign: 'left', padding: '10px 12px', background: 'var(--surface-3)', border: '1px solid var(--border-hi)', borderRadius: 12, cursor: 'pointer' }}
                      >
                        <div style={{ fontSize: '0.78rem', color: 'var(--text)', fontWeight: 600 }}>{item.command}</div>
                        <div style={{ fontSize: '0.7rem', color: 'var(--text-secondary)', marginTop: 4 }}>{item.effect}</div>
                      </button>
                    ))}
                  </div>
                </div>
              </div>

              <div style={{ display: 'grid', gridTemplateColumns: pythonOutputs.length > 0 ? '0.9fr 1.1fr' : '1fr', gap: 16 }}>
                {pythonOutputs.length > 0 && (
                  <div style={{ background: 'var(--surface-2)', border: '1px solid var(--border-hi)', borderRadius: 18, padding: 18, boxShadow: '0 12px 28px var(--accent-dim2)' }}>
                    <div style={{ fontSize: '0.72rem', color: 'var(--text-muted)', textTransform: 'uppercase', letterSpacing: '0.08em', marginBottom: 10 }}>Generated Files</div>
                    <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
                      {pythonOutputs.map(output => (
                        <a key={output.path} href={api.pythonDownloadUrl(output.path)} target="_blank" rel="noreferrer" className="btn btn-outline btn-sm" style={{ justifyContent: 'center' }}>
                          Open {output.label}
                        </a>
                      ))}
                    </div>
                  </div>
                )}
                <div style={{ background: 'var(--surface-2)', border: '1px solid var(--border-hi)', borderRadius: 18, padding: 18, boxShadow: '0 12px 28px var(--accent-dim2)' }}>
                  <div style={{ fontSize: '0.72rem', color: 'var(--text-muted)', textTransform: 'uppercase', letterSpacing: '0.08em', marginBottom: 10 }}>Latest Offline Result</div>
                  <pre style={{ margin: 0, whiteSpace: 'pre-wrap', wordBreak: 'break-word', fontSize: '0.73rem', lineHeight: 1.6, color: 'var(--text-secondary)', fontFamily: 'var(--font-mono, Consolas, monospace)' }}>
                    {pythonResult ? JSON.stringify(pythonResult, null, 2) : 'Run an offline command to see local workbook results here.'}
                  </pre>
                </div>
              </div>
            </div>
          )}

        </main>

        <div
          onMouseDown={() => setResizingPanel(true)}
          style={{
            width: 6,
            cursor: 'col-resize',
            flexShrink: 0,
            background: resizingPanel ? 'rgba(59,130,246,0.18)' : 'transparent',
            borderLeft: '1px solid rgba(148,163,184,0.16)'
          }}
        />

        <AIChatPanel 
          currentFile={currentFile}
          rawRows={rawGridValues}
          stats={curStats}
          panelWidth={aiPanelWidth}
          onGridUpdate={handleAIGridUpdate}
          onNewFile={handleOpenGeneratedFile}
          onShowChart={handleAIChartRequest}
          onAnalysisResult={handleChatAnalysisResult}
          onDocumentResult={handleChatDocumentResult}
          onApplyPreview={applyAiPreview}
          onDiscardPreview={discardAiPreview}
          onUndoLastChange={undoAiApply}
        />
      </div>
    </div>
  )
}

