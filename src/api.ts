const BASE = 'http://localhost:3001/api'
const PYTHON_BASE = 'http://127.0.0.1:8001'

async function req(method: string, url: string, body?: any, isFormData = false) {
  const opts: RequestInit = {
    method,
    headers: isFormData ? {} : { 'Content-Type': 'application/json' }
  }
  if (body) opts.body = isFormData ? body : JSON.stringify(body)

  console.log(`[API] ${method} ${BASE}${url}`)

  try {
    const res = await fetch(`${BASE}${url}`, opts)
    if (!res.ok) {
      const errText = await res.text()
      console.error(`[API] Error ${res.status}: ${errText}`)
      throw new Error(`HTTP ${res.status}: ${errText.slice(0, 100)}`)
    }
    return await res.json()
  } catch (err: any) {
    console.error(`[API] Request failed:`, err)
    throw err
  }
}

export const api = {
  health: () => req('GET', '/health'),
  pythonHealth: () => reqAbsolute(PYTHON_BASE, 'GET', '/'),

  // Auth
  login: (email: string, password: string) => req('POST', '/auth/login', { email, password }),
  register: (name: string, email: string, department: string) => req('POST', '/auth/register', { name, email, department }),

  // Files
  uploadFiles: async (files: File[], department: string, uploadedBy: string) => {
    const formData = new FormData()
    files.forEach(file => {
      formData.append('files', file)
      console.log(`[API] Adding file: ${file.name} (${file.size} bytes)`)
    })
    
    const url = `/files/upload?department=${encodeURIComponent(department)}&uploadedBy=${encodeURIComponent(uploadedBy)}`
    console.log(`[API] Uploading to: ${BASE}${url}`)
    
    const res = await fetch(`${BASE}${url}`, {
      method: 'POST',
      body: formData,
    })
    
    if (!res.ok) {
      const err = await res.text()
      throw new Error(`Upload failed: ${err}`)
    }
    return res.json()
  },
  
  getFiles: (params?: { department?: string; type?: string; search?: string }) => {
    const q = new URLSearchParams(params as any).toString()
    return req('GET', `/files${q ? '?' + q : ''}`)
  },
  
  deleteFile: (id: string) => req('DELETE', `/files/${id}`),
  browseDir: (path: string) => req('GET', `/files/browse?path=${encodeURIComponent(path)}`),
  mkdir: (path: string) => req('POST', '/files/mkdir', { path }),
  searchDisk: (query: string, directory?: string) => {
    const q = `query=${encodeURIComponent(query)}${directory ? '&directory=' + encodeURIComponent(directory) : ''}`
    return req('GET', `/files/search-disk?${q}`)
  },

  // Excel
  getExcelRaw: (fileId: string) => req('GET', `/excel/raw/${fileId}`),
  analyzeFile: (fileId: string) => req('GET', `/excel/analyze/${fileId}`),
  analyzeUpload: (file: File) => {
    const fd = new FormData()
    fd.append('file', file)
    return req('POST', '/excel/analyze', fd, true)
  },
  compareFiles: (fileId1: string, fileId2: string) => req('POST', '/excel/compare', { fileId1, fileId2 }),
  excelCommand: (fileId: string, command: string) => req('POST', '/excel/command', { fileId, command }),
  saveExcel: (fileId: string, rows: any[], sheetName?: string) => req('POST', `/excel/save/${fileId}`, { rows, sheetName }),
  generateExcel: (data: any[], filename: string, sheetName?: string, department?: string) =>
    req('POST', '/excel/generate', { data, filename, sheetName, department }),

  // Workflows
  getWorkflows: () => req('GET', '/workflows'),
  createWorkflow: (data: any) => req('POST', '/workflows', data),
  updateWorkflow: (id: string, data: any) => req('PATCH', `/workflows/${id}`, data),
  deleteWorkflow: (id: string) => req('DELETE', `/workflows/${id}`),
  runWorkflow: (id: string) => req('POST', `/workflows/${id}/run`),

  // Documents
  generateDoc: (type: string, data: any, title: string) => req('POST', '/documents/generate', { type, data, title }),
  generatePpt: (payload: { prompt: string; data?: any; audience?: string; theme?: string; goal?: string; department?: string }) =>
    req('POST', '/ppt/generate', payload),

  // Python engine
  pythonRead: (file: File, limit = 25) => {
    const fd = new FormData()
    fd.append('file', file)
    return reqAbsolute(PYTHON_BASE, 'POST', `/read?limit=${limit}`, fd, true)
  },
  pythonAnalyze: (file: File) => {
    const fd = new FormData()
    fd.append('file', file)
    return reqAbsolute(PYTHON_BASE, 'POST', '/analyze', fd, true)
  },
  pythonChart: (file: File, chartType = 'auto') => {
    const fd = new FormData()
    fd.append('file', file)
    return reqAbsolute(PYTHON_BASE, 'POST', `/chart?chart_type=${encodeURIComponent(chartType)}`, fd, true)
  },
  pythonClean: (file: File) => {
    const fd = new FormData()
    fd.append('file', file)
    return reqAbsolute(PYTHON_BASE, 'POST', '/clean', fd, true)
  },
  pythonReport: (file: File) => {
    const fd = new FormData()
    fd.append('file', file)
    return reqAbsolute(PYTHON_BASE, 'POST', '/report', fd, true)
  },
  pythonWord: (file: File) => {
    const fd = new FormData()
    fd.append('file', file)
    return reqAbsolute(PYTHON_BASE, 'POST', '/word', fd, true)
  },
  pythonPpt: (file: File) => {
    const fd = new FormData()
    fd.append('file', file)
    return reqAbsolute(PYTHON_BASE, 'POST', '/ppt', fd, true)
  },
  pythonAdvancedExcel: (file: File) => {
    const fd = new FormData()
    fd.append('file', file)
    return reqAbsolute(PYTHON_BASE, 'POST', '/excel-advanced', fd, true)
  },
  pythonTemplate: (rows = 10, includeSampleData = true) =>
    reqAbsolute(PYTHON_BASE, 'POST', `/template?rows=${rows}&include_sample_data=${includeSampleData}`),
  pythonProcess: (file: File, prompt: string) => {
    const fd = new FormData()
    fd.append('file', file)
    return reqAbsolute(PYTHON_BASE, 'POST', `/process?prompt=${encodeURIComponent(prompt)}`, fd, true)
  },
  pythonDownloadUrl: (path: string) => `${PYTHON_BASE}/download?path=${encodeURIComponent(path)}`,

  // Dashboard
  getDashboardStats: () => req('GET', '/dashboard/stats'),
  getUsers: () => req('GET', '/users'),
}

async function reqAbsolute(base: string, method: string, url: string, body?: any, isFormData = false) {
  const opts: RequestInit = {
    method,
    headers: isFormData ? {} : { 'Content-Type': 'application/json' }
  }
  if (body) opts.body = isFormData ? body : JSON.stringify(body)

  try {
    const res = await fetch(`${base}${url}`, opts)
    if (!res.ok) {
      const errText = await res.text()
      throw new Error(`HTTP ${res.status}: ${errText.slice(0, 140)}`)
    }
    return await res.json()
  } catch (err: any) {
    console.error(`[API] Absolute request failed:`, err)
    throw err
  }
}
