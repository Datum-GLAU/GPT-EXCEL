import { HashRouter, Routes, Route } from 'react-router-dom'
import { useEffect } from 'react'
import Loading from './pages/Loading'
import GetStarted from './pages/GetStarted'
import Auth from './pages/Auth'
import Dashboard from './pages/Dashboard'
import ExcelSheet from './pages/ExcelSheet'
import Settings from './pages/Settings'
import FileManager from './pages/FileManager'
import Workflow from './pages/Workflow'
import Documents from './pages/Documents'
import VoiceAssistant from './pages/VoiceAssistant'
import PowerPointStudio from './pages/PowerPointStudio'
import PluginMarketplace from './pages/PluginMarketplace'

export default function App() {
  useEffect(() => {
    const saved = localStorage.getItem('theme')
    if (saved === 'light') {
      document.documentElement.classList.add('light')
      document.documentElement.classList.remove('dark')
    } else {
      document.documentElement.classList.add('dark')
      document.documentElement.classList.remove('light')
    }
  }, [])

  return (
    <HashRouter>
      <Routes>
        <Route path="/" element={<Loading />} />
        <Route path="/get-started" element={<GetStarted />} />
        <Route path="/auth" element={<Auth />} />
        <Route path="/dashboard" element={<Dashboard />} />
        <Route path="/excel" element={<ExcelSheet />} />
        <Route path="/settings" element={<Settings />} />
        <Route path="/file-manager" element={<FileManager />} />
        <Route path="/workflow" element={<Workflow />} />
        <Route path="/documents" element={<Documents />} />
        <Route path="/voice" element={<VoiceAssistant />} />
        <Route path="/powerpoint" element={<PowerPointStudio />} />
        <Route path="/plugins" element={<PluginMarketplace />} />
      </Routes>
    </HashRouter>
  )
}
