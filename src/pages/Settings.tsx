import { useState } from 'react'
import { useSelector, useDispatch } from 'react-redux'
import { RootState, setTheme, setUser } from '../index'
import Sidebar from '../components/Sidebar'
import Header from '../components/Header'

const DEPARTMENTS = ['CSE','ECE','EEE','MECH','CIVIL','IT','MBA','MCA','PHY','CHEM','MATH','ENGLISH','Administration']
const ROLES = ['faculty','hod','admin','principal','lab_assistant']

export default function Settings() {
  const dispatch = useDispatch()
  const [sidebarOpen, setSidebarOpen] = useState(true)
  const theme = useSelector((s: RootState) => s.app.theme)
  const user = useSelector((s: RootState) => s.app.user)
  const [name, setName] = useState(user?.name || '')
  const [email, setEmail] = useState(user?.email || '')
  const [department, setDepartment] = useState(user?.department || '')
  const [apiKey, setApiKey] = useState(localStorage.getItem('openai_key') || '')
  const [showKey, setShowKey] = useState(false)
  const [passThreshold, setPassThreshold] = useState(localStorage.getItem('pass_threshold') || '40')
  const [attendanceThreshold, setAttendanceThreshold] = useState(localStorage.getItem('attendance_threshold') || '75')
  const [saved, setSaved] = useState(false)

  const handleTheme = () => {
    const t = theme === 'dark' ? 'light' : 'dark'
    dispatch(setTheme(t))
    document.documentElement.classList.toggle('dark', t === 'dark')
    document.documentElement.classList.toggle('light', t === 'light')
    localStorage.setItem('theme', t)
  }

  const saveProfile = () => {
    dispatch(setUser({ name, email, plan: user?.plan || 'free', department } as any))
    setSaved(true); setTimeout(() => setSaved(false), 2000)
  }

  const saveAnalysis = () => {
    localStorage.setItem('pass_threshold', passThreshold)
    localStorage.setItem('attendance_threshold', attendanceThreshold)
    if (apiKey) localStorage.setItem('openai_key', apiKey)
    setSaved(true); setTimeout(() => setSaved(false), 2000)
  }

  const Section = ({ title, children }: { title: string; children: React.ReactNode }) => (
    <div style={{ background: 'var(--surface)', border: '1px solid var(--border)', borderRadius: 12, padding: '18px 22px', marginBottom: 16 }}>
      <h2 style={{ fontSize: '0.9rem', fontWeight: 600, marginBottom: 16, color: 'var(--text)' }}>{title}</h2>
      {children}
    </div>
  )

  const Row = ({ label, sub, children }: { label: string; sub?: string; children: React.ReactNode }) => (
    <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '10px 0', borderBottom: '1px solid var(--border)' }}>
      <div>
        <div style={{ fontSize: '0.82rem', fontWeight: 500, color: 'var(--text)' }}>{label}</div>
        {sub && <div style={{ fontSize: '0.72rem', color: 'var(--text-muted)', marginTop: 2 }}>{sub}</div>}
      </div>
      <div>{children}</div>
    </div>
  )

  return (
    <div style={{ display: 'flex', flexDirection: 'column', height: '100vh', background: 'var(--bg)' }}>
      <Header toggleSidebar={() => setSidebarOpen(p => !p)} />
      <div style={{ display: 'flex', flex: 1, overflow: 'hidden' }}>
        <Sidebar isOpen={sidebarOpen} />
        <main style={{ flex: 1, overflowY: 'auto', padding: '24px 32px', background: 'var(--bg-2)' }}>
          <div style={{ maxWidth: 640 }}>
            <h1 style={{ fontSize: '1.4rem', letterSpacing: -0.5, marginBottom: 24 }}>Settings</h1>

            {saved && <div style={{ padding: '8px 14px', background: 'rgba(34,197,94,0.1)', border: '1px solid rgba(34,197,94,0.3)', borderRadius: 8, fontSize: '0.78rem', color: 'var(--green)', marginBottom: 16 }}>✓ Saved successfully</div>}

            <Section title="Your Profile">
              <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
                <div className="input-wrap"><label className="input-label">Full Name</label><input className="input" value={name} onChange={e => setName(e.target.value)} /></div>
                <div className="input-wrap"><label className="input-label">Email</label><input className="input" type="email" value={email} onChange={e => setEmail(e.target.value)} /></div>
                <div className="input-wrap">
                  <label className="input-label">Department</label>
                  <select className="input" value={department} onChange={e => setDepartment(e.target.value)}>
                    <option value="">Select department</option>
                    {DEPARTMENTS.map(d => <option key={d}>{d}</option>)}
                  </select>
                </div>
                <div style={{ display: 'flex', justifyContent: 'flex-end' }}>
                  <button className="btn btn-primary btn-sm" onClick={saveProfile}>Save Profile</button>
                </div>
              </div>
            </Section>

            <Section title="Appearance">
              <Row label="Dark Mode" sub="Toggle between dark and light theme">
                <button onClick={handleTheme} className={`switch ${theme === 'dark' ? 'on' : ''}`} style={{ width: 44, height: 24 }} />
              </Row>
            </Section>

            <Section title="Analysis Settings">
              <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
                <div className="input-wrap">
                  <label className="input-label">Pass Mark Threshold</label>
                  <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
                    <input className="input" type="number" min="0" max="100" value={passThreshold} onChange={e => setPassThreshold(e.target.value)} style={{ width: 80 }} />
                    <span style={{ fontSize: '0.78rem', color: 'var(--text-muted)' }}>Students scoring below this are marked as Failed</span>
                  </div>
                </div>
                <div className="input-wrap">
                  <label className="input-label">Attendance Alert Threshold (%)</label>
                  <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
                    <input className="input" type="number" min="0" max="100" value={attendanceThreshold} onChange={e => setAttendanceThreshold(e.target.value)} style={{ width: 80 }} />
                    <span style={{ fontSize: '0.78rem', color: 'var(--text-muted)' }}>Flag students below this attendance %</span>
                  </div>
                </div>
                <div className="input-wrap">
                  <label className="input-label">OpenAI API Key (optional)</label>
                  <div style={{ position: 'relative' }}>
                    <input className="input" type={showKey ? 'text' : 'password'} placeholder="sk-..." value={apiKey} onChange={e => setApiKey(e.target.value)} />
                    <button type="button" onClick={() => setShowKey(p => !p)} style={{ position: 'absolute', right: 12, top: '50%', transform: 'translateY(-50%)', background: 'none', border: 'none', color: 'var(--text-muted)', fontSize: '0.65rem', cursor: 'pointer' }}>{showKey ? 'HIDE' : 'SHOW'}</button>
                  </div>
                  <div style={{ fontSize: '0.7rem', color: 'var(--text-muted)', marginTop: 4 }}>Used for AI-powered document generation</div>
                </div>
                <div style={{ display: 'flex', justifyContent: 'flex-end' }}>
                  <button className="btn btn-primary btn-sm" onClick={saveAnalysis}>Save Settings</button>
                </div>
              </div>
            </Section>

            <Section title="Backend Server">
              <Row label="Server URL" sub="Express backend running locally">
                <code style={{ fontSize: '0.75rem', color: 'var(--accent)', background: 'var(--surface-2)', padding: '2px 8px', borderRadius: 4 }}>localhost:3001</code>
              </Row>
              <Row label="Start Backend" sub="Run this in a separate terminal">
                <code style={{ fontSize: '0.72rem', color: 'var(--green)', background: 'var(--surface-2)', padding: '3px 10px', borderRadius: 4 }}>cd server && npm install && npm start</code>
              </Row>
            </Section>

            <Section title="About">
              <div style={{ fontSize: '0.78rem', color: 'var(--text-secondary)', lineHeight: 1.7 }}>
                <strong style={{ color: 'var(--text)' }}>XtronExcel v2.0</strong> — University Data Management Platform<br />
                Built with Electron, React, TypeScript, Express.js, XLSX<br />
                Designed for GLA University — Datum_GLAU<br />
                <span style={{ color: 'var(--text-muted)' }}>© 2026 Mallavarapu Krishna Koushik Reddy et al.</span>
              </div>
            </Section>
          </div>
        </main>
      </div>
    </div>
  )
}
