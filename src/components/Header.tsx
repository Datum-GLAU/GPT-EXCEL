import { useState, useEffect } from 'react'
import { useNavigate } from 'react-router-dom'

interface Props { toggleSidebar: () => void }

export default function Header({ toggleSidebar }: Props) {
  const nav = useNavigate()
  const [dark, setDark] = useState(() => !document.documentElement.classList.contains('light'))
  const [cmd, setCmd] = useState(false)
  const [cmdQ, setCmdQ] = useState('')
  const [notifCount, setNotifCount] = useState(3)
  const [showNotifs, setShowNotifs] = useState(false)
  const [showProfile, setShowProfile] = useState(false)
  const user = (() => { try { return JSON.parse(localStorage.getItem('gpe_user') || 'null') } catch { return null } })()

  useEffect(() => {
    if (dark) {
      document.documentElement.classList.add('dark')
      document.documentElement.classList.remove('light')
      localStorage.setItem('theme', 'dark')
    } else {
      document.documentElement.classList.add('light')
      document.documentElement.classList.remove('dark')
      localStorage.setItem('theme', 'light')
    }
  }, [dark])

  useEffect(() => {
    const fn = (e: KeyboardEvent) => {
      if ((e.metaKey || e.ctrlKey) && e.key === 'k') { e.preventDefault(); setCmd(true) }
      if (e.key === 'Escape') { setCmd(false); setShowNotifs(false); setShowProfile(false) }
    }
    window.addEventListener('keydown', fn)
    return () => window.removeEventListener('keydown', fn)
  }, [])

  const notifs = [
    { id: 1, icon: '⊞', text: 'Excel file generated', sub: 'Q4 Model.xlsx · 2m ago', type: 'green', read: false },
    { id: 2, icon: '⌘', text: 'Workflow completed', sub: 'Weekly Report · 1h ago', type: 'blue', read: false },
    { id: 3, icon: '◬', text: 'API at 89% limit', sub: 'Upgrade for unlimited · 3h ago', type: 'yellow', read: false },
  ]

  const pages = [
    { icon: '◱', label: 'Dashboard', path: '/dashboard' },
    { icon: '⊞', label: 'Excel Sheet', path: '/excel' },
    { icon: '⚙', label: 'Settings', path: '/settings' },
  ]

  const initials = user?.name ? user.name.split(' ').map((n: string) => n[0]).join('').slice(0, 2).toUpperCase() : 'G'

  return (
    <>
      <header style={{ height: 44, display: 'flex', alignItems: 'center', padding: '0 14px', background: 'var(--bg)', borderBottom: '1px solid var(--border)', flexShrink: 0, gap: 8, zIndex: 40, WebkitAppRegion: 'drag' as any }}>

        {/* Sidebar toggle */}
        <button className="btn btn-icon-sm btn-ghost" onClick={toggleSidebar} style={{ WebkitAppRegion: 'no-drag' as any }} data-tip="Toggle sidebar">
          <svg width={14} height={14} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth={2}>
            <line x1={3} y1={6} x2={21} y2={6}/><line x1={3} y1={12} x2={21} y2={12}/><line x1={3} y1={18} x2={21} y2={18}/>
          </svg>
        </button>

        {/* Logo */}
        <div style={{ display: 'flex', alignItems: 'center', gap: 4, WebkitAppRegion: 'no-drag' as any, cursor: 'pointer' }} onClick={() => nav('/dashboard')}>
          <span style={{ fontFamily: 'var(--font-display)', fontWeight: 800, fontSize: 13, letterSpacing: -0.5, color: 'var(--text)' }}>XtronExcel</span>
        </div>

        {/* Command search */}
        <button onClick={() => setCmd(true)} style={{ flex: 1, maxWidth: 320, WebkitAppRegion: 'no-drag' as any, display: 'flex', alignItems: 'center', gap: 8, height: 28, padding: '0 10px', border: '1px solid var(--border)', borderRadius: 6, background: 'var(--surface-2)', cursor: 'text', fontSize: '0.75rem', color: 'var(--text-muted)', transition: 'all var(--tr)' }}
          onMouseEnter={e => (e.currentTarget.style.borderColor = 'var(--border-hi)')}
          onMouseLeave={e => (e.currentTarget.style.borderColor = 'var(--border)')}
        >
          <svg width={12} height={12} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth={2.5}><circle cx={11} cy={11} r={8}/><path d="m21 21-4.35-4.35"/></svg>
          <span style={{ flex: 1, textAlign: 'left', color: 'var(--text-muted)' }}>Search or command...</span>
          <div style={{ display: 'flex', gap: 2 }}><kbd style={{ fontSize: '0.6rem' }}>⌘</kbd><kbd style={{ fontSize: '0.6rem' }}>K</kbd></div>
        </button>

        {/* Right controls */}
        <div style={{ display: 'flex', gap: 3, marginLeft: 'auto', alignItems: 'center', WebkitAppRegion: 'no-drag' as any }}>

          {/* Theme toggle */}
          <button className="btn btn-icon-sm btn-ghost" onClick={() => setDark(d => !d)} data-tip={dark ? 'Light mode' : 'Dark mode'}
            style={{ color: 'var(--text-secondary)' }}
          >
            {dark ? (
              <svg width={14} height={14} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth={2}>
                <circle cx={12} cy={12} r={5}/>
                <line x1={12} y1={1} x2={12} y2={3}/><line x1={12} y1={21} x2={12} y2={23}/>
                <line x1={4.22} y1={4.22} x2={5.64} y2={5.64}/><line x1={18.36} y1={18.36} x2={19.78} y2={19.78}/>
                <line x1={1} y1={12} x2={3} y2={12}/><line x1={21} y1={12} x2={23} y2={12}/>
                <line x1={4.22} y1={19.78} x2={5.64} y2={18.36}/><line x1={18.36} y1={5.64} x2={19.78} y2={4.22}/>
              </svg>
            ) : (
              <svg width={14} height={14} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth={2}>
                <path d="M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z"/>
              </svg>
            )}
          </button>

          {/* Notifications */}
          <div style={{ position: 'relative' }}>
            <button className="btn btn-icon-sm btn-ghost" onClick={() => { setShowNotifs(p => !p); setShowProfile(false) }} style={{ position: 'relative', color: 'var(--text-secondary)' }}>
              <svg width={14} height={14} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth={2}>
                <path d="M18 8A6 6 0 0 0 6 8c0 7-3 9-3 9h18s-3-2-3-9"/><path d="M13.73 21a2 2 0 0 1-3.46 0"/>
              </svg>
              {notifCount > 0 && <span style={{ position: 'absolute', top: 2, right: 2, width: 7, height: 7, borderRadius: '50%', background: 'var(--red)', border: '1.5px solid var(--bg)' }}/>}
            </button>
            {showNotifs && (
              <div style={{ position: 'absolute', right: 0, top: 'calc(100% + 6px)', width: 300, background: 'var(--surface)', border: '1px solid var(--border)', borderRadius: 10, boxShadow: 'var(--shadow-lg)', zIndex: 200, animation: 'menuIn 0.14s ease', overflow: 'hidden' }}>
                <div style={{ padding: '10px 14px', borderBottom: '1px solid var(--border)', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                  <span style={{ fontSize: '0.8rem', fontWeight: 600, color: 'var(--text)' }}>Notifications</span>
                  <button className="btn btn-xs btn-ghost" onClick={() => setNotifCount(0)}>Mark all read</button>
                </div>
                {notifs.map(n => (
                  <div key={n.id} style={{ display: 'flex', gap: 10, padding: '10px 14px', borderBottom: '1px solid var(--border)', background: n.read ? 'transparent' : 'var(--accent-dim2)', cursor: 'pointer', transition: 'background var(--tr)' }}
                    onMouseEnter={e => (e.currentTarget.style.background = 'var(--accent-dim)')}
                    onMouseLeave={e => (e.currentTarget.style.background = n.read ? 'transparent' : 'var(--accent-dim2)')}
                  >
                    <div style={{ width: 28, height: 28, flexShrink: 0, display: 'flex', alignItems: 'center', justifyContent: 'center', background: `var(--${n.type}-dim)`, border: `1px solid var(--${n.type}-border)`, borderRadius: 6, fontSize: 13, color: `var(--${n.type})` }}>{n.icon}</div>
                    <div style={{ flex: 1 }}>
                      <div style={{ fontSize: '0.775rem', fontWeight: 500, color: 'var(--text)', lineHeight: 1.4 }}>{n.text}</div>
                      <div style={{ fontSize: '0.7rem', color: 'var(--text-muted)', marginTop: 2, fontFamily: 'var(--font-mono)' }}>{n.sub}</div>
                    </div>
                    {!n.read && <div style={{ width: 6, height: 6, borderRadius: '50%', background: 'var(--blue)', flexShrink: 0, marginTop: 4 }}/>}
                  </div>
                ))}
                <div style={{ padding: '8px 14px' }}>
                  <button className="btn btn-ghost btn-xs w-full">View all notifications</button>
                </div>
              </div>
            )}
          </div>

          {/* Profile */}
          <div style={{ position: 'relative' }}>
            <button className="btn btn-icon-sm btn-ghost" onClick={() => { setShowProfile(p => !p); setShowNotifs(false) }}>
              <div style={{ width: 22, height: 22, borderRadius: '50%', background: 'var(--blue)', display: 'flex', alignItems: 'center', justifyContent: 'center', color: '#fff', fontSize: '0.6rem', fontWeight: 800 }}>{initials}</div>
            </button>
            {showProfile && (
              <div style={{ position: 'absolute', right: 0, top: 'calc(100% + 6px)', width: 220, background: 'var(--surface)', border: '1px solid var(--border)', borderRadius: 10, boxShadow: 'var(--shadow-lg)', zIndex: 200, animation: 'menuIn 0.14s ease', padding: 4 }}>
                <div style={{ padding: '10px 12px', borderBottom: '1px solid var(--border)', marginBottom: 4 }}>
                  <div style={{ fontWeight: 600, fontSize: '0.8rem', color: 'var(--text)' }}>{user?.name || 'Guest User'}</div>
                  <div style={{ fontSize: '0.72rem', color: 'var(--text-muted)' }}>{user?.email || 'guest@demo.com'} · {user?.plan || 'Free'}</div>
                </div>
                {[{ icon: '◱', label: 'Profile' }, { icon: '⊞', label: 'API Keys' }, { icon: '◈', label: 'Integrations' }, { icon: '⌘', label: 'Shortcuts' }].map(item => (
                  <div key={item.label} className="ctx-item" onClick={() => { setShowProfile(false); nav('/settings') }}>
                    <span style={{ fontSize: 12, color: 'var(--text-secondary)' }}>{item.icon}</span>
                    <span style={{ fontSize: '0.8rem', color: 'var(--text)' }}>{item.label}</span>
                  </div>
                ))}
                <div className="ctx-separator"/>
                <div className="ctx-item" onClick={() => { setShowProfile(false); nav('/get-started') }}>
                  <span style={{ fontSize: 12, color: 'var(--text-secondary)' }}>→</span>
                  <span style={{ fontSize: '0.8rem', color: 'var(--text)' }}>Sign out</span>
                </div>
              </div>
            )}
          </div>
        </div>
      </header>

      {/* Command Palette */}
      {cmd && (
        <div className="modal-backdrop" onClick={() => setCmd(false)}>
          <div style={{ background: 'var(--surface)', border: '1px solid var(--border)', borderRadius: 12, width: 520, maxWidth: '90vw', boxShadow: 'var(--shadow-xl)', animation: 'scaleIn 0.15s ease', overflow: 'hidden' }} onClick={e => e.stopPropagation()}>
            <div className="search-bar" style={{ margin: 0, borderRadius: 0, border: 'none', borderBottom: '1px solid var(--border)', padding: '12px 16px' }}>
              <svg width={14} height={14} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth={2} style={{ color: 'var(--text-muted)', flexShrink: 0 }}><circle cx={11} cy={11} r={8}/><path d="m21 21-4.35-4.35"/></svg>
              <input autoFocus placeholder="Search pages, commands, files..." value={cmdQ} onChange={e => setCmdQ(e.target.value)} style={{ fontSize: '0.9rem', color: 'var(--text)' }}/>
              <kbd style={{ fontSize: '0.65rem' }}>ESC</kbd>
            </div>
            <div style={{ maxHeight: 360, overflowY: 'auto', padding: '6px 0' }}>
              <div style={{ padding: '6px 16px 3px', fontSize: '0.67rem', fontFamily: 'var(--font-mono)', color: 'var(--text-muted)', letterSpacing: '0.06em', textTransform: 'uppercase' }}>Pages</div>
              {pages.filter(p => !cmdQ || p.label.toLowerCase().includes(cmdQ.toLowerCase())).map(p => (
                <div key={p.path} className="ctx-item" style={{ padding: '8px 16px', margin: '1px 4px', borderRadius: 6 }} onClick={() => { setCmd(false); nav(p.path) }}>
                  <span style={{ width: 22, height: 22, display: 'flex', alignItems: 'center', justifyContent: 'center', background: 'var(--bg-3)', border: '1px solid var(--border)', fontSize: 12, borderRadius: 4, flexShrink: 0, color: 'var(--text)' }}>{p.icon}</span>
                  <span style={{ fontSize: '0.8125rem', color: 'var(--text)' }}>{p.label}</span>
                </div>
              ))}
            </div>
            <div style={{ padding: '8px 16px', borderTop: '1px solid var(--border)', display: 'flex', gap: 12, fontSize: '0.7rem', color: 'var(--text-muted)', fontFamily: 'var(--font-mono)' }}>
              <span><kbd>↵</kbd> Select</span><span><kbd>ESC</kbd> Close</span>
            </div>
          </div>
        </div>
      )}
    </>
  )
}
