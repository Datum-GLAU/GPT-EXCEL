import { useState, useEffect, useRef } from 'react'
import { useNavigate } from 'react-router-dom'
import { useDispatch } from 'react-redux'
import { setUser } from '..'

const DEMOS = [
  { email: 'admin@xtronexcel.com', pass: 'admin123', name: 'Admin User', plan: 'pro' as const },
  { email: 'krishna@datum.com', pass: 'datum2026', name: 'Krishna Koushik', plan: 'pro' as const },
  { email: 'demo@xtronexcel.com', pass: 'demo123', name: 'Demo User', plan: 'free' as const },
]

function BoxGrid() {
  const [hoverIdx, setHoverIdx] = useState<number | null>(null)
  const COLS = 20, ROWS = 10

  const getDist = (i: number) => {
    if (hoverIdx === null) return 99
    const hr = Math.floor(hoverIdx / COLS), hc = hoverIdx % COLS
    const ir = Math.floor(i / COLS), ic = i % COLS
    return Math.sqrt((hr - ir) ** 2 + (hc - ic) ** 2)
  }

  return (
    <div style={{ position: 'absolute', inset: 0, display: 'grid', gridTemplateColumns: `repeat(${COLS}, 1fr)`, gridTemplateRows: `repeat(${ROWS}, 1fr)`, pointerEvents: 'all', zIndex: 1 }}
      onMouseLeave={() => setHoverIdx(null)}
    >
      {Array(COLS * ROWS).fill(0).map((_, i) => {
        const dist = getDist(i)
        const intensity = dist < 1 ? 1 : dist < 2.2 ? 0.4 : dist < 3.5 ? 0.15 : 0
        return (
          <div key={i} onMouseEnter={() => setHoverIdx(i)} style={{
            border: '1px solid',
            borderColor: intensity > 0 ? `rgba(255,255,255,${intensity * 0.55})` : 'rgba(255,255,255,0.05)',
            background: intensity > 0 ? `rgba(255,255,255,${intensity * 0.1})` : 'transparent',
            boxShadow: intensity === 1 ? 'inset 0 0 20px rgba(255,255,255,0.12)' : 'none',
            transition: 'all 0.25s cubic-bezier(0.4,0,0.2,1)',
            animation: `gridGlitch ${3 + (i % 5)}s infinite`,
            animationDelay: `${(i % 7) * 0.8}s`,
          }}/>
        )
      })}
    </div>
  )
}

function PasswordStrength({ pass }: { pass: string }) {
  const checks = [
    { label: 'At least 6 chars', ok: pass.length >= 6 },
    { label: 'Has number', ok: /\d/.test(pass) },
    { label: 'Has uppercase', ok: /[A-Z]/.test(pass) },
    { label: 'Has special char', ok: /[^A-Za-z0-9]/.test(pass) },
  ]
  const score = checks.filter(c => c.ok).length
  const colors = ['#ef4444', '#f97316', '#eab308', '#22c55e']
  const labels = ['Weak', 'Fair', 'Good', 'Strong']

  if (!pass) return null

  return (
    <div style={{ marginTop: -8 }}>
      <div style={{ display: 'flex', gap: 4, marginBottom: 6 }}>
        {[0,1,2,3].map(i => (
          <div key={i} style={{ flex: 1, height: 3, borderRadius: 2, background: i < score ? colors[score - 1] : 'rgba(255,255,255,0.1)', transition: 'background 0.3s' }}/>
        ))}
      </div>
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
        <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
          {checks.map(c => (
            <span key={c.label} style={{ fontSize: '0.6rem', color: c.ok ? '#22c55e' : 'rgba(255,255,255,0.25)', display: 'flex', alignItems: 'center', gap: 3 }}>
              <span>{c.ok ? '✓' : '○'}</span>{c.label}
            </span>
          ))}
        </div>
        {score > 0 && <span style={{ fontSize: '0.65rem', color: colors[score - 1], fontWeight: 600 }}>{labels[score - 1]}</span>}
      </div>
    </div>
  )
}

export default function Auth() {
  const navigate = useNavigate()
  const dispatch = useDispatch()
  const [mode, setMode] = useState<'login' | 'register' | 'forgot'>('login')
  const [email, setEmail] = useState('')
  const [password, setPassword] = useState('')
  const [name, setName] = useState('')
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState('')
  const [success, setSuccess] = useState('')
  const [showPass, setShowPass] = useState(false)
  const [mounted, setMounted] = useState(false)
  const [showDemos, setShowDemos] = useState(false)
  const [forgotSent, setForgotSent] = useState(false)
  const emailRef = useRef<HTMLInputElement>(null)

  useEffect(() => { setTimeout(() => setMounted(true), 50) }, [])
  useEffect(() => { setError(''); setSuccess('') }, [mode])
  useEffect(() => { emailRef.current?.focus() }, [mode])

  const fillDemo = (d: typeof DEMOS[0]) => {
    setEmail(d.email); setPassword(d.pass)
    setMode('login'); setShowDemos(false); setError('')
  }

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault()
    setError(''); setSuccess('')
    if (!email) { setError('Email is required'); return }
    if (mode !== 'forgot' && !password) { setError('Password is required'); return }
    if (mode === 'register' && !name) { setError('Full name is required'); return }
    if (mode === 'register' && password.length < 6) { setError('Password must be at least 6 characters'); return }

    setLoading(true)
    await new Promise(r => setTimeout(r, 900))

    if (mode === 'login') {
      const found = DEMOS.find(d => d.email === email && d.pass === password)
      if (found) {
        dispatch(setUser(found))
        setSuccess(`Welcome back, ${found.name}!`)
        setTimeout(() => navigate('/dashboard'), 700)
      } else {
        const emailExists = DEMOS.find(d => d.email === email)
        setError(emailExists ? 'Incorrect password. Check demo credentials below.' : 'No account found with this email.')
      }
    } else if (mode === 'register') {
      dispatch(setUser({ name, email, plan: 'free' }))
      setSuccess(`Account created! Welcome, ${name}!`)
      setTimeout(() => navigate('/dashboard'), 700)
    } else {
      setForgotSent(true)
      setSuccess('Reset link sent! (Demo — no real email sent)')
    }
    setLoading(false)
  }

  return (
    <div style={{ height: '100vh', background: '#030303', color: '#fff', display: 'flex', alignItems: 'center', justifyContent: 'center', position: 'relative', overflow: 'hidden', opacity: mounted ? 1 : 0, transition: 'opacity 0.5s ease' }}>
      <style>{`
        @keyframes gridGlitch {
          0%,100% { border-color: rgba(255,255,255,0.05); }
          50% { border-color: rgba(255,255,255,0.18); }
          51% { border-color: rgba(255,255,255,0.05); }
          52% { border-color: rgba(255,255,255,0.22); }
        }
        @keyframes cardIn { from{opacity:0;transform:translateY(28px) scale(0.97)} to{opacity:1;transform:none} }
        @keyframes scanV { 0%{transform:translateY(-100%)} 100%{transform:translateY(100vh)} }
        @keyframes pulse-dot { 0%,100%{opacity:1} 50%{opacity:0.4} }
        @keyframes shake { 0%,100%{transform:translateX(0)} 20%,60%{transform:translateX(-6px)} 40%,80%{transform:translateX(6px)} }
        .auth-card { animation: cardIn 0.55s cubic-bezier(0.16,1,0.3,1) both; }
        .auth-input {
          width:100%; padding:14px 16px; background:rgba(255,255,255,0.04);
          border:1px solid rgba(255,255,255,0.1); border-radius:12px; color:#fff;
          outline:none; font-size:0.875rem; font-family:inherit; transition:all 0.2s;
          box-sizing:border-box;
        }
        .auth-input:focus { border-color:rgba(255,255,255,0.4); background:rgba(255,255,255,0.07); box-shadow:0 0 0 3px rgba(255,255,255,0.06); }
        .auth-input::placeholder { color:rgba(255,255,255,0.25); }
        .auth-input:hover { border-color:rgba(255,255,255,0.2); }
        .social-btn { flex:1; display:flex; align-items:center; justify-content:center; gap:10px; padding:12px; background:rgba(255,255,255,0.04); border:1px solid rgba(255,255,255,0.1); border-radius:12px; cursor:pointer; color:#fff; font-size:0.82rem; font-weight:500; transition:all 0.2s; font-family:inherit; }
        .social-btn:hover { background:rgba(255,255,255,0.08); border-color:rgba(255,255,255,0.3); transform:translateY(-1px); }
        .demo-row { display:flex; align-items:center; justify-content:space-between; padding:8px 10px; border-radius:9px; cursor:pointer; transition:background 0.15s; }
        .demo-row:hover { background:rgba(255,255,255,0.06); }
        .tab-btn { flex:1; padding:8px; background:none; border:none; color:rgba(255,255,255,0.4); font-size:0.78rem; font-weight:500; cursor:pointer; border-bottom:2px solid transparent; transition:all 0.2s; font-family:inherit; }
        .tab-btn.active { color:#fff; border-bottom-color:rgba(255,255,255,0.7); }
        .shake { animation:shake 0.4s ease; }
      `}</style>

      {/* BG Grid */}
      <div style={{ position: 'absolute', inset: 0, opacity: 0.5 }}><BoxGrid/></div>

      {/* Scan line */}
      <div style={{ position: 'absolute', left: 0, right: 0, height: 1, background: 'linear-gradient(90deg, transparent, rgba(255,255,255,0.15), transparent)', animation: 'scanV 5s linear infinite', pointerEvents: 'none', zIndex: 2 }}/>

      {/* Corner brackets */}
      {[{top:20,left:20},{top:20,right:20},{bottom:20,left:20},{bottom:20,right:20}].map((pos, i) => {
        const isR = 'right' in pos, isB = 'bottom' in pos
        return <div key={i} style={{ position:'absolute', width:18, height:18, borderTop:isB?'none':'1px solid rgba(255,255,255,0.15)', borderBottom:isB?'1px solid rgba(255,255,255,0.15)':'none', borderLeft:isR?'none':'1px solid rgba(255,255,255,0.15)', borderRight:isR?'1px solid rgba(255,255,255,0.15)':'none', ...pos as any, zIndex:3 }}/>
      })}

      {/* Back button */}
      <button onClick={() => navigate('/get-started')} style={{ position:'absolute', top:20, left:20, background:'none', border:'none', color:'rgba(255,255,255,0.4)', fontSize:'0.72rem', cursor:'pointer', fontFamily:'monospace', letterSpacing:'0.06em', zIndex:10, padding:'6px 10px' }}
        onMouseEnter={e => (e.currentTarget.style.color = 'rgba(255,255,255,0.8)')}
        onMouseLeave={e => (e.currentTarget.style.color = 'rgba(255,255,255,0.4)')}
      >← RETURN_HOME</button>

      {/* Status indicator */}
      <div style={{ position:'absolute', top:22, right:24, display:'flex', alignItems:'center', gap:6, zIndex:10 }}>
        <div style={{ width:6, height:6, borderRadius:'50%', background:'#22c55e', animation:'pulse-dot 2s infinite' }}/>
        <span style={{ fontSize:'0.62rem', color:'rgba(255,255,255,0.3)', fontFamily:'monospace', letterSpacing:'0.08em' }}>SYSTEM_ONLINE</span>
      </div>

      {/* Main card */}
      <div className="auth-card" style={{ position:'relative', zIndex:5, width:420, background:'rgba(8,8,8,0.9)', backdropFilter:'blur(32px)', border:'1px solid rgba(255,255,255,0.1)', borderRadius:24, padding:'38px 36px', boxShadow:'0 40px 100px rgba(0,0,0,0.9), inset 0 1px 0 rgba(255,255,255,0.06)' }}>

        {/* Logo */}
        <div style={{ textAlign:'center', marginBottom:28 }}>
          <div style={{ display:'inline-flex', alignItems:'center', gap:5, marginBottom:16 }}>
            <span style={{ fontFamily:'var(--font-display, sans-serif)', fontWeight:900, fontSize:24, letterSpacing:-1, color:'#fff' }}>XtronExcel</span>
          </div>

          {/* Mode tabs */}
          <div style={{ display:'flex', borderBottom:'1px solid rgba(255,255,255,0.08)', marginBottom:4 }}>
            {(['login','register'] as const).map(m => (
              <button key={m} className={`tab-btn${mode === m ? ' active' : ''}`} onClick={() => { setMode(m); setError(''); setSuccess('') }}>
                {m === 'login' ? 'Sign In' : 'Register'}
              </button>
            ))}
          </div>
          <p style={{ fontSize:'0.72rem', color:'rgba(255,255,255,0.3)', marginTop:10 }}>
            {mode === 'login' ? 'Access your AI workspace' : mode === 'register' ? 'Create your free account' : 'Reset your password'}
          </p>
        </div>

        {/* Social buttons */}
        {mode !== 'forgot' && (
          <div style={{ display:'flex', gap:10, marginBottom:20 }}>
            <button className="social-btn" onClick={() => { dispatch(setUser({ name:'Google User', email:'google@xtronexcel.com', plan:'free' })); navigate('/dashboard') }}>
              <svg width={16} height={16} viewBox="0 0 24 24"><path fill="#4285f4" d="M22.56 12.25c0-.78-.07-1.53-.2-2.25H12v4.26h5.92c-.26 1.37-1.04 2.53-2.21 3.31v2.77h3.57c2.08-1.92 3.28-4.74 3.28-8.09z"/><path fill="#34a853" d="M12 23c2.97 0 5.46-.98 7.28-2.66l-3.57-2.77c-.98.66-2.23 1.06-3.71 1.06-2.86 0-5.29-1.93-6.16-4.53H2.18v2.84C3.99 20.53 7.7 23 12 23z"/><path fill="#fbbc05" d="M5.84 14.09c-.22-.66-.35-1.36-.35-2.09s.13-1.43.35-2.09V7.07H2.18C1.43 8.55 1 10.22 1 12s.43 3.45 1.18 4.93l2.85-2.22.81-.62z"/><path fill="#ea4335" d="M12 5.38c1.62 0 3.06.56 4.21 1.64l3.15-3.15C17.45 2.09 14.97 1 12 1 7.7 1 3.99 3.47 2.18 7.07l3.66 2.84c.87-2.6 3.3-4.53 6.16-4.53z"/></svg>
              Google
            </button>
            <button className="social-btn" onClick={() => { dispatch(setUser({ name:'MS User', email:'ms@xtronexcel.com', plan:'free' })); navigate('/dashboard') }}>
              <svg width={16} height={16} viewBox="0 0 23 23"><path fill="#f3f3f3" d="M0 0h11v11H0zM12 0h11v11H12zM0 12h11v11H0zM12 12h11v11H12z"/></svg>
              Microsoft
            </button>
          </div>
        )}

        {/* Divider */}
        {mode !== 'forgot' && (
          <div style={{ display:'flex', alignItems:'center', gap:12, marginBottom:20 }}>
            <div style={{ flex:1, height:1, background:'rgba(255,255,255,0.07)' }}/>
            <span style={{ fontSize:'0.6rem', color:'rgba(255,255,255,0.25)', textTransform:'uppercase', letterSpacing:'0.1em', whiteSpace:'nowrap' }}>or with email</span>
            <div style={{ flex:1, height:1, background:'rgba(255,255,255,0.07)' }}/>
          </div>
        )}

        {/* Forgot password mode */}
        {mode === 'forgot' && (
          <div style={{ marginBottom:20, padding:'14px 16px', background:'rgba(59,130,246,0.08)', border:'1px solid rgba(59,130,246,0.2)', borderRadius:10 }}>
            <p style={{ fontSize:'0.78rem', color:'rgba(255,255,255,0.6)', lineHeight:1.6 }}>
              Enter your email and we'll send a reset link. Demo accounts can reset using their registered email.
            </p>
          </div>
        )}

        {/* Form */}
        <form onSubmit={handleSubmit} style={{ display:'flex', flexDirection:'column', gap:14 }}>
          {mode === 'register' && (
            <div>
              <input className="auth-input" type="text" placeholder="Full name" value={name} onChange={e => setName(e.target.value)} style={{ width:'100%' }}/>
            </div>
          )}

          <div>
            <input ref={emailRef} className="auth-input" type="email" placeholder="Email address" value={email} onChange={e => { setEmail(e.target.value); setError('') }} style={{ width:'100%' }}/>
          </div>

          {mode !== 'forgot' && (
            <div style={{ position:'relative' }}>
              <input className="auth-input" type={showPass ? 'text' : 'password'} placeholder="Password" value={password} onChange={e => { setPassword(e.target.value); setError('') }} style={{ width:'100%', paddingRight:60 }}/>
              <button type="button" onClick={() => setShowPass(p => !p)} style={{ position:'absolute', right:14, top:'50%', transform:'translateY(-50%)', background:'none', border:'none', color:'rgba(255,255,255,0.35)', fontSize:'0.6rem', fontWeight:700, cursor:'pointer', fontFamily:'monospace', letterSpacing:'0.06em' }}>
                {showPass ? 'HIDE' : 'SHOW'}
              </button>
            </div>
          )}

          {mode === 'register' && <PasswordStrength pass={password}/>}

          {mode === 'login' && (
            <div style={{ textAlign:'right', marginTop:-6 }}>
              <button type="button" onClick={() => setMode('forgot')} style={{ background:'none', border:'none', color:'rgba(255,255,255,0.35)', fontSize:'0.72rem', cursor:'pointer', fontFamily:'inherit' }}
                onMouseEnter={e => (e.currentTarget.style.color = 'rgba(255,255,255,0.7)')}
                onMouseLeave={e => (e.currentTarget.style.color = 'rgba(255,255,255,0.35)')}
              >Forgot password?</button>
            </div>
          )}

          {/* Error */}
          {error && (
            <div className="shake" style={{ padding:'10px 14px', background:'rgba(239,68,68,0.1)', border:'1px solid rgba(239,68,68,0.25)', borderRadius:10, fontSize:'0.78rem', color:'#fca5a5', display:'flex', alignItems:'center', gap:8 }}>
              <span style={{ flexShrink:0 }}>⚠</span>{error}
            </div>
          )}

          {/* Success */}
          {success && (
            <div style={{ padding:'10px 14px', background:'rgba(34,197,94,0.1)', border:'1px solid rgba(34,197,94,0.25)', borderRadius:10, fontSize:'0.78rem', color:'#86efac', display:'flex', alignItems:'center', gap:8 }}>
              <span style={{ flexShrink:0 }}>✓</span>{success}
            </div>
          )}

          {/* Submit */}
          <button type="submit" disabled={loading} style={{ width:'100%', padding:'14px', background: loading ? 'rgba(255,255,255,0.08)' : '#fff', color: loading ? 'rgba(255,255,255,0.5)' : '#000', border:'none', borderRadius:12, fontSize:'0.875rem', fontWeight:700, cursor: loading ? 'not-allowed' : 'pointer', transition:'all 0.2s', letterSpacing:'0.04em', fontFamily:'inherit', marginTop:4, display:'flex', alignItems:'center', justifyContent:'center', gap:8 }}>
            {loading ? (
              <><div style={{ width:14, height:14, border:'2px solid rgba(255,255,255,0.3)', borderTopColor:'#fff', borderRadius:'50%', animation:'spin 0.7s linear infinite' }}/> {mode === 'login' ? 'Authenticating...' : mode === 'register' ? 'Creating account...' : 'Sending...'}</>
            ) : (
              mode === 'login' ? 'Sign in →' : mode === 'register' ? 'Create account →' : 'Send reset link →'
            )}
          </button>
        </form>

        {/* Forgot back */}
        {mode === 'forgot' && (
          <div style={{ textAlign:'center', marginTop:16 }}>
            <button onClick={() => { setMode('login'); setForgotSent(false); setSuccess('') }} style={{ background:'none', border:'none', color:'rgba(255,255,255,0.4)', fontSize:'0.78rem', cursor:'pointer', fontFamily:'inherit' }}>
              ← Back to sign in
            </button>
          </div>
        )}

        {/* Divider + Skip */}
        <div style={{ margin:'20px 0 0', display:'flex', flexDirection:'column', gap:12 }}>
          <div style={{ height:1, background:'rgba(255,255,255,0.06)' }}/>
          <button onClick={() => { dispatch(setUser({ name:'Guest', email:'guest@xtronexcel.com', plan:'free' })); navigate('/dashboard') }} style={{ width:'100%', padding:'11px', background:'transparent', border:'1px solid rgba(255,255,255,0.08)', borderRadius:12, color:'rgba(255,255,255,0.45)', fontSize:'0.8rem', cursor:'pointer', fontFamily:'inherit', transition:'all 0.2s' }}
            onMouseEnter={e => { (e.currentTarget as HTMLElement).style.background='rgba(255,255,255,0.04)'; (e.currentTarget as HTMLElement).style.color='rgba(255,255,255,0.7)' }}
            onMouseLeave={e => { (e.currentTarget as HTMLElement).style.background='transparent'; (e.currentTarget as HTMLElement).style.color='rgba(255,255,255,0.45)' }}
          >Skip — explore as guest →</button>
        </div>

        {/* Demo accounts */}
        <div style={{ marginTop:16 }}>
          <button onClick={() => setShowDemos(p => !p)} style={{ width:'100%', background:'none', border:'none', color:'rgba(255,255,255,0.25)', fontSize:'0.65rem', cursor:'pointer', fontFamily:'monospace', letterSpacing:'0.08em', textTransform:'uppercase', display:'flex', alignItems:'center', justifyContent:'center', gap:6 }}>
            <span>{showDemos ? '▲' : '▼'}</span> Demo Accounts
          </button>
          {showDemos && (
            <div style={{ marginTop:10, background:'rgba(255,255,255,0.03)', border:'1px solid rgba(255,255,255,0.08)', borderRadius:12, overflow:'hidden' }}>
              {DEMOS.map((d, i) => (
                <div key={i} className="demo-row" onClick={() => fillDemo(d)} style={{ borderBottom: i < DEMOS.length - 1 ? '1px solid rgba(255,255,255,0.05)' : 'none' }}>
                  <div>
                    <div style={{ fontSize:'0.75rem', color:'rgba(255,255,255,0.8)', fontWeight:500 }}>{d.email}</div>
                    <div style={{ fontSize:'0.65rem', color:'rgba(255,255,255,0.3)', fontFamily:'monospace', marginTop:2 }}>pw: {d.pass} · {d.plan}</div>
                  </div>
                  <span style={{ fontSize:'0.65rem', color:'rgba(255,255,255,0.35)', padding:'4px 10px', border:'1px solid rgba(255,255,255,0.12)', borderRadius:6 }}>Use →</span>
                </div>
              ))}
            </div>
          )}
        </div>
      </div>

      {/* Footer */}
      <div style={{ position:'absolute', bottom:18, fontSize:'0.6rem', color:'rgba(255,255,255,0.12)', fontFamily:'monospace', letterSpacing:'0.14em', zIndex:3 }}>
        XtronExcel v2.0 · Powered by Datum_GLAU · © 2026
      </div>

      <style>{`@keyframes spin { to { transform: rotate(360deg); } }`}</style>
    </div>
  )
}
