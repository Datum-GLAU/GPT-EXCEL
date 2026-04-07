import { useNavigate } from 'react-router-dom';
import { useState, useEffect, useRef, useCallback } from 'react';

// ── DATA ──────────────────────────────────────────────────────────



const FEATURES = [
  { icon: '⊞', title: 'Excel AI Engine', desc: 'Natural language → production-ready Excel files with formulas, pivot tables, charts, and complete financial models from a single prompt.', tags: ['VLOOKUP', 'Pivot', 'Charts', 'Macros'], stat: '200+ formula types', color: 'var(--text)' },
  { icon: '◱', title: 'Document Intelligence', desc: 'Generate reports, proposals, contracts, and CVs. Multi-doc comparison, summarization, grammar correction, and instant PDF export.', tags: ['Reports', 'Contracts', 'Research', 'CVs'], stat: '50+ templates', color: 'var(--text)' },
  { icon: '◈', title: 'Smart File Manager', desc: 'AI-powered file categorization using semantic embeddings. Auto-tag, cluster, deduplicate, and intelligently search your entire workspace.', tags: ['Auto-Tag', 'Search', 'Dedup', 'Cluster'], stat: 'Embedding-based', color: 'var(--text)' },
  { icon: '⌘', title: 'Workflow Automation', desc: 'Build drag-and-drop automation workflows. Schedule recurring reports, automate email drafts, and create powerful AI-driven task pipelines.', tags: ['Schedule', 'Automate', 'Tasks', 'Pipelines'], stat: 'Drag & drop builder', color: 'var(--text)' },
  { icon: '◎', title: 'Voice Commands', desc: 'Hands-free via Web Speech API + Whisper. Voice-to-Excel, voice document creation, and real-time text-to-speech responses.', tags: ['Whisper', 'TTS', 'Offline', 'Live'], stat: 'Real-time transcription', color: 'var(--text)' },
  { icon: '◬', title: 'Data Visualization', desc: 'Interactive charts, financial dashboards, KPI monitors, and predictive forecasting graphs. 20+ chart types with full customization.', tags: ['Charts', 'KPI', 'Forecast', 'Dashboard'], stat: '20+ chart types', color: 'var(--text)' },
  { icon: '◻', title: 'PowerPoint Generator', desc: 'Auto-generate pitch decks with embedded charts and speaker notes. Custom themes, AI-written slide content.', tags: ['Slides', 'Charts', 'Notes', 'Themes'], stat: 'AI-powered decks', color: 'var(--text)' },
  { icon: '🔌', title: 'Plugin Marketplace', desc: 'Extend with 40+ industry-specific plugins. Automation packs, custom integrations, full developer SDK.', tags: ['Plugins', 'Extensions', 'SDK', 'Integrations'], stat: '40+ plugins', color: 'var(--text)' },
  { icon: '☁', title: 'Hybrid Cloud', desc: 'Firebase auth with cloud sync and local Python processing. Works offline for core features. Zero data lock-in.', tags: ['Firebase', 'Offline', 'Sync', 'Python'], stat: 'Hybrid architecture', color: 'var(--text)' },
];

const FAQS = [
  { q: 'Does it work offline?', a: 'Yes. XtronExcel uses a hybrid architecture. Core Excel generation, file management, and basic voice commands work fully offline using local Python. AI features require internet for LLM calls.' },
  { q: 'What AI models does it use?', a: 'Primary: OpenAI GPT-4o for generation. Azure OpenAI for enterprise. Whisper for voice transcription. Configure your own API keys in Settings for full model control.' },
  { q: 'What file formats are supported?', a: 'Excel (.xlsx, .xls, .xlsm), CSV, PDF, Word (.docx), PowerPoint (.pptx), JSON, and plain text. Export to any format from generated content in one click.' },
  { q: 'Is there a free tier?', a: '50 AI generations/month, basic Excel features, and file manager. Pro ($12/mo) unlocks unlimited generations, all voice features, advanced automation, and priority support.' },
  { q: 'How is my data protected?', a: 'Files are processed locally using Python. Only prompts are sent to AI APIs — never raw file contents. Firebase handles auth securely. No vendor lock-in.' },
  { q: 'Can I use my own API key?', a: 'Yes. Settings → API Keys. Paste your OpenAI, Azure, or Anthropic key for direct billing control and access to latest models.' },
];

const TECH = ['Electron.js','React 18','TypeScript','Python 3.11','FastAPI','OpenAI GPT-4o','Whisper API','OpenPyXL','Pandas','Firebase','Redux Toolkit','Monaco Editor','Recharts','Socket.io','SQLite','Stripe','Vite','Electron Builder'];

// ── HOOKS ─────────────────────────────────────────────────────────

function useMouse() {
  const [pos, setPos] = useState({ x: -999, y: -999 });
  useEffect(() => {
    const fn = (e: MouseEvent) => setPos({ x: e.clientX, y: e.clientY });
    window.addEventListener('mousemove', fn, { passive: true });
    return () => window.removeEventListener('mousemove', fn);
  }, []);
  return pos;
}

function useInView(ref: React.RefObject<HTMLElement>, threshold = 0.15) {
  const [vis, setVis] = useState(false);
  useEffect(() => {
    if (!ref.current) return;
    const obs = new IntersectionObserver(([e]) => { if (e.isIntersecting) setVis(true); }, { threshold });
    obs.observe(ref.current);
    return () => obs.disconnect();
  }, []);
  return vis;
}

function useCounter(target: number, dur = 1400, active = false) {
  const [v, setV] = useState(0);
  useEffect(() => {
    if (!active) return;
    let t0: number;
    const step = (ts: number) => {
      if (!t0) t0 = ts;
      const p = Math.min((ts - t0) / dur, 1);
      setV(Math.floor((1 - Math.pow(1 - p, 3)) * target));
      if (p < 1) requestAnimationFrame(step);
    };
    requestAnimationFrame(step);
  }, [target, dur, active]);
  return v;
}

// ── SUBCOMPONENTS ─────────────────────────────────────────────────

function BoxGrid() {
  const [hoverIdx, setHoverIdx] = useState<number | null>(null);
  const COLS = 20, ROWS = 9;

  // Helper to calculate distance for the "glow around" effect
  const getDist = (i: number) => {
    if (hoverIdx === null) return 99;
    const hr = Math.floor(hoverIdx / COLS), hc = hoverIdx % COLS;
    const ir = Math.floor(i / COLS), ic = i % COLS;
    return Math.sqrt((hr - ir) ** 2 + (hc - ic) ** 2);
  };

  return (
    <div 
      style={{ 
        position: 'absolute', inset: 0, 
        display: 'grid', 
        gridTemplateColumns: `repeat(${COLS}, 1fr)`, 
        gridTemplateRows: `repeat(${ROWS}, 1fr)`, 
        pointerEvents: 'all', zIndex: 1 
      }}
      onMouseLeave={() => setHoverIdx(null)}
    >
      {Array(COLS * ROWS).fill(0).map((_, i) => {
        const dist = getDist(i);
        // The center box is 100% bright, neighbors are 40%, others 0%
        const intensity = dist < 1 ? 1 : dist < 2.5 ? 0.4 : 0;
        
        return (
          <div 
            key={i} 
            onMouseEnter={() => setHoverIdx(i)}
            className="glitch-box"
            style={{
              // Individual random flicker for each box
              animation: `gridGlitch ${2 + Math.random() * 4}s infinite`,
              animationDelay: `${Math.random() * 5}s`,
              
              // Hover Glow Logic
              background: intensity > 0 
                ? `rgba(255,255,255, ${intensity * 0.12})` 
                : 'transparent',
              borderColor: intensity > 0 
                ? `rgba(255,255,255, ${intensity * 0.4})` 
                : undefined,
              boxShadow: intensity === 1 
                ? 'inset 0 0 12px rgba(255,255,255,0.1)' 
                : 'none',
            }}
          />
        );
      })}
    </div>
  );
}
function GlowOrb({ x, y }: { x: number; y: number }) {
  return (
    <div style={{ position: 'fixed', left: x, top: y, width: 480, height: 480, borderRadius: '50%', background: 'radial-gradient(circle, rgba(255,255,255,0.06) 0%, transparent 62%)', transform: 'translate(-50%,-50%)', pointerEvents: 'none', zIndex: 0, transition: 'left 0.1s linear, top 0.1s linear' }}/>
  );
}

function TerminalWidget() {
  const ref = useRef<HTMLDivElement>(null);
  const inView = useInView(ref as any);
  const [lines, setLines] = useState<typeof TERMINAL_LINES>([]);
  const [blink, setBlink] = useState(true);
  const [loop, setLoop] = useState(0);
  const TERMINAL_LINES = [
    { txt: '$ xtronexcel generate --prompt "Q4 budget tracker"', col: 'rgba(255,255,255,0.35)', d: 0 },
    { txt: '⠿ Analyzing prompt with GPT-4o...', col: 'rgba(255,255,255,0.72)', d: 600 },
    { txt: '⠿ Generating Python + openpyxl code...', col: 'rgba(255,255,255,0.72)', d: 1300 },
    { txt: '⠿ Executing locally on your machine...', col: 'rgba(255,255,255,0.72)', d: 2000 },
    { txt: '⠿ Building 7 sheets, 24 formulas, 3 charts...', col: 'rgba(255,255,255,0.72)', d: 2800 },
    { txt: '✓ budget_q4_2026.xlsx — 1.2 MB — Done', col: '#ffffff', d: 3500 },
    { txt: '✓ Generated in 4.8s  |  Ready to preview', col: '#ffffff', d: 3900 },
  ];

  useEffect(() => { const t = setInterval(() => setBlink(b => !b), 520); return () => clearInterval(t); }, []);

  useEffect(() => {
    if (!inView) return;
    setLines([]);
    TERMINAL_LINES.forEach((l, i) => {
      setTimeout(() => setLines(p => [...p, l]), l.d + loop * 5200);
    });
    const reset = setTimeout(() => { setLines([]); setLoop(n => n + 1); }, 5600);
    return () => clearTimeout(reset);
  }, [inView, loop]);

  return (
    <div ref={ref} style={{ background: '#080808', border: '1px solid rgba(255,255,255,0.08)', borderRadius: 12, overflow: 'hidden', fontFamily: 'var(--font-mono)', fontSize: '0.76rem', boxShadow: '0 20px 60px rgba(0,0,0,0.5)' }}>
      <div style={{ display: 'flex', alignItems: 'center', gap: 6, padding: '9px 14px', borderBottom: '1px solid rgba(255,255,255,0.06)', background: 'rgba(255,255,255,0.025)' }}>
        <div style={{ width: 10, height: 10, borderRadius: '50%', background: '#2a2a2a' }}/>
        <div style={{ width: 10, height: 10, borderRadius: '50%', background: '#2a2a2a' }}/>
        <div style={{ width: 10, height: 10, borderRadius: '50%', background: '#2a2a2a' }}/>
        <span style={{ marginLeft: 10, color: 'rgba(255,255,255,0.22)', fontSize: '0.68rem', letterSpacing: '0.06em' }}>xtronexcel — bash</span>
        <div style={{ marginLeft: 'auto', display: 'flex', alignItems: 'center', gap: 5 }}>
          <div style={{ width: 6, height: 6, borderRadius: '50%', background: '#ffffff', animation: 'pulse-soft 2s infinite' }}/>
          <span style={{ fontSize: '0.65rem', color: 'rgba(255,255,255,0.85)', letterSpacing: '0.06em' }}>LIVE</span>
        </div>
      </div>
      <div style={{ padding: '16px 18px', minHeight: 172, display: 'flex', flexDirection: 'column', gap: 7 }}>
        {lines.map((l, i) => (
          <div key={`${loop}-${i}`} style={{ color: l.col, animation: 'tLine 0.2s ease', lineHeight: 1.6 }}>
            {l.txt}
          </div>
        ))}
        {lines.length < TERMINAL_LINES.length && (
          <span style={{ color: 'rgba(255,255,255,0.3)', borderRight: blink ? '1.5px solid rgba(255,255,255,0.35)' : '1.5px solid transparent', paddingRight: 1 }}>&nbsp;</span>
        )}
      </div>
    </div>
  );
}


function FeatCard({ feat, idx, active, onClick }: { feat: typeof FEATURES[0]; idx: number; active: boolean; onClick: () => void }) {
  const ref = useRef<HTMLDivElement>(null);
  const inView = useInView(ref as any, 0.08);
  const [hov, setHov] = useState(false);

  return (
    <div ref={ref} onClick={onClick} onMouseEnter={() => setHov(true)} onMouseLeave={() => setHov(false)}
      style={{ background: active || hov ? 'var(--surface-2)' : 'var(--surface)', border: `1px solid ${active ? feat.color : hov ? 'var(--accent)' : 'var(--border)'}`, borderRadius: 12, padding: '20px', cursor: 'pointer', transition: 'all 0.22s ease', transform: inView ? (hov || active ? 'translateY(-4px)' : 'translateY(0)') : 'translateY(22px)', opacity: inView ? 1 : 0, transitionDelay: `${idx * 0.045}s`, boxShadow: active || hov ? `0 10px 28px rgba(0,0,0,0.2), 0 0 0 1px ${feat.color}18` : 'none', position: 'relative', overflow: 'hidden' }}
    >
      {(active || hov) && <div style={{ position: 'absolute', inset: 0, background: `radial-gradient(circle at 15% 15%, ${feat.color}08 0%, transparent 55%)`, pointerEvents: 'none' }}/>}
      <div style={{ width: 38, height: 38, fontSize: 18, display: 'flex', alignItems: 'center', justifyContent: 'center', background: `${feat.color}15`, border: `1px solid ${feat.color}28`, borderRadius: 9, marginBottom: 12, color: feat.color, transition: 'transform 0.2s', transform: hov ? 'scale(1.08)' : 'scale(1)' }}>{feat.icon}</div>
      <h3 style={{ marginBottom: 7, fontSize: '0.9rem', color: 'var(--text)' }}>{feat.title}</h3>
      <p style={{ fontSize: '0.78rem', lineHeight: 1.6, marginBottom: 12, color: 'var(--text-sec)' }}>{feat.desc}</p>
      <div style={{ display: 'flex', flexWrap: 'wrap', gap: 4, marginBottom: 10 }}>{feat.tags.map(t => <span key={t} className="tag">{t}</span>)}</div>
      <div style={{ fontSize: '0.7rem', fontFamily: 'var(--font-mono)', color: feat.color, letterSpacing: '0.05em' }}>→ {feat.stat}</div>
    </div>
  );
}

function TechBadge({ label, idx }: { label: string; idx: number }) {
  const [hov, setHov] = useState(false);
  const cols = ['var(--text)','var(--text-sec)','var(--text-muted)','var(--text)','var(--text-sec)','var(--text-muted)','var(--text)'];
  const c = cols[idx % cols.length];
  return (
    <span onMouseEnter={() => setHov(true)} onMouseLeave={() => setHov(false)}
      style={{ padding: '5px 12px', fontSize: '0.75rem', fontFamily: 'var(--font-mono)', letterSpacing: '0.03em', borderRadius: 5, cursor: 'default', transition: 'all 0.18s ease', border: `1px solid ${hov ? c : 'var(--border)'}`, color: hov ? c : 'var(--text-sec)', background: hov ? `${c}10` : 'var(--surface-2)', transform: hov ? 'translateY(-2px)' : 'none', boxShadow: hov ? `0 4px 14px ${c}22` : 'none' }}
    >{label}</span>
  );
}

function FaqRow({ faq, open, onToggle, idx }: { faq: typeof FAQS[0]; open: boolean; onToggle: () => void; idx: number }) {
  const ref = useRef<HTMLDivElement>(null);
  const inView = useInView(ref as any, 0.1);
  return (
    <div ref={ref} className="faq-item" style={{ opacity: inView ? 1 : 0, transform: inView ? 'none' : 'translateX(-14px)', transition: `all 0.4s ease ${idx * 0.06}s` }}>
      <div className="faq-q" onClick={onToggle}>
        <span style={{ color: 'var(--text)' }}>{faq.q}</span>
        <span style={{ color: open ? 'var(--text)' : 'var(--text-muted)', transition: 'transform 0.25s ease, color 0.2s', transform: open ? 'rotate(45deg)' : 'none', display: 'inline-block', fontSize: 20 }}>+</span>
      </div>
      <div className={`faq-a${open ? ' open' : ''}`}>{faq.a}</div>
    </div>
  );
}

// ── MAIN ──────────────────────────────────────────────────────────

export default function GetStarted() {
  const nav = useNavigate();
  const mouse = useMouse();
  const [vis, setVis] = useState(false);
  const [openFaq, setOpenFaq] = useState<number | null>(null);
  const [activeFeat, setActiveFeat] = useState(0);
  const [dark, setDark] = useState(() => !document.documentElement.classList.contains('light'));
  const [scrolled, setScrolled] = useState(false);

  useEffect(() => { setTimeout(() => setVis(true), 60); }, []);
  useEffect(() => { const t = setInterval(() => setActiveFeat(p => (p + 1) % FEATURES.length), 3800); return () => clearInterval(t); }, []);
  useEffect(() => { const fn = () => setScrolled(window.scrollY > 24); window.addEventListener('scroll', fn, { passive: true }); return () => window.removeEventListener('scroll', fn); }, []);

  const toggleTheme = () => {
    const nd = !dark; setDark(nd);
    if (nd) { document.documentElement.classList.add('dark'); document.documentElement.classList.remove('light'); }
    else { document.documentElement.classList.add('light'); document.documentElement.classList.remove('dark'); }
    localStorage.setItem('theme', nd ? 'dark' : 'light');
  };

  return (
    <div style={{ height: '100vh', overflowY: 'auto', background: 'var(--bg)', color: 'var(--text)', fontFamily: 'var(--font-body)', opacity: vis ? 1 : 0, transition: 'opacity 0.4s ease' }}>
      <style>{`
       
      @keyframes gridGlitch {
  0%, 100% { border-color: rgba(255,255,255,0.03); }
  50% { border-color: rgba(255,255,255,0.15); }
  51% { border-color: rgba(255,255,255,0.03); }
  52% { border-color: rgba(255,255,255,0.2); }
}

.glitch-box {
  border: 1px solid rgba(255,255,255,0.03);
  transition: background 0.4s ease, border-color 0.2s ease, box-shadow 0.4s ease;
}
        @keyframes tLine { from{opacity:0;transform:translateX(-8px)} to{opacity:1;transform:none} }
        @keyframes heroOrb { 0%,100%{opacity:0.5;transform:translate(-50%,-50%) scale(1)} 50%{opacity:1;transform:translate(-50%,-50%) scale(1.12)} }
        @keyframes scanH { 0%{transform:translateY(-100%)} 100%{transform:translateY(700px)} }
        @keyframes slideW { from{opacity:0;transform:translateY(24px)} to{opacity:1;transform:none} }
        @keyframes popIn { from{opacity:0;transform:scale(0.92) translateY(8px)} to{opacity:1;transform:none} }
        @keyframes borderPulse { 0%,100%{box-shadow:0 0 0 0 transparent} 50%{box-shadow:0 0 0 3px rgba(255,255,255,0.15)} }
        @keyframes shimMove { 0%{transform:translateX(-100%)} 100%{transform:translateX(300%)} }
        @keyframes borderGlitch {0%, 100% { border-color: rgba(255,255,255,0.05); }33% { border-color: rgba(255,255,255,0.12); }66% { border-color: rgba(255,255,255,0.08); }}
        @keyframes boxFill {from { background: transparent; transform: scale(0.95); }to { background: rgba(255,255,255,0.08); transform: scale(1); }}
        .w1{animation:slideW 0.5s ease 0.1s both}
        .w2{animation:slideW 0.5s ease 0.2s both}
        .w3{animation:slideW 0.5s ease 0.3s both}
        .w4{animation:slideW 0.5s ease 0.4s both}
        .w5{animation:slideW 0.5s ease 0.5s both}
        .hlift{transition:all 0.2s ease}
        .hlift:hover{transform:translateY(-3px);box-shadow:var(--shadow-lg), 0 0 0 1px var(--accent)}
        .gs-nav{backdrop-filter:blur(14px);-webkit-backdrop-filter:blur(14px)}
      `}</style>

      

      {/* Navbar */}
      <nav className="gs-nav" style={{ position: 'sticky', top: 0, zIndex: 50, display: 'flex', alignItems: 'center', padding: '0 28px', height: 52, background: scrolled ? 'rgba(10,10,10,0.9)' : 'var(--bg)', borderBottom: `1px solid ${scrolled ? 'var(--border)' : 'transparent'}`, transition: 'all 0.3s ease' }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 5, cursor: 'pointer' }} onClick={() => window.scrollTo({ top: 0, behavior: 'smooth' })}>
          <span style={{ fontFamily: 'var(--font-display)', fontSize: 16, fontWeight: 800, letterSpacing: -0.5, color: 'var(--text)' }}>XtronExcel</span>
        </div>
        <div style={{ display: 'flex', gap: 2, marginLeft: 28 }}>
          {['Features', 'Pricing', 'Docs', 'About'].map(l => <span key={l} className="nav-link">{l}</span>)}
        </div>
        <div style={{ marginLeft: 'auto', display: 'flex', gap: 6, alignItems: 'center' }}>
          <span style={{ fontSize: '0.7rem', color: 'var(--text-muted)', fontFamily: 'var(--font-mono)' }}>Powered by Datum_GLAU</span>
          <div style={{ width: 1, height: 14, background: 'var(--border)' }}/>
          <button className="btn btn-icon-sm btn-ghost" onClick={toggleTheme} data-tip={dark ? 'Light mode' : 'Dark mode'} style={{ color: 'var(--text-sec)' }}>
            {dark
              ? <svg width={14} height={14} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth={2}><circle cx={12} cy={12} r={5}/><line x1={12} y1={1} x2={12} y2={3}/><line x1={12} y1={21} x2={12} y2={23}/><line x1={1} y1={12} x2={3} y2={12}/><line x1={21} y1={12} x2={23} y2={12}/><line x1={4.22} y1={4.22} x2={5.64} y2={5.64}/><line x1={18.36} y1={18.36} x2={19.78} y2={19.78}/><line x1={4.22} y1={19.78} x2={5.64} y2={18.36}/><line x1={18.36} y1={5.64} x2={19.78} y2={4.22}/></svg>
              : <svg width={14} height={14} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth={2}><path d="M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z"/></svg>
            }
          </button>
          <button className="btn btn-ghost btn-sm" onClick={() => nav('/auth')} style={{ color: 'var(--text-sec)' }}>Sign in</button>
          <button className="btn btn-primary btn-sm" onClick={() => nav('/auth')}>Get Started →</button>
        </div>
      </nav>

      {/* Hero Section */}
      <section style={{ padding: '72px 0 56px', display: 'flex', flexDirection: 'column', alignItems: 'center', textAlign: 'center', borderBottom: '1px solid var(--border)', position: 'relative', overflow: 'hidden', minHeight: 580 }}>
        <BoxGrid/>
        <div style={{ position: 'absolute', width: 640, height: 640, borderRadius: '50%', background: 'radial-gradient(circle, rgba(255,255,255,0.04) 0%, transparent 65%)', top: '50%', left: '50%', transform: 'translate(-50%,-50%)', animation: 'heroOrb 5s ease infinite', pointerEvents: 'none', zIndex: 1 }}/>
        <div style={{ position: 'absolute', left: 0, right: 0, height: 1.5, background: 'linear-gradient(90deg, transparent 0%, rgba(255,255,255,0.3) 50%, transparent 100%)', animation: 'scanH 6s linear infinite', pointerEvents: 'none', zIndex: 2 }}/>

        <div style={{ position: 'relative', zIndex: 3, padding: '0 24px', maxWidth: 800 }}>
          <div style={{ marginBottom: 20, animation: 'popIn 0.4s ease 0.05s both' }}>
            <span className="badge badge-outline" style={{ fontSize: '0.7rem', letterSpacing: '0.06em', padding: '4px 14px', animation: 'borderPulse 3s ease infinite' }}>
              ● NEW — XtronExcel v2.0 — Now Live
            </span>
          </div>

          <h1 style={{ fontFamily: 'var(--font-display)', fontSize: 'clamp(2.5rem, 6vw, 4.2rem)', fontWeight: 800, letterSpacing: -2.5, lineHeight: 1.07, marginBottom: 22 }}>
            <span className="w1" style={{ display: 'inline-block' }}>AI-Powered</span>{' '}
            <span className="w2" style={{ display: 'inline-block' }}>Spreadsheet</span>{' '}
            <br/>
            <span className="w3" style={{ display: 'inline-block', color: 'var(--text)' }}>Intelligence</span>{' '}
            <span className="w4" style={{ display: 'inline-block' }}>for</span>{' '}
            <span className="w5" style={{ display: 'inline-block' }}>Everyone</span>
          </h1>

          <p style={{ fontSize: '1.0rem', color: 'var(--text-sec)', lineHeight: 1.78, maxWidth: 540, margin: '0 auto 34px', animation: 'popIn 0.5s ease 0.5s both', opacity: 0 }}>
            XtronExcel combines LLM intelligence with Python automation to generate Excel files, documents, and data visualizations from natural language.
          </p>

          <div style={{ display: 'flex', gap: 10, justifyContent: 'center', flexWrap: 'wrap', animation: 'popIn 0.5s ease 0.65s both', opacity: 0, marginBottom: 22 }}>
            <button className="btn btn-primary btn-xl hlift" onClick={() => nav('/auth')} style={{ boxShadow: '0 0 24px rgba(255,255,255,0.12)' }}>Start for free →</button>
            <button className="btn btn-outline btn-xl hlift" onClick={() => nav('/dashboard')}>View demo</button>
          </div>

          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 22, fontSize: '0.75rem', color: 'var(--text-muted)', animation: 'popIn 0.5s ease 0.8s both', opacity: 0 }}>
            {['No credit card', '50 free generations/mo', 'Offline capable', 'Open source'].map(t => (
              <span key={t} style={{ display: 'flex', alignItems: 'center', gap: 4 }}><span style={{ color: 'var(--text)' }}>✓</span>{t}</span>
            ))}
          </div>
        </div>
      </section>

   

      {/* Terminal Showcase */}
      <section style={{ padding: '64px 28px', borderBottom: '1px solid var(--border)', background: 'var(--bg-2)' }}>
        <div style={{ maxWidth: 940, margin: '0 auto', display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 40, alignItems: 'center' }}>
          <div>
            <div className="section-label">Under the hood</div>
            <h2 style={{ marginBottom: 14, color: 'var(--text)' }}>Python-powered. <span style={{ color: 'var(--text)' }}>AI-driven.</span></h2>
            <p style={{ marginBottom: 22, lineHeight: 1.75, fontSize: '0.875rem' }}>GPT-4o generates production-quality Python using openpyxl and pandas. Our engine validates, tests, and executes it locally no server-side processing, maximum privacy.</p>
            <div style={{ display: 'flex', flexDirection: 'column', gap: 9 }}>
              {[
                { icon: '◈', txt: 'Conditional formatting with color scales', c: 'var(--text)' },
                { icon: '◬', txt: 'Embedded bar, line, pie, scatter charts', c: 'var(--text)' },
                { icon: '⊞', txt: 'Pivot tables with multi-level grouping', c: 'var(--text)' },
                { icon: '⌘', txt: 'Named ranges, data validation, dropdowns', c: 'var(--text)' },
                { icon: '◱', txt: 'Cross-sheet formulas with full references', c: 'var(--text)' },
              ].map(f => (
                <div key={f.txt} style={{ display: 'flex', alignItems: 'center', gap: 10, padding: '8px 12px', background: 'var(--surface)', border: '1px solid var(--border)', borderRadius: 8, transition: 'all 0.18s ease', cursor: 'default' }}
                  onMouseEnter={e => { const el = e.currentTarget as HTMLElement; el.style.borderColor = f.c; el.style.background = `${f.c}08`; el.style.transform = 'translateX(4px)'; el.style.boxShadow = `0 0 0 1px ${f.c}40`; }}
                  onMouseLeave={e => { const el = e.currentTarget as HTMLElement; el.style.borderColor = 'var(--border)'; el.style.background = 'var(--surface)'; el.style.transform = 'none'; el.style.boxShadow = 'none'; }}
                >
                  <span style={{ color: f.c, fontSize: 14, flexShrink: 0 }}>{f.icon}</span>
                  <span style={{ fontSize: '0.8rem', color: 'var(--text-sec)' }}>{f.txt}</span>
                </div>
              ))}
            </div>
          </div>
          <TerminalWidget/>
        </div>
      </section>

      {/* Features */}
      <section style={{ padding: '64px 28px', borderBottom: '1px solid var(--border)' }}>
        <div style={{ maxWidth: 1090, margin: '0 auto' }}>
          <div className="section-label">Features</div>
          <div style={{ display: 'flex', alignItems: 'flex-end', justifyContent: 'space-between', marginBottom: 36 }}>
            <div>
              <h2 style={{ marginBottom: 8, color: 'var(--text)' }}>Everything you need. Nothing you don't.</h2>
              <p style={{ maxWidth: 460, fontSize: '0.835rem' }}>A complete AI productivity suite built for analysts, developers, and business teams.</p>
            </div>
            <div style={{ display: 'flex', gap: 3 }}>
              {FEATURES.map((_, i) => (
                <div key={i} onClick={() => setActiveFeat(i)} style={{ width: i === activeFeat ? 22 : 7, height: 7, borderRadius: 99, background: i === activeFeat ? 'var(--text)' : 'var(--border-2)', transition: 'all 0.3s ease', cursor: 'pointer' }}/>
              ))}
            </div>
          </div>
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: 12 }}>
            {FEATURES.map((f, i) => <FeatCard key={i} feat={f} idx={i} active={activeFeat === i} onClick={() => setActiveFeat(i)}/>)}
          </div>
        </div>
      </section>

      {/* How It Works */}
      <section style={{ padding: '64px 28px', borderBottom: '1px solid var(--border)', background: 'var(--bg-2)' }}>
        <div style={{ maxWidth: 860, margin: '0 auto' }}>
          <div className="section-label">How it works</div>
          <h2 style={{ marginBottom: 48, color: 'var(--text)' }}>Three steps to your spreadsheet</h2>
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', position: 'relative' }}>
            <div style={{ position: 'absolute', top: 22, left: '16.5%', right: '16.5%', height: 1, background: 'linear-gradient(90deg, transparent, var(--border-hi), transparent)' }}/>
            {[
              { step: '01', title: 'Describe your need', desc: 'Write in plain English. "Monthly budget with 6 expense categories, running totals, overspend alerts, and a comparison bar chart." Specificity gives better results.', icon: '✍️', color: 'var(--text)' },
              { step: '02', title: 'AI generates & runs', desc: 'GPT-4o writes Python using openpyxl/pandas. Our engine validates, tests, and executes it locally — no data leaves your machine, ever.', icon: '⚙️', color: 'var(--text)' },
              { step: '03', title: 'Preview & download', desc: 'Live spreadsheet preview before downloading. Export as .xlsx, .csv, .pdf, or .docx. Share directly or embed in other tools.', icon: '📥', color: 'var(--text)' },
            ].map((s, i) => {
              const r = useRef<HTMLDivElement>(null);
              const iv = useInView(r as any, 0.2);
              return (
                <div key={i} ref={r} style={{ padding: '0 24px', textAlign: 'center', position: 'relative', zIndex: 1, opacity: iv ? 1 : 0, transform: iv ? 'none' : 'translateY(22px)', transition: `all 0.5s ease ${i * 0.14}s` }}>
                  <div style={{ width: 44, height: 44, border: `2px solid ${s.color}35`, background: `${s.color}12`, display: 'flex', alignItems: 'center', justifyContent: 'center', margin: '0 auto 14px', fontFamily: 'var(--font-mono)', fontSize: '0.7rem', color: s.color, letterSpacing: '0.05em', fontWeight: 700, borderRadius: 10 }}>{s.step}</div>
                  <div style={{ fontSize: 30, marginBottom: 12 }}>{s.icon}</div>
                  <h3 style={{ marginBottom: 10, fontSize: '0.9rem', color: 'var(--text)' }}>{s.title}</h3>
                  <p style={{ fontSize: '0.78rem', lineHeight: 1.65, color: 'var(--text-sec)' }}>{s.desc}</p>
                </div>
              );
            })}
          </div>
        </div>
      </section>

      {/* Pricing */}
      <section style={{ padding: '64px 28px', borderBottom: '1px solid var(--border)' }}>
        <div style={{ maxWidth: 740, margin: '0 auto', textAlign: 'center' }}>
          <div className="section-label" style={{ textAlign: 'center' }}>Pricing</div>
          <h2 style={{ marginBottom: 8, color: 'var(--text)' }}>Simple, transparent pricing</h2>
          <p style={{ marginBottom: 44 }}>Start free. Upgrade when you need more power.</p>
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 18 }}>
            {[
              { name: 'Free', price: '$0', desc: 'Perfect for individuals getting started', color: 'var(--text)', features: ['50 AI generations/month', 'Basic Excel generation', '1 GB file storage', '3 workflow automations', 'Community support', 'Core voice commands'], btn: 'Get started free', cls: 'btn-outline' },
              { name: 'Pro', price: '$12', desc: 'Everything for power users & teams', color: 'var(--text)', features: ['Unlimited AI generations', 'All Excel + Document features', '50 GB file storage', 'Unlimited workflows', 'Whisper voice commands', 'Priority support (24h)', 'Plugin marketplace', 'Custom AI model selection', 'Team collaboration'], btn: 'Start Pro trial', cls: 'btn-primary', popular: true },
            ].map((p, i) => {
              const r = useRef<HTMLDivElement>(null);
              const iv = useInView(r as any, 0.2);
              return (
                <div ref={r} key={i} style={{ background: 'var(--surface)', border: `1px solid ${p.popular ? 'var(--border-hi)' : 'var(--border)'}`, borderRadius: 14, padding: '26px', position: 'relative', textAlign: 'left', opacity: iv ? 1 : 0, transform: iv ? 'none' : 'translateY(22px)', transition: `all 0.5s ease ${i * 0.1}s`, boxShadow: p.popular ? '0 0 0 1px var(--border-hi), var(--shadow-lg)' : 'none' }}>
                  {p.popular && <div style={{ position: 'absolute', top: -11, left: '50%', transform: 'translateX(-50%)' }}><span className="badge badge-outline" style={{ fontSize: '0.62rem', padding: '3px 10px', letterSpacing: '0.06em', color: 'var(--text)', borderColor: 'var(--border-hi)' }}>MOST POPULAR</span></div>}
                  <div style={{ marginBottom: 18 }}>
                    <div style={{ fontFamily: 'var(--font-display)', fontSize: '1rem', fontWeight: 700, color: 'var(--text)', marginBottom: 4 }}>{p.name}</div>
                    <div style={{ display: 'flex', alignItems: 'baseline', gap: 3, marginBottom: 5 }}>
                      <span style={{ fontFamily: 'var(--font-display)', fontSize: '2.4rem', fontWeight: 800, color: p.color }}>{p.price}</span>
                      <span style={{ fontSize: '0.75rem', color: 'var(--text-muted)' }}>/month</span>
                    </div>
                    <div style={{ fontSize: '0.78rem', color: 'var(--text-muted)' }}>{p.desc}</div>
                  </div>
                  <div style={{ display: 'flex', flexDirection: 'column', gap: 9, marginBottom: 22 }}>
                    {p.features.map((f, fi) => (
                      <div key={fi} style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: '0.8rem', color: 'var(--text-sec)' }}>
                        <span style={{ color: p.color, flexShrink: 0 }}>✓</span>{f}
                      </div>
                    ))}
                  </div>
                  <button className={`btn ${p.cls} w-full hlift`} style={{ justifyContent: 'center' }} onClick={() => nav('/auth')}>{p.btn}</button>
                </div>
              );
            })}
          </div>
        </div>
      </section>

      {/* Testimonials */}
      <section style={{ padding: '64px 28px', borderBottom: '1px solid var(--border)', background: 'var(--bg-2)' }}>
        <div style={{ maxWidth: 980, margin: '0 auto' }}>
          <div className="section-label">What people say</div>
          <h2 style={{ marginBottom: 40, color: 'var(--text)' }}>Loved by data teams worldwide</h2>
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: 16 }}>
            {[
              { quote: 'XtronExcel saved our finance team 12+ hours every week. What used to take half a day now takes under 2 minutes.', name: 'Sarah K.', role: 'CFO, SaaS Startup', rating: 5, color: 'var(--text)' },
              { quote: 'The VLOOKUP and pivot generation is insane. I just describe what I need and it works perfectly, every time.', name: 'Marcus T.', role: 'Data Analyst, Fortune 500', rating: 5, color: 'var(--text)' },
              { quote: 'As a non-technical founder, this is magic. I can now create financial models that impress investors without knowing Python.', name: 'Priya M.', role: 'Founder, EdTech', rating: 5, color: 'var(--text)' },
            ].map((t, i) => {
              const r = useRef<HTMLDivElement>(null);
              const iv = useInView(r as any, 0.2);
              return (
                <div ref={r} key={i} style={{ background: 'var(--surface)', border: '1px solid var(--border)', borderRadius: 12, padding: '22px', opacity: iv ? 1 : 0, transform: iv ? 'none' : 'translateY(20px)', transition: `all 0.5s ease ${i * 0.11}s`, cursor: 'default' }}
                  onMouseEnter={e => { const el = e.currentTarget as HTMLElement; el.style.borderColor = t.color; el.style.boxShadow = `0 10px 28px rgba(0,0,0,0.2), 0 0 0 1px ${t.color}`; }}
                  onMouseLeave={e => { const el = e.currentTarget as HTMLElement; el.style.borderColor = 'var(--border)'; el.style.boxShadow = 'none'; }}
                >
                  <div style={{ display: 'flex', gap: 2, marginBottom: 14 }}>
                    {Array(t.rating).fill(0).map((_, s) => <span key={s} style={{ color: 'var(--text)', fontSize: 14 }}>★</span>)}
                  </div>
                  <p style={{ fontSize: '0.84rem', lineHeight: 1.7, color: 'var(--text-sec)', marginBottom: 18, fontStyle: 'italic' }}>"{t.quote}"</p>
                  <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
                    <div style={{ width: 34, height: 34, borderRadius: '50%', background: `${t.color}18`, border: `1px solid ${t.color}35`, display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: '0.75rem', fontWeight: 700, color: t.color, flexShrink: 0 }}>{t.name[0]}</div>
                    <div><div style={{ fontSize: '0.8rem', fontWeight: 600, color: 'var(--text)' }}>{t.name}</div><div style={{ fontSize: '0.7rem', color: 'var(--text-muted)' }}>{t.role}</div></div>
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      </section>

      {/* Tech Stack */}
      <section style={{ padding: '48px 28px', borderBottom: '1px solid var(--border)' }}>
        <div style={{ maxWidth: 900, margin: '0 auto' }}>
          <div className="section-label">Tech Stack</div>
          <h3 style={{ marginBottom: 22, fontSize: '0.9rem', color: 'var(--text)' }}>Built with the best tools in the ecosystem</h3>
          <div style={{ display: 'flex', flexWrap: 'wrap', gap: 8 }}>
            {TECH.map((t, i) => <TechBadge key={t} label={t} idx={i}/>)}
          </div>
        </div>
      </section>

      {/* FAQ */}
      <section style={{ padding: '64px 28px', borderBottom: '1px solid var(--border)', background: 'var(--bg-2)' }}>
        <div style={{ maxWidth: 640, margin: '0 auto' }}>
          <div className="section-label">FAQ</div>
          <h2 style={{ marginBottom: 34, color: 'var(--text)' }}>Common questions</h2>
          {FAQS.map((f, i) => <FaqRow key={i} faq={f} idx={i} open={openFaq === i} onToggle={() => setOpenFaq(openFaq === i ? null : i)}/>)}
        </div>
      </section>

      {/* CTA */}
      <section style={{ padding: '80px 28px', textAlign: 'center', borderBottom: '1px solid var(--border)', position: 'relative', overflow: 'hidden' }}>
        <div style={{ position: 'absolute', inset: 0, backgroundImage: `linear-gradient(var(--border) 1px, transparent 1px), linear-gradient(90deg, var(--border) 1px, transparent 1px)`, backgroundSize: '32px 32px', opacity: 0.18, pointerEvents: 'none' }}/>
        <div style={{ position: 'absolute', width: 440, height: 440, borderRadius: '50%', background: 'radial-gradient(circle, rgba(255,255,255,0.04) 0%, transparent 70%)', top: '50%', left: '50%', transform: 'translate(-50%,-50%)', pointerEvents: 'none' }}/>
        <div style={{ position: 'relative', zIndex: 1 }}>
          <div className="section-label" style={{ textAlign: 'center', marginBottom: 16 }}>Get started today</div>
          <h2 style={{ fontSize: 'clamp(1.7rem, 4vw, 2.5rem)', marginBottom: 14, letterSpacing: -1, color: 'var(--text)' }}>Build smarter spreadsheets with AI.</h2>
          <p style={{ marginBottom: 38, color: 'var(--text-sec)', maxWidth: 440, margin: '0 auto 38px' }}>50 free generations per month. No credit card required. Start in under 60 seconds.</p>
          <div style={{ display: 'flex', gap: 10, justifyContent: 'center' }}>
            <button className="btn btn-primary btn-xl hlift" onClick={() => nav('/auth')} style={{ boxShadow: '0 0 28px rgba(255,255,255,0.12)' }}>Start for free →</button>
            <button className="btn btn-outline btn-xl hlift" onClick={() => nav('/dashboard')}>Skip to dashboard</button>
          </div>
          <div style={{ marginTop: 22, fontSize: '0.75rem', color: 'var(--text-muted)' }}>
            <span style={{ marginRight: 18 }}>✓ No credit card</span>
            <span style={{ marginRight: 18 }}>✓ Cancel anytime</span>
            <span>✓ GDPR compliant</span>
          </div>
        </div>
      </section>

      {/* Footer */}
      <footer style={{ padding: '22px 28px', display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: 12 }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
          <span style={{ fontFamily: 'var(--font-display)', fontWeight: 800, fontSize: 13, color: 'var(--text)' }}>XtronExcel</span>
          <span style={{ color: 'var(--text-muted)', fontSize: '0.75rem' }}>© 2026</span>
        </div>
        <div style={{ display: 'flex', gap: 16 }}>
          {['Privacy', 'Terms', 'Docs', 'GitHub', 'Contact'].map(l => (
            <span key={l} style={{ fontSize: '0.75rem', color: 'var(--text-muted)', cursor: 'pointer', transition: 'color var(--tr)' }}
              onMouseEnter={e => (e.currentTarget.style.color = 'var(--text)')}
              onMouseLeave={e => (e.currentTarget.style.color = 'var(--text-muted)')}
            >{l}</span>
          ))}
        </div>
        <div style={{ fontSize: '0.7rem', fontFamily: 'var(--font-mono)', color: 'var(--text-muted)', letterSpacing: '0.05em' }}>Powered by Datum_GLAU · 2026</div>
      </footer>
    </div>
  );
}
