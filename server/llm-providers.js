// ════════════════════════════════════════════════════════════════
// llm-providers.js — Drop this file in your server/ folder
// Supports: Gemini (primary) + HuggingFace (fallback)
// ════════════════════════════════════════════════════════════════

const GEMINI_MODEL = 'gemini-2.0-flash'
const HF_MODEL = 'Qwen/Qwen2.5-7B-Instruct'
const HF_CHAT_URL = 'https://router.huggingface.co/v1/chat/completions'

// ── GEMINI ────────────────────────────────────────────────────────
async function callGemini(prompt, systemPrompt, jsonMode = false, apiKey) {
  const key = apiKey || process.env.GEMINI_API_KEY
  if (!key) throw new Error('NO_GEMINI_KEY')

  const url = `https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODEL}:generateContent?key=${key}`
  const system = systemPrompt || (jsonMode
    ? 'You are a data manipulation AI. Respond with valid JSON only. No markdown fences, no explanation, just raw JSON.'
    : 'You are Xtron, an expert AI assistant for student data analysis. Be concise, specific, and helpful.')

  const res = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      systemInstruction: { parts: [{ text: system }] },
      contents: [{ role: 'user', parts: [{ text: prompt }] }],
      generationConfig: { temperature: jsonMode ? 0.1 : 0.7, maxOutputTokens: 4096 }
    })
  })

  if (!res.ok) {
    const err = await res.json()
    const msg = err.error?.message || `Gemini ${res.status}`
    if (res.status === 429 || msg.includes('quota') || msg.includes('RESOURCE_EXHAUSTED')) {
      throw new Error('QUOTA_EXCEEDED')
    }
    throw new Error(msg)
  }

  const data = await res.json()
  const text = data.candidates?.[0]?.content?.parts?.[0]?.text || ''
  return jsonMode ? parseJSON(text) : text
}

// ── HUGGING FACE (FIXED - Proper text extraction) ─────────────────
async function callHuggingFace(prompt, systemPrompt, jsonMode = false, apiKey) {
  const key = apiKey || process.env.HF_API_KEY
  if (!key) throw new Error('NO_HF_KEY')

  const system = systemPrompt || (jsonMode
    ? 'You are a data manipulation AI. Respond with valid JSON only. No markdown fences, no explanation, just raw JSON.'
    : 'You are Xtron, an expert AI assistant for student data analysis. Be concise, specific, and helpful.')

  try {
    const res = await fetch(HF_CHAT_URL, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${key}`
      },
      body: JSON.stringify({
        model: HF_MODEL,
        messages: [
          { role: 'system', content: system },
          { role: 'user', content: prompt }
        ],
        max_tokens: 2048,
        temperature: jsonMode ? 0.1 : 0.7
      })
    })

    if (!res.ok) {
      const err = await res.json().catch(() => ({}))
      throw new Error(normalizeError(err.error || err, `HuggingFace ${res.status}`))
    }

    const data = await res.json()
    let text = extractHFContent(data?.choices?.[0]?.message?.content)
    
    // Clean up any object stringification
    if (text === '[object Object]' || text.includes('[object Object]')) {
      text = 'I processed your request. Here is the response: ' + JSON.stringify(data)
    }
    
    return jsonMode ? parseJSON(text) : text
  } catch (error) {
    console.error('HuggingFace error:', error)
    throw error
  }
}

// ── STREAMING: GEMINI ─────────────────────────────────────────────
async function streamGemini(prompt, systemPrompt, apiKey, onChunk, onDone, onError) {
  const key = apiKey || process.env.GEMINI_API_KEY
  if (!key) { onError('NO_GEMINI_KEY'); return false }

  const url = `https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODEL}:streamGenerateContent?alt=sse&key=${key}`
  const system = systemPrompt || 'You are Xtron, expert AI for student data analysis. Be concise and use actual data.'

  try {
    const res = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        systemInstruction: { parts: [{ text: system }] },
        contents: [{ role: 'user', parts: [{ text: prompt }] }],
        generationConfig: { temperature: 0.7, maxOutputTokens: 2048 }
      })
    })

    if (!res.ok) {
      const errData = await res.json()
      const msg = errData.error?.message || `Gemini ${res.status}`
      if (res.status === 429 || msg.includes('quota') || msg.includes('RESOURCE_EXHAUSTED')) {
        return false
      }
      onError(msg); return true
    }

    const reader = res.body.getReader()
    const decoder = new TextDecoder()
    let buf = ''

    while (true) {
      const { done, value } = await reader.read()
      if (done) break
      buf += decoder.decode(value, { stream: true })
      const lines = buf.split('\n')
      buf = lines.pop() || ''
      for (const line of lines) {
        if (!line.startsWith('data: ')) continue
        const raw = line.slice(6).trim()
        if (raw === '[DONE]') continue
        try {
          const parsed = JSON.parse(raw)
          const text = parsed.candidates?.[0]?.content?.parts?.[0]?.text
          if (text) onChunk(text)
        } catch {}
      }
    }
    onDone(); return true
  } catch (e) {
    if (e.message?.includes('quota') || e.message?.includes('RESOURCE_EXHAUSTED')) return false
    onError(e.message); return true
  }
}

// ── STREAMING: HUGGING FACE (FIXED) ───────────────────────────────
async function streamHuggingFace(prompt, systemPrompt, apiKey, onChunk, onDone, onError) {
  const key = apiKey || process.env.HF_API_KEY
  if (!key) { onError('NO_HF_KEY'); return }

  const system = systemPrompt || 'You are Xtron, expert AI for student data analysis.'

  try {
    const res = await fetch(HF_CHAT_URL, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${key}`
      },
      body: JSON.stringify({
        model: HF_MODEL,
        messages: [
          { role: 'system', content: system },
          { role: 'user', content: prompt }
        ],
        max_tokens: 2048,
        temperature: 0.7,
        stream: true
      })
    })

    if (!res.ok) {
      const err = await res.json().catch(() => ({}))
      onError(normalizeError(err.error || err, `HuggingFace ${res.status}`))
      return
    }

    const reader = res.body.getReader()
    const decoder = new TextDecoder()
    let buf = ''

    while (true) {
      const { done, value } = await reader.read()
      if (done) break
      buf += decoder.decode(value, { stream: true })
      const lines = buf.split('\n')
      buf = lines.pop() || ''

      for (const line of lines) {
        if (!line.startsWith('data: ')) continue
        const raw = line.slice(6).trim()
        if (!raw || raw === '[DONE]') continue

        try {
          const parsed = JSON.parse(raw)
          const text = extractHFContent(parsed?.choices?.[0]?.delta?.content)
          if (text) onChunk(text)
        } catch {}
      }
    }

    onDone()
  } catch (e) {
    onError(normalizeError(e, 'HuggingFace request failed'))
  }
}

function extractHFContent(content) {
  if (typeof content === 'string') return content
  if (content && typeof content === 'object') {
    if (typeof content.text === 'string') return content.text
    if (typeof content.content === 'string') return content.content
    if (Array.isArray(content.content)) return extractHFContent(content.content)
  }
  if (Array.isArray(content)) {
    return content
      .map(part => {
        if (typeof part === 'string') return part
        if (part?.type === 'text') return part.text || ''
        if (part && typeof part === 'object') return extractHFContent(part)
        return ''
      })
      .join('')
  }
  return ''
}

function normalizeError(error, fallback = 'Unknown error') {
  if (!error) return fallback
  if (typeof error === 'string') return error
  if (error instanceof Error) return error.message || fallback
  if (typeof error === 'object') {
    if (typeof error.message === 'string') return error.message
    if (typeof error.error === 'string') return error.error
    if (typeof error.detail === 'string') return error.detail
    try { return JSON.stringify(error) } catch {}
  }
  return String(error)
}

// ── SHARED JSON PARSER ────────────────────────────────────────────
function parseJSON(text) {
  if (!text || text === '[object Object]') {
    throw new Error('Invalid response from LLM')
  }
  const cleaned = text.replace(/```json\n?/g, '').replace(/```\n?/g, '').trim()
  try { 
    return JSON.parse(cleaned) 
  }
  catch { 
    throw new Error('Invalid JSON from LLM: ' + cleaned.slice(0, 300)) 
  }
}

// ── UNIFIED CALL (auto-fallback) ──────────────────────────────────
async function callLLM(prompt, systemPrompt, jsonMode, keys = {}) {
  if (keys.gemini || process.env.GEMINI_API_KEY) {
    try {
      return await callGemini(prompt, systemPrompt, jsonMode, keys.gemini)
    } catch (e) {
      if (e.message !== 'QUOTA_EXCEEDED') throw e
      console.log('[LLM] Gemini quota exceeded, falling back to HuggingFace...')
    }
  }

  if (keys.hf || process.env.HF_API_KEY) {
    return await callHuggingFace(prompt, systemPrompt, jsonMode, keys.hf)
  }

  throw new Error('No API keys configured. Set GEMINI_API_KEY or HF_API_KEY.')
}

// ── UNIFIED STREAM (auto-fallback) ────────────────────────────────
async function streamLLM(prompt, systemPrompt, keys = {}, onChunk, onDone, onError) {
  if (keys.gemini || process.env.GEMINI_API_KEY) {
    const handled = await streamGemini(prompt, systemPrompt, keys.gemini, onChunk, onDone, onError)
    if (handled) return
    console.log('[LLM] Gemini quota exceeded, streaming from HuggingFace...')
  }

  if (keys.hf || process.env.HF_API_KEY) {
    await streamHuggingFace(prompt, systemPrompt, keys.hf, onChunk, onDone, onError)
    return
  }

  onError('No API keys configured.')
}

module.exports = { callLLM, streamLLM, callGemini, callHuggingFace, streamGemini, streamHuggingFace }
