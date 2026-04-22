/**
 * RightCut — Session history store.
 * Reads/writes to Supabase when available, otherwise uses localStorage.
 */
import { create } from 'zustand'
import { supabase, supabaseEnabled } from '../lib/supabase'

const LOCAL_KEY = 'rightcut_sessions'

function loadLocalSessions() {
  try {
    return JSON.parse(localStorage.getItem(LOCAL_KEY) || '[]')
  } catch { return [] }
}

function saveLocalSessions(sessions) {
  try { localStorage.setItem(LOCAL_KEY, JSON.stringify(sessions)) } catch {}
}

const useHistoryStore = create((set, get) => ({
  sessions: [],
  loading: false,

  // Load all sessions for the current user
  loadSessions: async (userId) => {
    if (!supabaseEnabled || !userId) {
      set({ sessions: loadLocalSessions() })
      return
    }
    set({ loading: true })
    const { data, error } = await supabase
      .from('sessions')
      .select('id, session_id, title, created_at, updated_at')
      .eq('user_id', userId)
      .order('updated_at', { ascending: false })
      .limit(50)
    set({ loading: false, sessions: error ? loadLocalSessions() : (data || []) })
  },

  // Create or update a session record
  upsertSession: async (sessionId, userId, title) => {
    const now = new Date().toISOString()
    if (!supabaseEnabled || !userId) {
      const existing = loadLocalSessions()
      const idx = existing.findIndex((s) => s.session_id === sessionId)
      const entry = { id: sessionId, session_id: sessionId, title, created_at: now, updated_at: now }
      if (idx >= 0) {
        existing[idx] = { ...existing[idx], title, updated_at: now }
      } else {
        existing.unshift(entry)
      }
      saveLocalSessions(existing)
      set({ sessions: existing })
      return
    }
    await supabase.from('sessions').upsert({
      session_id: sessionId,
      user_id: userId,
      title,
      updated_at: now,
    }, { onConflict: 'session_id' })
    get().loadSessions(userId)
  },

  // Save messages to Supabase
  saveMessage: async (sessionId, userId, message) => {
    if (!supabaseEnabled || !userId) return
    await supabase.from('messages').upsert({
      message_id: message.id,
      session_id: sessionId,
      user_id: userId,
      role: message.role,
      text: message.text,
      metadata: message.sheetRefs ? { sheetRefs: message.sheetRefs } : null,
      created_at: new Date(message.timestamp).toISOString(),
    }, { onConflict: 'message_id' })
  },

  // Load messages for a specific session from Supabase
  loadMessages: async (sessionId, userId) => {
    if (!supabaseEnabled || !userId) return []
    const { data, error } = await supabase
      .from('messages')
      .select('message_id, role, text, created_at, metadata')
      .eq('session_id', sessionId)
      .order('created_at', { ascending: true })
    if (error) return []
    return (data || []).map((m) => ({
      id: m.message_id,
      role: m.role,
      text: m.text,
      timestamp: new Date(m.created_at).getTime(),
      timeline: [],
      sheetRefs: m.metadata?.sheetRefs || [],
    }))
  },

  // Save workbook snapshot for a session
  saveWorkbook: async (sessionId, userId, workbookState) => {
    if (!workbookState) return
    // Always save to localStorage keyed by session
    try {
      localStorage.setItem(`rightcut_wb_${sessionId}`, JSON.stringify(workbookState))
    } catch {}
    if (!supabaseEnabled || !userId) return
    await supabase.from('workbook_snapshots').upsert({
      session_id: sessionId,
      user_id: userId,
      snapshot: workbookState,
      updated_at: new Date().toISOString(),
    }, { onConflict: 'session_id' })
  },

  // Load workbook snapshot for a session
  loadWorkbook: async (sessionId, userId) => {
    // Try Supabase first
    if (supabaseEnabled && userId) {
      const { data } = await supabase
        .from('workbook_snapshots')
        .select('snapshot')
        .eq('session_id', sessionId)
        .single()
      if (data?.snapshot) return data.snapshot
    }
    // Fallback to localStorage
    try {
      const raw = localStorage.getItem(`rightcut_wb_${sessionId}`)
      if (raw) return JSON.parse(raw)
    } catch {}
    return null
  },
}))

export default useHistoryStore
