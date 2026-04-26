/**
 * RightCut — Session history store.
 * Reads/writes to Supabase when available, otherwise uses localStorage.
 */
import { create } from 'zustand'
import { supabase, supabaseEnabled } from '../lib/supabase'
import useAuthStore from './authStore'

// Check if current user is a guest — guest users skip Supabase writes (use localStorage only)
function _isGuestUser() {
  try { return useAuthStore.getState().isGuest } catch { return false }
}

const LOCAL_SESSIONS_KEY  = 'rightcut_sessions'
const LOCAL_MESSAGES_KEY  = (sid) => `rightcut_msgs_${sid}`

// ── localStorage helpers ──────────────────────────────────────────────────────

function loadLocalSessions() {
  try { return JSON.parse(localStorage.getItem(LOCAL_SESSIONS_KEY) || '[]') }
  catch { return [] }
}

function saveLocalSessions(sessions) {
  try { localStorage.setItem(LOCAL_SESSIONS_KEY, JSON.stringify(sessions)) }
  catch {}
}

function loadLocalMessages(sessionId) {
  try { return JSON.parse(localStorage.getItem(LOCAL_MESSAGES_KEY(sessionId)) || '[]') }
  catch { return [] }
}

function saveLocalMessage(sessionId, message) {
  try {
    const existing = loadLocalMessages(sessionId)
    const idx = existing.findIndex((m) => m.id === message.id)
    if (idx >= 0) existing[idx] = message
    else existing.push(message)
    localStorage.setItem(LOCAL_MESSAGES_KEY(sessionId), JSON.stringify(existing))
  } catch {}
}

/** Batch-save all messages at once — much cheaper than N individual saves */
function saveLocalMessagesBatch(sessionId, messages) {
  try {
    const saveable = messages
      .filter((m) => m.role === 'user' || m.role === 'agent')
      .map(({ id, role, text, timestamp, sheetRefs }) => ({ id, role, text, timestamp, sheetRefs }))
    localStorage.setItem(LOCAL_MESSAGES_KEY(sessionId), JSON.stringify(saveable))
  } catch {}
}

// ── Workbook save debounce ────────────────────────────────────────────────────
const _wbTimers = {}
function debounceWbSave(sessionId, fn, delay = 2000) {
  clearTimeout(_wbTimers[sessionId])
  _wbTimers[sessionId] = setTimeout(fn, delay)
}

// ── Store ─────────────────────────────────────────────────────────────────────

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

  // Create or update a session record — updates local state directly, no reload
  upsertSession: async (sessionId, userId, title) => {
    const now = new Date().toISOString()

    // Always update local store immediately so sidebar reflects title right away
    set((s) => {
      const existing = [...s.sessions]
      const idx = existing.findIndex((s) => s.session_id === sessionId)
      if (idx >= 0) {
        existing[idx] = { ...existing[idx], title, updated_at: now }
        // Move to front (most recent)
        const [updated] = existing.splice(idx, 1)
        existing.unshift(updated)
      } else {
        existing.unshift({ id: sessionId, session_id: sessionId, title, created_at: now, updated_at: now })
      }
      return { sessions: existing }
    })

    if (!supabaseEnabled || !userId || _isGuestUser()) {
      saveLocalSessions(get().sessions)
      return
    }

    // Fire-and-forget to Supabase — no reload needed
    supabase.from('sessions').upsert({
      session_id: sessionId,
      user_id: userId,
      title,
      updated_at: now,
    }, { onConflict: 'session_id' }).then(({ error }) => {
      if (error) console.warn('upsertSession Supabase error:', error.message)
    })
  },

  // Save a single message (Supabase + localStorage)
  saveMessage: async (sessionId, userId, message) => {
    saveLocalMessage(sessionId, message)

    if (!supabaseEnabled || !userId || _isGuestUser()) return

    supabase.from('messages').upsert({
      message_id: message.id,
      session_id: sessionId,
      user_id: userId,
      role: message.role,
      text: message.text,
      metadata: message.sheetRefs ? { sheetRefs: message.sheetRefs } : null,
      created_at: new Date(message.timestamp).toISOString(),
    }, { onConflict: 'message_id' }).then(({ error }) => {
      if (error) console.warn('saveMessage Supabase error:', error.message)
    })
  },

  // Batch-save all messages for a session — single localStorage write + bulk Supabase upsert
  saveAllMessages: async (sessionId, userId, messages) => {
    const saveable = messages.filter((m) => m.role === 'user' || m.role === 'agent')
    if (saveable.length === 0) return

    // Single localStorage write
    saveLocalMessagesBatch(sessionId, saveable)

    if (!supabaseEnabled || !userId || _isGuestUser()) return

    // Bulk upsert to Supabase
    const rows = saveable.map((m) => ({
      message_id: m.id,
      session_id: sessionId,
      user_id: userId,
      role: m.role,
      text: m.text,
      metadata: m.sheetRefs ? { sheetRefs: m.sheetRefs } : null,
      created_at: new Date(m.timestamp).toISOString(),
    }))
    supabase.from('messages').upsert(rows, { onConflict: 'message_id' }).then(({ error }) => {
      if (error) console.warn('saveAllMessages Supabase error:', error.message)
    })
  },

  // Load messages for a specific session
  loadMessages: async (sessionId, userId) => {
    if (supabaseEnabled && userId) {
      const { data, error } = await supabase
        .from('messages')
        .select('message_id, role, text, created_at, metadata')
        .eq('session_id', sessionId)
        .order('created_at', { ascending: true })
      if (!error && data?.length) {
        return data.map((m) => ({
          id: m.message_id,
          role: m.role,
          text: m.text,
          timestamp: new Date(m.created_at).getTime(),
          timeline: [],
          sheetRefs: m.metadata?.sheetRefs || [],
        }))
      }
    }
    // Fallback to localStorage
    return loadLocalMessages(sessionId)
  },

  // Save workbook snapshot — debounced to avoid spamming on every cell update
  saveWorkbook: async (sessionId, userId, workbookState) => {
    if (!workbookState) return

    // Always save to localStorage immediately (keyed by session)
    // Include _sessionId tag so we can detect cross-session corruption on load
    try {
      const tagged = { ...workbookState, _sessionId: sessionId }
      localStorage.setItem(`rightcut_wb_${sessionId}`, JSON.stringify(tagged))
    } catch {}

    if (!supabaseEnabled || !userId || _isGuestUser()) return

    // Debounce Supabase writes: only flush after 2s of inactivity
    debounceWbSave(sessionId, () => {
      supabase.from('workbook_snapshots').upsert({
        session_id: sessionId,
        user_id: userId,
        snapshot: { ...workbookState, _sessionId: sessionId },
        updated_at: new Date().toISOString(),
      }, { onConflict: 'session_id' }).then(({ error }) => {
        if (error) console.warn('saveWorkbook Supabase error:', error.message)
      })
    })
  },

  // Load workbook snapshot for a session
  loadWorkbook: async (sessionId, userId) => {
    if (supabaseEnabled && userId) {
      const { data } = await supabase
        .from('workbook_snapshots')
        .select('snapshot')
        .eq('session_id', sessionId)
        .single()
      if (data?.snapshot) {
        const snap = data.snapshot
        // Verify the snapshot belongs to this session (detect cross-session corruption)
        if (snap._sessionId && snap._sessionId !== sessionId) {
          // Corrupted entry — delete it
          supabase.from('workbook_snapshots').delete().eq('session_id', sessionId)
          return null
        }
        // Old entry without tag — could be corrupted, discard it
        if (!snap._sessionId) {
          supabase.from('workbook_snapshots').delete().eq('session_id', sessionId)
          return null
        }
        // Remove the internal tag before returning
        const { _sessionId, ...wb } = snap
        return wb
      }
    }
    // Fallback to per-session localStorage key
    try {
      const raw = localStorage.getItem(`rightcut_wb_${sessionId}`)
      if (raw) {
        const parsed = JSON.parse(raw)
        // Verify this workbook belongs to the requested session (detect cross-session corruption)
        if (parsed._sessionId && parsed._sessionId !== sessionId) {
          localStorage.removeItem(`rightcut_wb_${sessionId}`)  // clean up corrupted key
          return null
        }
        // Remove the internal tag before returning
        if (parsed._sessionId) {
          const { _sessionId, ...wb } = parsed
          return wb
        }
        // Old entry without session tag — potentially corrupted by a previous
        // bug that saved the wrong workbook under this session's key.
        // Delete it to prevent cross-session data leaks. The workbook will
        // be re-saved correctly (with tag) the next time this session is active.
        localStorage.removeItem(`rightcut_wb_${sessionId}`)
        return null
      }
    } catch {}
    return null
  },

  // Delete a session and its messages/workbook
  deleteSession: async (sessionId, userId) => {
    // Remove from local store
    set((s) => ({ sessions: s.sessions.filter((s) => s.session_id !== sessionId) }))
    saveLocalSessions(get().sessions)

    // Remove localStorage artifacts
    try {
      localStorage.removeItem(LOCAL_MESSAGES_KEY(sessionId))
      localStorage.removeItem(`rightcut_wb_${sessionId}`)
    } catch {}

    if (!supabaseEnabled || !userId) return

    // Cascade delete — messages and workbook_snapshots have FK ON DELETE CASCADE
    // so deleting the session row is enough if schema is set up correctly.
    // Otherwise delete child rows first.
    await Promise.all([
      supabase.from('messages').delete().eq('session_id', sessionId),
      supabase.from('workbook_snapshots').delete().eq('session_id', sessionId),
    ])
    await supabase.from('sessions').delete().eq('session_id', sessionId)
  },
}))

export default useHistoryStore
