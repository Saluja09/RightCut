/**
 * RightCut — Zustand workspace store.
 * Single source of truth for all UI state.
 */
import { create } from 'zustand'

// Keys to persist across reloads
const PERSIST_KEY = 'rightcut_store'

function loadPersistedState() {
  try {
    const raw = localStorage.getItem(PERSIST_KEY)
    if (!raw) return {}
    const { workbookState, tabs, activeTab, activeSheet } = JSON.parse(raw)
    // Don't restore messages — avoids duplicate key issues and stale pending states
    const wb = workbookState || null
    return {
      workbookState: wb,
      tabs: tabs || [],
      activeTab: activeTab || null,
      activeSheet: activeSheet || null,
      // If we have a saved workbook, mark it for backend restore on next WS connect
      pendingRestore: wb ? { workbook_state: wb, messages: [] } : null,
      // Existing sessions (have workbook) skip the role modal; fresh sessions show it
      sessionRole: wb ? 'finance' : null,
    }
  } catch (_) {
    return {}
  }
}

function persistState(state) {
  try {
    const { workbookState, tabs, activeTab, activeSheet } = state
    localStorage.setItem(PERSIST_KEY, JSON.stringify({ workbookState, tabs, activeTab, activeSheet }))
  } catch (_) {}
}

const useWorkspaceStore = create((set, get) => ({
  ...loadPersistedState(),
  // ── Session ──────────────────────────────────────────────────────────────
  sessionId: null,
  initSession: () => {
    // Reuse existing session across reloads
    const existing = localStorage.getItem('rightcut_session_id')
    const id = existing || crypto.randomUUID()
    localStorage.setItem('rightcut_session_id', id)
    set({ sessionId: id })
    return id
  },

  // ── Chat messages ─────────────────────────────────────────────────────────
  // message shape: { id, role: 'user'|'agent'|'error', text, timestamp, timeline? }
  messages: [],
  addMessage: (msg) =>
    set((s) => ({ messages: [...s.messages, { id: msg.id || crypto.randomUUID(), ...msg }] })),

  // Update an existing message by id (e.g. append streaming text)
  updateMessage: (id, patch) =>
    set((s) => ({
      messages: s.messages.map((m) => (m.id === id ? { ...m, ...patch } : m)),
    })),

  // ── Tool timeline ─────────────────────────────────────────────────────────
  // Maps pendingMessageId → ToolStep[]
  toolTimelines: {},
  addToolStep: (messageId, step) =>
    set((s) => ({
      toolTimelines: {
        ...s.toolTimelines,
        [messageId]: [...(s.toolTimelines[messageId] || []), step],
      },
    })),

  // ── Pending agent message (accumulates tool steps before agent responds) ──
  pendingMessageId: null,
  setPendingMessageId: (id) => set({ pendingMessageId: id }),

  // ── Workbook state ────────────────────────────────────────────────────────
  // workbookState shape: { sheets: SheetState[], active_sheet: string }
  workbookState: null,
  setWorkbookState: (state) => {
    const activeSheet = state.active_sheet || state.sheets?.[0]?.name || null
    set({ workbookState: state, activeSheet })

    // Auto-register sheet tabs
    const { addTab } = get()
    for (const sheet of state.sheets || []) {
      addTab({ id: sheet.name, name: sheet.name, type: 'sheet' })
    }
  },

  // ── Active sheet ──────────────────────────────────────────────────────────
  activeSheet: null,
  setActiveSheet: (name) => set({ activeSheet: name }),

  // ── Tabs ──────────────────────────────────────────────────────────────────
  // tab shape: { id, name, type: 'sheet'|'document' }
  tabs: [],
  activeTab: null,
  addTab: (tab) =>
    set((s) => ({
      tabs: s.tabs.find((t) => t.id === tab.id) ? s.tabs : [...s.tabs, tab],
      activeTab: s.activeTab ?? tab.id,
    })),
  setActiveTab: (id) => {
    const { tabs, workbookState } = get()
    const tab = tabs.find((t) => t.id === id)
    set({ activeTab: id })
    if (tab?.type === 'sheet') {
      set({ activeSheet: id })
    }
  },

  // ── Uploaded documents ────────────────────────────────────────────────────
  // documents shape: { [file_id]: { filename, file_type, file_id } }
  documents: {},
  addDocument: (fileId, doc) =>
    set((s) => ({
      documents: { ...s.documents, [fileId]: doc },
    })),
  pendingFileIds: [],   // file_ids staged for next message
  stagePendingFile: (fileId) =>
    set((s) => ({ pendingFileIds: [...s.pendingFileIds, fileId] })),
  clearPendingFiles: () => set({ pendingFileIds: [] }),

  // ── Session restore ───────────────────────────────────────────────────────
  // Set when switching to a saved session; cleared by WS session_ready handler
  pendingRestore: null,

  // ── Session role ──────────────────────────────────────────────────────────
  // 'finance' | 'general' — set by RoleSelectModal at session start
  sessionRole: null,
  setSessionRole: (role) => set({ sessionRole: role }),

  // ── WebSocket / agent status ───────────────────────────────────────────────
  // 'disconnected' | 'connecting' | 'connected' | 'thinking'
  wsStatus: 'disconnected',
  setWsStatus: (status) => set({ wsStatus: status }),

  // ── Validation stats (from validate_workbook tool) ────────────────────────
  formulaCount: 0,
  hardcodedCount: 0,
  sheetCount: 0,
  setValidationStats: ({ formula_count = 0, hardcoded_count = 0 }) =>
    set({ formulaCount: formula_count, hardcodedCount: hardcoded_count }),
  updateSheetCount: () =>
    set((s) => ({ sheetCount: s.workbookState?.sheets?.length || 0 })),
}))

// Persist relevant state to localStorage on every change
useWorkspaceStore.subscribe((state) => persistState(state))

export default useWorkspaceStore
