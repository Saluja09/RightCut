/**
 * RightCut — Zustand workspace store.
 * Single source of truth for all UI state.
 *
 * WORKBOOK DATA FLOW (unidirectional):
 *   Backend → (WS workbook_update) → workspaceStore.workbookState → UI render
 *   workspaceStore.workbookState → (subscriber) → historyStore.saveWorkbook (backup)
 *   historyStore.loadWorkbook → POST /restore → Backend → (WS workbook_update) → store
 *
 * The frontend NEVER renders workbook data loaded directly from localStorage/Supabase.
 * It only uses saved snapshots to restore the backend, then renders what the backend sends.
 */
import { create } from 'zustand'

// Persist only sessionRole across reloads (not workbook — backend is source of truth)
const PERSIST_KEY = 'rightcut_store_v2'

// Migration: clear old workbook persistence keys that caused cross-session contamination.
// The old system stored workbook snapshots in localStorage and rendered them directly.
// The new system only uses the backend as the source of truth.
const MIGRATION_KEY = 'rightcut_migration_v2'
if (!localStorage.getItem(MIGRATION_KEY)) {
  const keysToRemove = []
  for (let i = 0; i < localStorage.length; i++) {
    const key = localStorage.key(i)
    if (key?.startsWith('rightcut_wb_') || key === 'rightcut_store' || key === 'rightcut_migration_v1') {
      keysToRemove.push(key)
    }
  }
  keysToRemove.forEach((k) => localStorage.removeItem(k))
  localStorage.setItem(MIGRATION_KEY, '1')
}

function loadPersistedState() {
  try {
    const raw = localStorage.getItem(PERSIST_KEY)
    if (!raw) return {}
    const parsed = JSON.parse(raw)
    return { sessionRole: parsed.sessionRole || null }
  } catch (_) {
    return {}
  }
}

function persistState(state) {
  try {
    localStorage.setItem(PERSIST_KEY, JSON.stringify({
      sessionId: state.sessionId,
      sessionRole: state.sessionRole,
    }))
  } catch (_) {}
}

const useWorkspaceStore = create((set, get) => ({
  ...loadPersistedState(),
  // ── Session ──────────────────────────────────────────────────────────────
  sessionId: null,
  initSession: () => {
    const existing = localStorage.getItem('rightcut_session_id')
    const id = existing || crypto.randomUUID()
    localStorage.setItem('rightcut_session_id', id)
    set({ sessionId: id })
    return id
  },

  // ── Chat messages ─────────────────────────────────────────────────────────
  messages: [],
  addMessage: (msg) =>
    set((s) => {
      const id = msg.id || crypto.randomUUID()
      const exists = s.messages.findIndex((m) => m.id === id)
      if (exists >= 0) {
        const updated = [...s.messages]
        updated[exists] = { ...updated[exists], ...msg, id }
        return { messages: updated }
      }
      return { messages: [...s.messages, { ...msg, id }] }
    }),

  updateMessage: (id, patch) =>
    set((s) => ({
      messages: s.messages.map((m) => (m.id === id ? { ...m, ...patch } : m)),
    })),

  // ── Tool timeline ─────────────────────────────────────────────────────────
  toolTimelines: {},
  addToolStep: (messageId, step) =>
    set((s) => ({
      toolTimelines: {
        ...s.toolTimelines,
        [messageId]: [...(s.toolTimelines[messageId] || []), step],
      },
    })),

  // ── Pending agent message ─────────────────────────────────────────────────
  pendingMessageId: null,
  setPendingMessageId: (id) => set({ pendingMessageId: id }),

  // ── Workbook state (set ONLY from backend WS workbook_update) ─────────────
  workbookState: null,
  setWorkbookState: (state) => {
    const activeSheet = state.active_sheet || state.sheets?.[0]?.name || null

    // Build tabs from incoming workbook state
    const newTabs = []
    for (const sheet of state.sheets || []) {
      newTabs.push({ id: sheet.name, name: sheet.name, type: 'sheet' })
      for (let i = 0; i < (sheet.charts || []).length; i++) {
        const chart = sheet.charts[i]
        const chartId = `chart__${sheet.name}__${i}`
        newTabs.push({
          id: chartId,
          name: chart.title || `${sheet.name} Chart`,
          type: 'chart',
          sheetName: sheet.name,
          chartIndex: i,
        })
      }
    }

    // Preserve document tabs
    const { tabs: oldTabs } = get()
    const docTabs = oldTabs.filter((t) => t.type === 'document')
    const allTabs = [...newTabs, ...docTabs]

    const currentActive = get().activeTab
    const validActive = allTabs.find((t) => t.id === currentActive)
      ? currentActive
      : (activeSheet || allTabs[0]?.id || null)

    set({
      workbookState: state,
      activeSheet,
      tabs: allTabs,
      activeTab: validActive,
    })
  },

  // ── Active sheet ──────────────────────────────────────────────────────────
  activeSheet: null,
  setActiveSheet: (name) => set({ activeSheet: name }),

  // ── Tabs ──────────────────────────────────────────────────────────────────
  tabs: [],
  activeTab: null,
  addTab: (tab) =>
    set((s) => ({
      tabs: s.tabs.find((t) => t.id === tab.id) ? s.tabs : [...s.tabs, tab],
      activeTab: s.activeTab ?? tab.id,
    })),
  setActiveTab: (id) => {
    const { tabs } = get()
    const tab = tabs.find((t) => t.id === id)
    set({ activeTab: id })
    if (tab?.type === 'sheet') {
      set({ activeSheet: id })
    }
  },

  // ── Uploaded documents ────────────────────────────────────────────────────
  documents: {},
  addDocument: (fileId, doc) =>
    set((s) => ({
      documents: { ...s.documents, [fileId]: doc },
    })),
  pendingFileIds: [],
  stagePendingFile: (fileId) =>
    set((s) => ({ pendingFileIds: [...s.pendingFileIds, fileId] })),
  clearPendingFiles: () => set({ pendingFileIds: [] }),

  // ── Session restore ───────────────────────────────────────────────────────
  // Holds a saved workbook snapshot for the WS session_ready handler to restore.
  // Flow: handleSelectSession loads from storage → sets pendingRestore →
  //       WS connects → session_ready(has_workbook:false) → POST /restore →
  //       backend sends workbook_update → store.setWorkbookState → UI renders.
  pendingRestore: null,

  // ── Session role ──────────────────────────────────────────────────────────
  sessionRole: null,
  setSessionRole: (role) => set({ sessionRole: role }),

  // ── Session restore status ─────────────────────────────────────────────────
  restoring: false,
  setRestoring: (v) => set({ restoring: v }),

  // ── WebSocket / agent status ───────────────────────────────────────────────
  wsStatus: 'disconnected',
  setWsStatus: (status) => set({ wsStatus: status }),

  // ── Validation stats ──────────────────────────────────────────────────────
  formulaCount: 0,
  hardcodedCount: 0,
  sheetCount: 0,
  setValidationStats: ({ formula_count = 0, hardcoded_count = 0 }) =>
    set({ formulaCount: formula_count, hardcodedCount: hardcoded_count }),
  updateSheetCount: () =>
    set((s) => ({ sheetCount: s.workbookState?.sheets?.length || 0 })),
}))

// Persist only lightweight state (sessionRole) — workbook is NOT persisted here
let _persistTimer = null
useWorkspaceStore.subscribe((state) => {
  clearTimeout(_persistTimer)
  _persistTimer = setTimeout(() => persistState(state), 300)
})

export default useWorkspaceStore
