/**
 * RightCut — Zustand workspace store.
 * Single source of truth for all UI state.
 */
import { create } from 'zustand'

// Keys to persist across reloads
const PERSIST_KEY = 'rightcut_store'
const MIGRATION_KEY = 'rightcut_migration_v1'

// One-time migration: clear potentially contaminated per-session workbook keys
// from before the cross-session contamination fix.
if (!localStorage.getItem(MIGRATION_KEY)) {
  const keysToRemove = []
  for (let i = 0; i < localStorage.length; i++) {
    const key = localStorage.key(i)
    if (key?.startsWith('rightcut_wb_')) keysToRemove.push(key)
  }
  keysToRemove.forEach((k) => localStorage.removeItem(k))
  // Also clear the global persisted state to remove stale workbooks
  localStorage.removeItem(PERSIST_KEY)
  localStorage.setItem(MIGRATION_KEY, '1')
}

function loadPersistedState() {
  try {
    const raw = localStorage.getItem(PERSIST_KEY)
    if (!raw) return {}
    const parsed = JSON.parse(raw)
    const { workbookState, tabs, activeTab, activeSheet, sessionRole } = parsed

    // Only restore workbook if it belongs to the current session
    // (prevents cross-session contamination after session switches)
    const currentSessionId = localStorage.getItem('rightcut_session_id')
    const storedSessionId = parsed.sessionId
    // If the persisted state includes a sessionId, it must match the current one.
    // If it doesn't include one (old format), discard the workbook — we can't verify ownership.
    const sessionMatch = storedSessionId ? storedSessionId === currentSessionId : false
    const wb = (workbookState && sessionMatch) ? workbookState : null
    const role = sessionRole || (wb ? 'general' : null)

    // If workbook was discarded, also discard tabs/activeSheet to avoid stale UI
    let resolvedTabs = wb ? (tabs || []) : []
    let resolvedActiveTab = wb ? (activeTab || null) : null
    let resolvedActiveSheet = wb ? (activeSheet || null) : null
    if (wb && resolvedTabs.length === 0 && wb.sheets?.length > 0) {
      resolvedTabs = wb.sheets.map((s) => ({ id: s.name, name: s.name, type: 'sheet' }))
      resolvedActiveSheet = resolvedActiveSheet || wb.active_sheet || wb.sheets[0]?.name || null
      resolvedActiveTab = resolvedActiveSheet
    }

    return {
      workbookState: wb,
      tabs: resolvedTabs,
      activeTab: resolvedActiveTab,
      activeSheet: resolvedActiveSheet,
      pendingRestore: wb ? { workbook_state: wb, messages: [] } : null,
      sessionRole: role,
    }
  } catch (_) {
    return {}
  }
}

function persistState(state) {
  try {
    const { sessionId, workbookState, tabs, activeTab, activeSheet, sessionRole } = state
    localStorage.setItem(PERSIST_KEY, JSON.stringify({ sessionId, workbookState, tabs, activeTab, activeSheet, sessionRole }))
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
    set((s) => {
      const id = msg.id || crypto.randomUUID()
      // Deduplicate: if a message with this id already exists, replace it (handles WS retries)
      const exists = s.messages.findIndex((m) => m.id === id)
      if (exists >= 0) {
        const updated = [...s.messages]
        updated[exists] = { ...updated[exists], ...msg, id }
        return { messages: updated }
      }
      return { messages: [...s.messages, { ...msg, id }] }
    }),

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
    const sid = get().sessionId
    console.debug(`[setWorkbookState] sessionId=${sid?.slice(0,8)} sheets=${(state.sheets||[]).map(s=>s.name).join(',')}`)
    const activeSheet = state.active_sheet || state.sheets?.[0]?.name || null

    // Build the complete tab list from the incoming workbook state.
    // This replaces stale tabs from previous sessions — only tabs matching
    // current sheets/charts survive.
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

    // Preserve any document tabs (they aren't in workbook state)
    const { tabs: oldTabs } = get()
    const docTabs = oldTabs.filter((t) => t.type === 'document')
    const allTabs = [...newTabs, ...docTabs]

    // Pick a valid active tab
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

  // ── Session restore status ─────────────────────────────────────────────────
  // true while a /restore call is in flight — blocks sending messages
  restoring: false,
  setRestoring: (v) => set({ restoring: v }),

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

// Persist relevant state to localStorage — debounced to avoid writes on every keystroke/scroll
let _persistTimer = null
useWorkspaceStore.subscribe((state) => {
  clearTimeout(_persistTimer)
  _persistTimer = setTimeout(() => persistState(state), 300)
})

export default useWorkspaceStore
