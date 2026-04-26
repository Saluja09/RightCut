/**
 * RightCut — Root application component.
 * Sidebar + chat + preview layout.
 *
 * WORKBOOK DATA FLOW (unidirectional):
 *   Backend is the single source of truth for workbook state.
 *   Frontend renders what the backend sends via WS workbook_update.
 *   Saved snapshots (localStorage/Supabase) are only used to POST /restore
 *   when the backend session is empty — never rendered directly.
 */
import { useEffect } from 'react'
import ChatPanel from './components/ChatPanel'
import PreviewPanel from './components/PreviewPanel'
import StatusBar from './components/StatusBar'
import LoginPage from './components/LoginPage'
import LeftSidebar from './components/LeftSidebar'
import useWorkspaceStore from './stores/workspaceStore'
import useAuthStore from './stores/authStore'
import useThemeStore from './stores/themeStore'
import useHistoryStore from './stores/historyStore'
import { useWebSocket } from './hooks/useWebSocket'

// Initialize session synchronously so sessionId exists before any hook runs
const store = useWorkspaceStore.getState()
if (!store.sessionId) {
  store.initSession()
}

export default function App() {
  const sessionId   = useWorkspaceStore((s) => s.sessionId)
  const messages    = useWorkspaceStore((s) => s.messages)
  const setSessionRole = useWorkspaceStore((s) => s.setSessionRole)
  const { user, isGuest, loading: authLoading, initAuth } = useAuthStore()
  const { initTheme } = useThemeStore()
  const { upsertSession, saveAllMessages, loadSessions, saveWorkbook } = useHistoryStore()

  const { sendMessage, sendCellEdit } = useWebSocket()

  // Init auth + theme on mount, and reload messages for the current active session
  useEffect(() => {
    initAuth()
    initTheme()
    if (sessionId) {
      // Restore chat messages from localStorage/Supabase
      const hs = useHistoryStore.getState()
      hs.loadMessages(sessionId, undefined).then((msgs) => {
        if (msgs.length > 0 && useWorkspaceStore.getState().messages.length === 0) {
          useWorkspaceStore.setState({ messages: msgs })
        }
      })
      // Load saved workbook snapshot — NOT for rendering, but to feed pendingRestore
      // so the WS session_ready handler can POST /restore if the backend is empty.
      hs.loadWorkbook(sessionId, undefined).then((wb) => {
        if (wb) {
          useWorkspaceStore.setState({
            pendingRestore: { workbook_state: wb, messages: [] },
          })
        }
      })
    }
  }, []) // eslint-disable-line react-hooks/exhaustive-deps

  // Always set role — no more role selection modal
  useEffect(() => {
    if (!useWorkspaceStore.getState().sessionRole) {
      setSessionRole('general')
    }
  }, [sessionId]) // eslint-disable-line react-hooks/exhaustive-deps

  // Persist session title and save ALL unsaved messages whenever messages change
  useEffect(() => {
    if (!sessionId || messages.length === 0) return
    const userId = user?.id

    const firstUser = messages.find((m) => m.role === 'user')
    if (firstUser) {
      const title = firstUser.text.slice(0, 60) + (firstUser.text.length > 60 ? '…' : '')
      upsertSession(sessionId, userId, title)
    }

    saveAllMessages(sessionId, userId, messages)
  }, [messages.length, sessionId]) // eslint-disable-line react-hooks/exhaustive-deps

  // Reload sessions when user changes
  useEffect(() => {
    if (user?.id) loadSessions(user.id)
  }, [user?.id]) // eslint-disable-line react-hooks/exhaustive-deps

  // Save workbook snapshot as backup whenever it changes (write-only, for recovery)
  useEffect(() => {
    const unsub = useWorkspaceStore.subscribe((state, prev) => {
      if (state.workbookState && state.workbookState !== prev.workbookState && state.sessionId) {
        saveWorkbook(state.sessionId, user?.id, state.workbookState)
      }
    })
    return unsub
  }, [user?.id]) // eslint-disable-line react-hooks/exhaustive-deps

  if (authLoading) {
    return (
      <div style={{
        display: 'flex', alignItems: 'center', justifyContent: 'center',
        height: '100vh', background: 'var(--bg-app)', color: 'var(--text-muted)'
      }}>
        Loading…
      </div>
    )
  }

  if (!user && !isGuest) {
    return <LoginPage />
  }

  const handleSelectSession = async (newSessionId) => {
    // ── 1. Switch sessionId immediately (before any async work) ──
    // This makes the WS dispatch guard reject stale messages from the old session.
    localStorage.setItem('rightcut_session_id', newSessionId)
    useWorkspaceStore.setState({
      sessionId: newSessionId,
      workbookState: null,   // cleared — backend will send via WS
      messages: [],
      tabs: [],
      activeTab: null,
      activeSheet: null,
      pendingRestore: null,
      restoring: false,
      pendingMessageId: null,
      sessionRole: 'general',
    })

    // ── 2. Load saved data (async) ──
    const { loadMessages, loadWorkbook } = useHistoryStore.getState()
    const [msgs, wb] = await Promise.all([
      loadMessages(newSessionId, user?.id),
      loadWorkbook(newSessionId, user?.id),
    ])

    // ── 3. Apply messages + set pendingRestore for WS handler ──
    // workbookState stays null — the backend will provide it via workbook_update
    // after the WS connects (if it has the workbook) or after /restore (if empty).
    useWorkspaceStore.setState({
      messages: msgs,
      pendingRestore: wb ? { workbook_state: wb, messages: msgs } : null,
      sessionRole: 'general',
    })
  }

  return (
    <div className="app-root">
      <LeftSidebar onSelectSession={handleSelectSession} />
      <div className="app-right">
        <div className="app-main">
          <div className="chat-pane">
            <ChatPanel
              sessionId={sessionId}
              onSendMessage={sendMessage}
            />
          </div>
          <div className="preview-pane">
            <PreviewPanel sessionId={sessionId} onCellEdit={sendCellEdit} />
          </div>
        </div>
        <StatusBar />
      </div>
    </div>
  )
}
