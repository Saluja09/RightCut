/**
 * RightCut — Root application component.
 * Sidebar + chat + preview layout.
 */
import { useEffect } from 'react'
import ChatPanel from './components/ChatPanel'
import PreviewPanel from './components/PreviewPanel'
import StatusBar from './components/StatusBar'
import LoginPage from './components/LoginPage'
import LeftSidebar from './components/LeftSidebar'
import RoleSelectModal from './components/RoleSelectModal'
import useWorkspaceStore from './stores/workspaceStore'
import useAuthStore from './stores/authStore'
import useThemeStore from './stores/themeStore'
import useHistoryStore from './stores/historyStore'
import { useWebSocket } from './hooks/useWebSocket'
import { apiUrl } from './utils/api'

// Initialize session synchronously so sessionId exists before any hook runs
const store = useWorkspaceStore.getState()
if (!store.sessionId) {
  store.initSession()
}

export default function App() {
  const sessionId   = useWorkspaceStore((s) => s.sessionId)
  const messages    = useWorkspaceStore((s) => s.messages)
  const sessionRole = useWorkspaceStore((s) => s.sessionRole)
  const setSessionRole = useWorkspaceStore((s) => s.setSessionRole)
  const { user, isGuest, loading: authLoading, initAuth } = useAuthStore()
  const { initTheme } = useThemeStore()
  const { upsertSession, saveAllMessages, loadSessions, saveWorkbook, loadWorkbook } = useHistoryStore()

  // Single WebSocket for the entire app
  const { sendMessage, sendCellEdit } = useWebSocket()

  // Init auth + theme on mount, and reload messages for the current active session
  useEffect(() => {
    initAuth()
    initTheme()
    // Restore messages for the current session on page load
    // (workspaceStore persists workbookState but not messages)
    if (sessionId) {
      const hs = useHistoryStore.getState()
      hs.loadMessages(sessionId, undefined).then((msgs) => {
        if (msgs.length > 0 && useWorkspaceStore.getState().messages.length === 0) {
          useWorkspaceStore.setState({ messages: msgs })
        }
      })
      // Also ensure the current session's workbook is saved to per-session key
      // so that switching away and back restores it correctly
      const wb = useWorkspaceStore.getState().workbookState
      if (wb) {
        hs.saveWorkbook(sessionId, undefined, wb)
      }
    }
  }, []) // eslint-disable-line react-hooks/exhaustive-deps

  // If messages load into a session that has no role set, default to general
  // (covers restored sessions where the modal would incorrectly appear)
  useEffect(() => {
    if (sessionRole === null && messages.length > 0) {
      setSessionRole('general')
    }
  }, [messages.length]) // eslint-disable-line react-hooks/exhaustive-deps

  // Persist session title and save ALL unsaved messages whenever messages change
  useEffect(() => {
    if (!sessionId || messages.length === 0) return
    const userId = user?.id

    // Title: derived from first user message
    const firstUser = messages.find((m) => m.role === 'user')
    if (firstUser) {
      const title = firstUser.text.slice(0, 60) + (firstUser.text.length > 60 ? '…' : '')
      upsertSession(sessionId, userId, title)
    }

    // Batch-save all user/agent messages — deduplicates by id
    saveAllMessages(sessionId, userId, messages)
  }, [messages.length, sessionId]) // eslint-disable-line react-hooks/exhaustive-deps

  // Reload sessions when user changes
  useEffect(() => {
    if (user?.id) loadSessions(user.id)
  }, [user?.id]) // eslint-disable-line react-hooks/exhaustive-deps

  // Save workbook snapshot whenever it changes
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
    // Save the current session's workbook before switching so it can be restored later
    const cur = useWorkspaceStore.getState()
    if (cur.sessionId && cur.workbookState) {
      try {
        localStorage.setItem(`rightcut_wb_${cur.sessionId}`, JSON.stringify({ ...cur.workbookState, _sessionId: cur.sessionId }))
      } catch (_) {}
    }

    // ── CRITICAL: switch sessionId + clear workbook BEFORE any async work ──
    // This ensures the WS dispatch guard (`s.sessionId !== sessionId`) rejects
    // stale workbook_update messages from the old session's WebSocket while we
    // await loadMessages / loadWorkbook below.
    localStorage.setItem('rightcut_session_id', newSessionId)
    useWorkspaceStore.setState({
      sessionId: newSessionId,
      workbookState: null,
      tabs: [],
      activeTab: null,
      activeSheet: null,
      pendingRestore: null,
      restoring: false,
      pendingMessageId: null,
    })

    const { loadMessages, loadWorkbook } = useHistoryStore.getState()
    const [msgs, wb] = await Promise.all([
      loadMessages(newSessionId, user?.id),
      loadWorkbook(newSessionId, user?.id),
    ])
    const tabs = []
    if (wb?.sheets) {
      for (const sheet of wb.sheets) {
        tabs.push({ id: sheet.name, name: sheet.name, type: 'sheet' })
      }
    }
    const activeSheet = wb?.active_sheet || wb?.sheets?.[0]?.name || null

    // Apply loaded data — sessionId is already set above
    useWorkspaceStore.setState({
      messages: msgs,
      workbookState: wb || null,
      tabs,
      activeTab: activeSheet,
      activeSheet,
      // Mark that this session needs backend restoration (WS onopen will pick this up)
      pendingRestore: wb ? { workbook_state: wb, messages: msgs } : null,
      // Block message sends until restore is done
      restoring: !!wb,
      // Existing sessions skip the role modal — default to general
      sessionRole: msgs.length > 0 ? 'general' : null,
    })

    // Eagerly call restore so backend is ready before the next user message
    if (wb) {
      try {
        await fetch(apiUrl(`/restore/${newSessionId}`), {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            workbook_state: wb,
            messages: msgs.map((m) => ({ role: m.role, text: m.text, timestamp: m.timestamp })),
            role: useWorkspaceStore.getState().sessionRole || 'general',
          }),
        })
      } catch (e) {
        console.warn('Eager restore failed, will retry on WS connect:', e)
      } finally {
        useWorkspaceStore.setState({ restoring: false })
      }
    }
  }

  const showRoleModal = sessionRole === null && sessionId

  const handleRoleConfirm = (role) => {
    setSessionRole(role)
  }

  return (
    <div className="app-root">
      {showRoleModal && (
        <RoleSelectModal
          sessionId={sessionId}
          onConfirm={handleRoleConfirm}
        />
      )}
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
