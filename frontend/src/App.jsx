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
  const { upsertSession, saveMessage, loadSessions, saveWorkbook, loadWorkbook } = useHistoryStore()

  // Single WebSocket for the entire app
  const { sendMessage, sendCellEdit } = useWebSocket()

  // Init auth + theme on mount
  useEffect(() => {
    initAuth()
    initTheme()
  }, []) // eslint-disable-line react-hooks/exhaustive-deps

  // If messages load into a session that has no role set, default to finance
  // (covers restored sessions where the modal would incorrectly appear)
  useEffect(() => {
    if (sessionRole === null && messages.length > 0) {
      setSessionRole('finance')
    }
  }, [messages.length]) // eslint-disable-line react-hooks/exhaustive-deps

  // Persist session to history whenever messages change
  useEffect(() => {
    if (!sessionId || messages.length === 0) return
    const userId = user?.id
    const firstUser = messages.find((m) => m.role === 'user')
    const title = firstUser
      ? firstUser.text.slice(0, 60) + (firstUser.text.length > 60 ? '…' : '')
      : 'Session ' + sessionId.slice(0, 8)
    upsertSession(sessionId, userId, title)
    const last = messages[messages.length - 1]
    if (last && (last.role === 'user' || last.role === 'agent')) {
      saveMessage(sessionId, userId, last)
    }
  }, [messages]) // eslint-disable-line react-hooks/exhaustive-deps

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
    localStorage.setItem('rightcut_session_id', newSessionId)
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

    // Switch frontend state first so the UI updates immediately
    useWorkspaceStore.setState({
      sessionId: newSessionId,
      messages: msgs,
      workbookState: wb || null,
      tabs,
      activeTab: activeSheet,
      activeSheet,
      // Mark that this session needs backend restoration (WS onopen will pick this up)
      pendingRestore: wb ? { workbook_state: wb, messages: msgs } : null,
      // Existing sessions skip the role modal — treat as finance (default)
      sessionRole: msgs.length > 0 ? 'finance' : null,
    })

    // Also eagerly call restore so backend is ready before the next user message
    if (wb) {
      try {
        await fetch(`/restore/${newSessionId}`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            workbook_state: wb,
            messages: msgs.map((m) => ({ role: m.role, text: m.text, timestamp: m.timestamp })),
          }),
        })
      } catch (e) {
        // Non-fatal: WS session_ready handler will retry if needed
        console.warn('Eager restore failed, will retry on WS connect:', e)
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
