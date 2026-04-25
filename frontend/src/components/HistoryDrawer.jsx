import { useEffect } from 'react'
import { X, Plus, BarChart3, Clock } from 'lucide-react'
import useHistoryStore from '../stores/historyStore'
import useAuthStore from '../stores/authStore'
import useWorkspaceStore from '../stores/workspaceStore'

function formatDate(iso) {
  if (!iso) return ''
  const d = new Date(iso)
  const now = new Date()
  const diffDays = Math.floor((now - d) / 86400000)
  if (diffDays === 0) return 'Today'
  if (diffDays === 1) return 'Yesterday'
  if (diffDays < 7) return `${diffDays}d ago`
  return d.toLocaleDateString('en-US', { month: 'short', day: 'numeric' })
}

export default function HistoryDrawer({ onClose, onSelectSession }) {
  const { user } = useAuthStore()
  const { sessions, loading, loadSessions } = useHistoryStore()
  const currentSessionId = useWorkspaceStore((s) => s.sessionId)

  useEffect(() => {
    loadSessions(user?.id)
  }, [user?.id]) // eslint-disable-line

  const handleNewSession = () => {
    const cur = useWorkspaceStore.getState()
    if (cur.sessionId && cur.workbookState) {
      try {
        localStorage.setItem(`rightcut_wb_${cur.sessionId}`, JSON.stringify(cur.workbookState))
      } catch (_) {}
    }
    const newId = crypto.randomUUID()
    localStorage.setItem('rightcut_session_id', newId)
    useWorkspaceStore.setState({
      sessionId: newId,
      messages: [],
      workbookState: null,
      tabs: [],
      activeTab: null,
      activeSheet: null,
    })
    onClose()
  }

  return (
    <div className="history-overlay" onClick={onClose} role="dialog" aria-modal="true" aria-label="Session history">
      <div className="history-drawer" onClick={(e) => e.stopPropagation()}>
        <div className="history-header">
          <span className="history-title">Session History</span>
          <button className="close-btn" onClick={onClose} aria-label="Close history">
            <X size={15} />
          </button>
        </div>

        <button className="history-new-btn" onClick={handleNewSession}>
          <Plus size={14} />
          New session
        </button>

        <div className="history-list" role="list">
          {loading && (
            <div className="history-empty">
              <Clock size={22} style={{ opacity: 0.4 }} />
              Loading…
            </div>
          )}

          {!loading && sessions.length === 0 && (
            <div className="history-empty">
              <BarChart3 size={24} style={{ opacity: 0.3 }} />
              No past sessions yet.
              <span>Start chatting to create one.</span>
            </div>
          )}

          {sessions.map((session) => {
            const isActive = session.session_id === currentSessionId
            return (
              <button
                key={session.id || session.session_id}
                className={`history-item ${isActive ? 'history-item--active' : ''}`}
                onClick={() => { onSelectSession(session.session_id); onClose() }}
                role="listitem"
              >
                <div className="history-item-ico">
                  <BarChart3 size={13} />
                </div>
                <div className="history-item-body">
                  <div className="history-item-title">
                    {session.title || 'Untitled session'}
                  </div>
                  <div className="history-item-date">
                    {formatDate(session.updated_at || session.created_at)}
                  </div>
                </div>
              </button>
            )
          })}
        </div>
      </div>
    </div>
  )
}
