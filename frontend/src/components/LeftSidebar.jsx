/**
 * RightCut — Persistent left sidebar.
 * Replaces the history drawer and top ribbon.
 */
import { useEffect, useState } from 'react'
import {
  SquarePen, FolderOpen, LayoutTemplate, Archive,
  ChevronDown, ChevronRight, Sun, Moon, LogOut,
  LayoutGrid, FileText, BarChart3, Trash2
} from 'lucide-react'
import useHistoryStore from '../stores/historyStore'
import useAuthStore from '../stores/authStore'
import useThemeStore from '../stores/themeStore'
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

export default function LeftSidebar({ onSelectSession }) {
  const { user, isGuest, signOut } = useAuthStore()
  const { theme, toggleTheme } = useThemeStore()
  const { sessions, loadSessions, deleteSession } = useHistoryStore()
  const currentSessionId = useWorkspaceStore((s) => s.sessionId)
  const tabs = useWorkspaceStore((s) => s.tabs)
  const setActiveTab = useWorkspaceStore((s) => s.setActiveTab)
  const [expandedSession, setExpandedSession] = useState(null)
  const [deletingId, setDeletingId] = useState(null)

  useEffect(() => {
    loadSessions(user?.id)
  }, [user?.id]) // eslint-disable-line

  const handleDeleteSession = async (e, sessionId) => {
    e.stopPropagation()
    if (!window.confirm('Delete this session? This cannot be undone.')) return
    setDeletingId(sessionId)
    await deleteSession(sessionId, user?.id)
    setDeletingId(null)
    // If we deleted the active session, start a new one
    if (sessionId === currentSessionId) {
      const newId = crypto.randomUUID()
      localStorage.setItem('rightcut_session_id', newId)
      useWorkspaceStore.setState({
        sessionId: newId, messages: [], workbookState: null,
        tabs: [], activeTab: null, activeSheet: null, sessionRole: null,
        pendingRestore: null, restoring: false, pendingMessageId: null,
      })
    }
  }

  const handleNewSession = () => {
    const newId = crypto.randomUUID()
    localStorage.setItem('rightcut_session_id', newId)
    useWorkspaceStore.setState({
      sessionId: newId,
      messages: [],
      workbookState: null,
      tabs: [],
      activeTab: null,
      activeSheet: null,
      sessionRole: 'general',
      pendingRestore: null,
      restoring: false,
      pendingMessageId: null,
    })
  }

  // Get sheet tabs for a given session (only available for current session)
  const getSessionTabs = (sessionId) => {
    if (sessionId !== currentSessionId) return []
    return tabs.filter((t) => t.type === 'sheet')
  }

  const userDisplay = isGuest ? 'Guest' : (user?.email?.split('@')[0] || 'User')
  const userInitials = isGuest ? 'G' : (user?.email?.slice(0, 2).toUpperCase() || 'U')

  return (
    <aside className="left-sidebar">
      {/* Logo */}
      <div className="sidebar-logo">
        <div className="sidebar-logo-mark">
          <BarChart3 size={14} color="#fff" />
        </div>
        <span className="sidebar-logo-text">RightCut</span>
      </div>

      {/* New Analysis */}
      <div className="sidebar-section sidebar-section--top">
        <button className="sidebar-new-btn" onClick={handleNewSession}>
          <SquarePen size={14} />
          New Analysis
        </button>
      </div>

      {/* Nav items */}
      <nav className="sidebar-nav">
        <button className="sidebar-nav-item sidebar-nav-item--active">
          <FolderOpen size={14} />
          Projects
        </button>
        <button className="sidebar-nav-item">
          <LayoutTemplate size={14} />
          Templates
        </button>
        <button className="sidebar-nav-item">
          <Archive size={14} />
          Vault
        </button>
      </nav>

      <div className="sidebar-divider" />

      {/* Recent sessions */}
      <div className="sidebar-section-label">Recent</div>
      <div className="sidebar-sessions">
        {sessions.length === 0 && (
          <div className="sidebar-empty">No sessions yet</div>
        )}
        {sessions.map((session) => {
          const isActive = session.session_id === currentSessionId
          const sessionTabs = getSessionTabs(session.session_id)
          const isExpanded = expandedSession === session.session_id || isActive

          const handleSessionClick = () => {
            if (!isActive) {
              // Switch to session — expand happens naturally via isActive becoming true
              onSelectSession(session.session_id)
              setExpandedSession(session.session_id)
            } else {
              // Already active — just toggle sheet list expand
              setExpandedSession(isExpanded ? null : session.session_id)
            }
          }

          return (
            <div key={session.session_id} className="sidebar-session-group">
              <div className="sidebar-session-row">
                <button
                  className={`sidebar-session-item ${isActive ? 'sidebar-session-item--active' : ''}`}
                  onClick={handleSessionClick}
                >
                  <span className="sidebar-session-chevron">
                    {isExpanded && sessionTabs.length > 0
                      ? <ChevronDown size={11} />
                      : <ChevronRight size={11} />
                    }
                  </span>
                  <span className="sidebar-session-title">
                    {session.title || 'Untitled'}
                  </span>
                  {isActive && <span className="sidebar-session-dot" />}
                </button>
                <button
                  className="sidebar-session-delete"
                  onClick={(e) => handleDeleteSession(e, session.session_id)}
                  disabled={deletingId === session.session_id}
                  title="Delete session"
                >
                  <Trash2 size={11} />
                </button>
              </div>

              {/* Sheet sub-items */}
              {isExpanded && sessionTabs.length > 0 && (
                <div className="sidebar-sub-items">
                  {sessionTabs.map((tab) => (
                    <button
                      key={tab.id}
                      className="sidebar-sub-item"
                      onClick={() => setActiveTab(tab.id)}
                    >
                      {tab.type === 'sheet'
                        ? <LayoutGrid size={11} />
                        : <FileText size={11} />
                      }
                      <span>{tab.name}</span>
                    </button>
                  ))}
                </div>
              )}
            </div>
          )
        })}
      </div>

      {/* Bottom: user + theme */}
      <div className="sidebar-bottom">
        <div className="sidebar-divider" />
        <button className="sidebar-theme-btn" onClick={toggleTheme}>
          {theme === 'light' ? <Moon size={13} /> : <Sun size={13} />}
          {theme === 'light' ? 'Dark mode' : 'Light mode'}
        </button>
        <button className="sidebar-user-btn" onClick={signOut} title="Sign out">
          <div className="sidebar-user-avatar">{userInitials}</div>
          <span className="sidebar-user-name">{userDisplay}</span>
          <LogOut size={11} style={{ opacity: 0.5, marginLeft: 'auto' }} />
        </button>
      </div>
    </aside>
  )
}
