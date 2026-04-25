/**
 * RightCut — Message bubble.
 * User-facing checklist timeline, live step updates, markdown rendering.
 */
import { Bot, User, AlertCircle, LayoutGrid, FileSpreadsheet, ArrowRight } from 'lucide-react'
import ToolTimeline from './ToolTimeline'
import useWorkspaceStore from '../stores/workspaceStore'

export default function MessageBubble({ message }) {
  const toolTimelines = useWorkspaceStore((s) => s.toolTimelines)
  const documents     = useWorkspaceStore((s) => s.documents)
  const setActiveTab  = useWorkspaceStore((s) => s.setActiveTab)
  const { id, role, text, timeline, fileIds, sheetRefs, timestamp } = message

  const steps  = timeline || toolTimelines[id] || []
  const isLive = role === 'agent_pending'

  /* ── Pending (agent is working) ── */
  if (isLive) {
    return (
      <div className="message message--agent">
        <div className="agent-avatar"><Bot size={13} /></div>
        <div className="message-body">
          {steps.length === 0 ? (
            <div className="thinking-indicator">
              <span className="thinking-dot" />
              <span className="thinking-dot" />
              <span className="thinking-dot" />
            </div>
          ) : (
            <ToolTimeline steps={steps} isLive />
          )}
        </div>
      </div>
    )
  }

  /* ── System (compaction notice, etc.) ── */
  if (role === 'system') {
    return (
      <div className="message message--system">
        <div className="message-system-text">{text}</div>
      </div>
    )
  }

  /* ── Error ── */
  if (role === 'error') {
    return (
      <div className="message message--error">
        <div className="error-avatar"><AlertCircle size={13} /></div>
        <div className="message-body">
          <div className="message-text error-text">{text}</div>
          <div className="message-time">{formatTime(timestamp)}</div>
        </div>
      </div>
    )
  }

  /* ── User ── */
  if (role === 'user') {
    return (
      <div className="message message--user">
        <div className="message-body">
          <div className="message-text">
            <MarkdownText text={text} />
          </div>
          {fileIds?.length > 0 && (
            <div className="file-chips">
              {fileIds.map((fid) => {
                const doc = documents[fid]
                return (
                  <span key={fid} className="file-chip">
                    {doc?.filename || fid}
                  </span>
                )
              })}
            </div>
          )}
          <div className="message-time">{formatTime(timestamp)}</div>
        </div>
        <div className="user-avatar-msg"><User size={12} /></div>
      </div>
    )
  }

  /* ── Agent ── */
  return (
    <div className="message message--agent">
      <div className="agent-avatar"><Bot size={13} /></div>
      <div className="message-body">
        {/* Tool call timeline — shown above the text response */}
        {steps.length > 0 && <ToolTimeline steps={steps} isLive={false} />}
        <div className="message-text agent-text">
          <MarkdownText text={text} />
        </div>
        {/* File created card — shown when workbook was built this turn */}
        {sheetRefs?.length > 0 && (
          <div className="created-card">
            <div className="created-card-label">CREATED</div>
            <button
              className="created-card-file"
              onClick={() => setActiveTab(sheetRefs[0].name)}
            >
              <FileSpreadsheet size={14} className="created-card-icon" />
              <span className="created-card-name">
                {sheetRefs.map((r) => r.name).join(', ')}
              </span>
              <ArrowRight size={13} className="created-card-arrow" />
            </button>
            {sheetRefs.length > 1 && (
              <div className="created-card-sheets">
                {sheetRefs.map((ref) => (
                  <button
                    key={ref.name}
                    className="sheet-ref-pill"
                    onClick={() => setActiveTab(ref.name)}
                  >
                    <LayoutGrid size={11} />
                    {ref.name}
                  </button>
                ))}
              </div>
            )}
          </div>
        )}
        <div className="message-time">{formatTime(timestamp)}</div>
      </div>
    </div>
  )
}

/* ── Minimal markdown: **bold**, `code`, headers, bullet lists ── */
function MarkdownText({ text }) {
  if (!text) return null
  const lines = text.split('\n')
  return (
    <div className="md-body">
      {lines.map((line, i) => {
        // H3/H2/H1
        if (line.startsWith('### ')) return <div key={i} className="md-h3">{parseInline(line.slice(4))}</div>
        if (line.startsWith('## '))  return <div key={i} className="md-h2">{parseInline(line.slice(3))}</div>
        if (line.startsWith('# '))   return <div key={i} className="md-h1">{parseInline(line.slice(2))}</div>
        // Bullet list
        if (line.match(/^[-*] /)) return <div key={i} className="md-li"><span className="md-bullet">·</span>{parseInline(line.slice(2))}</div>
        // Numbered list
        if (line.match(/^\d+\. /)) return <div key={i} className="md-li"><span className="md-bullet">{line.match(/^(\d+)\./)[1]}.</span>{parseInline(line.replace(/^\d+\. /, ''))}</div>
        // Blank line = spacer
        if (!line.trim()) return <div key={i} className="md-spacer" />
        // Normal paragraph
        return <div key={i} className="md-p">{parseInline(line)}</div>
      })}
    </div>
  )
}

function parseInline(text) {
  const segments = []
  const regex = /(\*\*[^*]+\*\*|`[^`]+`)/g
  let last = 0, match, idx = 0

  while ((match = regex.exec(text)) !== null) {
    if (match.index > last) segments.push(<span key={idx++}>{text.slice(last, match.index)}</span>)
    const raw = match[0]
    if (raw.startsWith('**')) {
      segments.push(<strong key={idx++}>{raw.slice(2, -2)}</strong>)
    } else {
      segments.push(<code key={idx++} className="inline-code">{raw.slice(1, -1)}</code>)
    }
    last = match.index + raw.length
  }
  if (last < text.length) segments.push(<span key={idx++}>{text.slice(last)}</span>)
  return segments.length ? segments : [<span key={0}>{text}</span>]
}

function formatTime(ts) {
  if (!ts) return ''
  return new Date(ts).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })
}
