/**
 * RightCut — Chat panel with @ mention support.
 */
import { useCallback, useEffect, useRef, useState } from 'react'
import {
  Send, Paperclip,
  FileText, FileSpreadsheet, FileBarChart2, AtSign,
  Building2, TrendingUp, Calculator, Download
} from 'lucide-react'
import MessageBubble from './MessageBubble'
import useWorkspaceStore from '../stores/workspaceStore'

function fileIcon(fileType) {
  if (!fileType) return <FileText size={13} />
  const t = fileType.toLowerCase()
  if (t === 'pdf') return <FileText size={13} />
  if (t === 'csv' || t === 'xlsx' || t === 'xls') return <FileSpreadsheet size={13} />
  return <FileBarChart2 size={13} />
}

export default function ChatPanel({ sessionId, onSendMessage }) {
  const [inputText, setInputText]       = useState('')
  const [mentionQuery, setMentionQuery] = useState(null)  // null = closed, string = open
  const [mentionIdx, setMentionIdx]     = useState(0)
  const [historyIdx, setHistoryIdx]     = useState(-1)   // -1 = not browsing history
  const [savedDraft, setSavedDraft]     = useState('')   // draft saved when entering history
  const messagesEndRef = useRef(null)
  const fileInputRef   = useRef(null)
  const textareaRef    = useRef(null)

  const messages         = useWorkspaceStore((s) => s.messages)
  const wsStatus         = useWorkspaceStore((s) => s.wsStatus)
  const pendingFileIds   = useWorkspaceStore((s) => s.pendingFileIds)
  const documents        = useWorkspaceStore((s) => s.documents)
  const stagePendingFile = useWorkspaceStore((s) => s.stagePendingFile)
  const addDocument      = useWorkspaceStore((s) => s.addDocument)
  const addTab           = useWorkspaceStore((s) => s.addTab)

  const isThinking = wsStatus === 'thinking'
  const canSend    = inputText.trim().length > 0 && !isThinking && wsStatus === 'connected'

  const handleDownloadSummary = async (format = 'md') => {
    if (!sessionId) return
    try {
      const res = await fetch(`/summary/${sessionId}?format=${format}`)
      if (!res.ok) {
        const err = await res.json().catch(() => ({}))
        alert(err.detail || 'Failed to generate summary.')
        return
      }
      const blob = await res.blob()
      const url = URL.createObjectURL(blob)
      const a = document.createElement('a')
      a.href = url
      a.download = `rightcut_summary.${format}`
      a.click()
      URL.revokeObjectURL(url)
    } catch {
      alert('Failed to generate summary.')
    }
  }

  // Filter docs for @ mention
  const docList = Object.values(documents)
  const filteredDocs = mentionQuery !== null
    ? docList.filter((d) =>
        d.filename.toLowerCase().includes((mentionQuery || '').toLowerCase())
      )
    : []

  // Auto-scroll
  useEffect(() => {
    messagesEndRef.current?.scrollIntoView({ behavior: 'smooth' })
  }, [messages])

  // Auto-resize textarea
  useEffect(() => {
    const ta = textareaRef.current
    if (!ta) return
    ta.style.height = 'auto'
    ta.style.height = `${Math.min(ta.scrollHeight, 130)}px`
  }, [inputText])

  // Ordered list of user messages for history navigation (newest last)
  const userMessages = messages.filter((m) => m.role === 'user').map((m) => m.text)

  const handleSend = useCallback(() => {
    const text = inputText.trim()
    if (!text || isThinking) return
    onSendMessage(text, pendingFileIds)
    setInputText('')
    setHistoryIdx(-1)
    setSavedDraft('')
  }, [inputText, isThinking, onSendMessage, pendingFileIds])

  const handleKeyDown = (e) => {
    // Handle @ mention navigation first
    if (mentionQuery !== null && filteredDocs.length > 0) {
      if (e.key === 'ArrowDown') {
        e.preventDefault()
        setMentionIdx((i) => (i + 1) % filteredDocs.length)
        return
      }
      if (e.key === 'ArrowUp') {
        e.preventDefault()
        setMentionIdx((i) => (i - 1 + filteredDocs.length) % filteredDocs.length)
        return
      }
      if (e.key === 'Enter' || e.key === 'Tab') {
        e.preventDefault()
        insertMention(filteredDocs[mentionIdx])
        return
      }
      if (e.key === 'Escape') {
        setMentionQuery(null)
        return
      }
    }

    // Message history navigation — only when cursor is at start of empty/single-line input
    if (e.key === 'ArrowUp' && userMessages.length > 0) {
      const ta = textareaRef.current
      const atTop = !ta || ta.selectionStart === 0
      if (atTop) {
        e.preventDefault()
        const newIdx = historyIdx === -1 ? userMessages.length - 1 : Math.max(0, historyIdx - 1)
        if (historyIdx === -1) setSavedDraft(inputText)
        setHistoryIdx(newIdx)
        setInputText(userMessages[newIdx])
        return
      }
    }
    if (e.key === 'ArrowDown' && historyIdx !== -1) {
      e.preventDefault()
      if (historyIdx >= userMessages.length - 1) {
        // Back to draft
        setHistoryIdx(-1)
        setInputText(savedDraft)
      } else {
        const newIdx = historyIdx + 1
        setHistoryIdx(newIdx)
        setInputText(userMessages[newIdx])
      }
      return
    }

    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault()
      handleSend()
    }
  }

  const handleInput = (e) => {
    const val = e.target.value
    setInputText(val)
    // Typing while in history mode exits it
    if (historyIdx !== -1) {
      setHistoryIdx(-1)
      setSavedDraft('')
    }

    // Detect @ trigger
    const cursor = e.target.selectionStart
    const textBeforeCursor = val.slice(0, cursor)
    const atMatch = textBeforeCursor.match(/@(\w*)$/)
    if (atMatch && docList.length > 0) {
      setMentionQuery(atMatch[1])
      setMentionIdx(0)
    } else {
      setMentionQuery(null)
    }
  }

  const insertMention = (doc) => {
    const ta = textareaRef.current
    if (!ta) return
    const cursor = ta.selectionStart
    const before = inputText.slice(0, cursor)
    const after  = inputText.slice(cursor)
    // Replace the @query with @filename
    const replaced = before.replace(/@(\w*)$/, `@${doc.filename} `)
    setInputText(replaced + after)
    setMentionQuery(null)
    setTimeout(() => {
      ta.focus()
      ta.setSelectionRange(replaced.length, replaced.length)
    }, 0)
  }

  const handleFileChange = async (e) => {
    const files = Array.from(e.target.files || [])
    e.target.value = ''
    for (const file of files) {
      try {
        const formData = new FormData()
        formData.append('file', file)
        const res = await fetch(`/upload/${sessionId}`, { method: 'POST', body: formData })
        if (!res.ok) {
          const err = await res.json().catch(() => ({ detail: 'Upload failed' }))
          useWorkspaceStore.getState().addMessage({
            id: crypto.randomUUID(), role: 'error',
            text: `Upload failed: ${err.detail || res.statusText}`, timestamp: Date.now(),
          })
          continue
        }
        const data = await res.json()
        const doc = {
          file_id: data.file_id, filename: data.filename,
          file_type: data.file_type, page_count: data.page_count,
          content: '', tables: [],
        }
        addDocument(data.file_id, doc)
        stagePendingFile(data.file_id)
        addTab({ id: data.file_id, name: data.filename, type: 'document' })
        useWorkspaceStore.getState().addMessage({
          id: crypto.randomUUID(), role: 'agent', timeline: [], timestamp: Date.now(),
          text: `Uploaded **${data.filename}** (${data.file_type.toUpperCase()}${data.page_count ? `, ${data.page_count} pages` : ''}). Type \`@${data.filename}\` to reference it.`,
        })
      } catch (err) {
        useWorkspaceStore.getState().addMessage({
          id: crypto.randomUUID(), role: 'error',
          text: `Upload error: ${err.message}`, timestamp: Date.now(),
        })
      }
    }
  }

  return (
    <div className="chat-panel">
      {/* Messages */}
      <div className="messages-area" role="log" aria-live="polite">
        {messages.length === 0 && (
          <EmptyChat onPromptClick={(text) => { setInputText(text); textareaRef.current?.focus() }} />
        )}
        {messages.map((msg) => (
          <MessageBubble key={msg.id} message={msg} />
        ))}
        <div ref={messagesEndRef} />
      </div>

      {/* Staged files */}
      {pendingFileIds.length > 0 && (
        <div className="staged-files">
          {pendingFileIds.map((fid) => {
            const doc = documents[fid]
            return (
              <span key={fid} className="staged-chip">
                {fileIcon(doc?.file_type)}
                {doc?.filename || fid}
              </span>
            )
          })}
        </div>
      )}

      {/* Input */}
      <div className="input-wrapper">
        {/* @ mention dropdown */}
        {mentionQuery !== null && filteredDocs.length > 0 && (
          <div className="mention-dropdown" role="listbox">
            <div className="mention-header">Documents</div>
            {filteredDocs.map((doc, i) => (
              <button
                key={doc.file_id}
                className={`mention-item ${i === mentionIdx ? 'selected' : ''}`}
                onMouseDown={(e) => { e.preventDefault(); insertMention(doc) }}
                role="option"
                aria-selected={i === mentionIdx}
              >
                <span className="mention-item-icon">{fileIcon(doc.file_type)}</span>
                <span className="mention-item-name">{doc.filename}</span>
                <span className="mention-item-type">{doc.file_type?.toUpperCase()}</span>
              </button>
            ))}
          </div>
        )}

        <div className="input-row">
          <textarea
            ref={textareaRef}
            className="chat-textarea"
            value={inputText}
            onChange={handleInput}
            onKeyDown={handleKeyDown}
            placeholder={
              isThinking ? 'Agent is working…'
              : wsStatus === 'connected' ? 'Ask anything, or @ to mention a file'
              : 'Connecting…'
            }
            disabled={isThinking || wsStatus !== 'connected'}
            rows={1}
            aria-label="Chat input"
          />
          <div className="input-actions">
            <button
              className="input-icon-btn"
              onClick={() => fileInputRef.current?.click()}
              title="Upload document (PDF, DOCX, CSV, XLSX)"
              disabled={isThinking}
              aria-label="Attach file"
            >
              <Paperclip size={15} />
            </button>
            <button
              className="input-icon-btn"
              onClick={() => {
                setInputText((t) => t + '@')
                textareaRef.current?.focus()
                setMentionQuery('')
              }}
              title="Mention a document"
              disabled={isThinking || docList.length === 0}
              aria-label="Mention document"
            >
              <AtSign size={15} />
            </button>
            <button
              className="send-btn"
              onClick={handleSend}
              disabled={!canSend}
              title="Send (Enter)"
              aria-label="Send message"
            >
              {isThinking
                ? <span className="send-spinner" />
                : <Send size={14} />
              }
            </button>
          </div>
        </div>
        <input
          type="file"
          ref={fileInputRef}
          onChange={handleFileChange}
          accept=".pdf,.docx,.doc,.csv,.xlsx,.xls"
          multiple
          hidden
        />
        <div className="input-hint">
          <span className="input-hint-text">@ to mention · ↑↓ history</span>
          <div className="input-hint-right">
            {messages.length > 1 && (
              <div className="summary-dropdown">
                <button
                  className="input-hint-summary"
                  onClick={() => handleDownloadSummary('md')}
                  title="Download session summary"
                >
                  <Download size={11} />
                  Summary
                </button>
                <button
                  className="input-hint-summary input-hint-summary--txt"
                  onClick={() => handleDownloadSummary('txt')}
                  title="Download as .txt"
                >
                  .txt
                </button>
              </div>
            )}
            <span className="input-hint-kbd">↵ Send</span>
          </div>
        </div>
      </div>
    </div>
  )
}


const QUICK_PROMPTS = [
  { icon: Building2, text: 'Build a DCF model for a $50M EBITDA company' },
  { icon: TrendingUp, text: 'Create a SaaS comps table with EV/Revenue' },
  { icon: Calculator, text: 'Build an LBO model with MOIC and IRR' },
]

function EmptyChat({ onPromptClick }) {
  return (
    <div className="empty-chat">
      <div className="empty-chat-icon">
        <FileBarChart2 size={24} />
      </div>
      <div className="empty-chat-title">Start your analysis</div>
      <div className="empty-chat-text">
        Ask RightCut to build a model, or upload a document and use @ to reference it.
      </div>
      <div className="empty-chat-prompts">
        {QUICK_PROMPTS.map(({ icon: Icon, text }, i) => (
          <button
            key={i}
            className="empty-chat-prompt"
            onClick={() => onPromptClick?.(text)}
          >
            <Icon size={12} />
            {text}
          </button>
        ))}
      </div>
    </div>
  )
}
