/**
 * RightCut — WebSocket hook.
 *
 * WORKBOOK DATA FLOW (unidirectional):
 *   1. WS connects to /ws/{sessionId}
 *   2. Backend sends session_ready { has_workbook }
 *   3a. If has_workbook: backend sends workbook_update → store renders it
 *   3b. If !has_workbook && pendingRestore: POST /restore → backend rebuilds →
 *       backend sends workbook_update on next WS message (or reconnect)
 *   4. All subsequent workbook_update messages are rendered directly
 *
 *   The frontend NEVER renders workbook data from localStorage — only from the backend.
 */
import { useEffect, useRef } from 'react'
import useWorkspaceStore from '../stores/workspaceStore'
import { apiUrl, wsUrl } from '../utils/api'

const RECONNECT_BASE_MS = 1500
const MAX_RECONNECT_ATTEMPTS = 10

export function useWebSocket() {
  const sessionId = useWorkspaceStore((s) => s.sessionId)
  const wsRef = useRef(null)
  const reconnectTimer = useRef(null)
  const attempts = useRef(0)
  const unmounted = useRef(false)

  const sendRef = useRef(null)
  const sendCellEditRef = useRef(null)

  useEffect(() => {
    if (!sessionId) return
    unmounted.current = false

    function dispatch(msg) {
      const s = useWorkspaceStore.getState()

      // Guard: reject messages from a stale WS after session switch
      if (s.sessionId !== sessionId) return

      switch (msg.type) {
        case 'session_ready': {
          // Backend already has a workbook → it will send workbook_update next.
          // Backend is empty → restore from saved snapshot if we have one.
          if (!msg.has_workbook && s.pendingRestore && !s.restoring) {
            const pendingRestore = s.pendingRestore
            useWorkspaceStore.setState({ restoring: true })
            fetch(apiUrl(`/restore/${sessionId}`), {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({
                workbook_state: pendingRestore.workbook_state,
                messages: (pendingRestore.messages || []).map((m) => ({
                  role: m.role,
                  text: m.text,
                  timestamp: m.timestamp,
                })),
                role: s.sessionRole || 'general',
              }),
            })
              .then(() => useWorkspaceStore.setState({ restoring: false, pendingRestore: null }))
              .catch((e) => {
                console.warn('Restore failed:', e)
                useWorkspaceStore.setState({ restoring: false, pendingRestore: null })
              })
          } else {
            useWorkspaceStore.setState({ pendingRestore: null })
          }
          break
        }

        case 'tool_call':
          if (s.pendingMessageId) s.addToolStep(s.pendingMessageId, msg.step)
          break

        case 'workbook_update':
          // This is the ONLY path that sets workbookState — straight from the backend.
          s.setWorkbookState(msg.state)
          s.updateSheetCount()
          break

        case 'agent_response': {
          const id = s.pendingMessageId || crypto.randomUUID()
          const MUTATING_TOOLS = new Set([
            'create_sheet', 'insert_data', 'add_formula', 'edit_cell',
            'apply_formatting', 'sort_range', 'create_model_scaffold', 'clean_data',
          ])
          const timeline = msg.timeline || []
          const builtModel = timeline.some((t) => MUTATING_TOOLS.has(t.tool))
          const sheets = s.workbookState?.sheets || []
          const sheetRefs = builtModel ? sheets.map((sh) => ({ name: sh.name })) : []
          s.addMessage({ id, role: 'agent', text: msg.text, timeline, sheetRefs, timestamp: Date.now() })
          s.setPendingMessageId(null)
          s.setWsStatus('connected')
          break
        }

        case 'new_tab':
          s.addTab(msg.tab)
          break

        case 'thinking':
          s.setWsStatus('thinking')
          break

        case 'error':
          s.addMessage({ id: crypto.randomUUID(), role: 'error', text: msg.message, timestamp: Date.now() })
          s.setWsStatus('connected')
          s.setPendingMessageId(null)
          break

        case 'history_compacted': {
          const strategyLabels = {
            tool_result: 'Tool results compressed',
            summarization: 'Older turns summarized',
            sliding_window: 'History trimmed',
          }
          const label = strategyLabels[msg.strategy] || 'History compacted'
          const saving = msg.tokens_before && msg.tokens_after
            ? ` · ${Math.round((1 - msg.tokens_after / msg.tokens_before) * 100)}% token saving`
            : ''
          s.addMessage({
            id: crypto.randomUUID(),
            role: 'system',
            text: `${label}${saving}`,
            timestamp: Date.now(),
          })
          break
        }

        case 'pong':
          break
        default:
          break
      }
    }

    function connect() {
      if (unmounted.current) return
      if (wsRef.current && wsRef.current.readyState < 2) {
        wsRef.current.onclose = null
        wsRef.current.close()
      }

      const ws = new WebSocket(wsUrl(sessionId))
      wsRef.current = ws
      useWorkspaceStore.getState().setWsStatus('connecting')

      ws.onopen = () => {
        attempts.current = 0
        useWorkspaceStore.getState().setWsStatus('connected')
      }

      ws.onmessage = (e) => {
        try { dispatch(JSON.parse(e.data)) } catch (_) {}
      }

      ws.onclose = () => {
        if (unmounted.current) return
        const st = useWorkspaceStore.getState()
        st.setWsStatus('disconnected')
        if (st.pendingMessageId) {
          st.addMessage({ id: st.pendingMessageId, role: 'agent', text: '*(Connection lost — please resend your message)*', timestamp: Date.now() })
          st.setPendingMessageId(null)
        }
        if (attempts.current < MAX_RECONNECT_ATTEMPTS) {
          const delay = Math.min(RECONNECT_BASE_MS * Math.pow(2, attempts.current), 30_000)
          attempts.current++
          reconnectTimer.current = setTimeout(connect, delay)
        }
      }

      ws.onerror = () => ws.close()
    }

    connect()

    const ping = setInterval(() => {
      if (wsRef.current?.readyState === WebSocket.OPEN) {
        wsRef.current.send(JSON.stringify({ type: 'ping' }))
      }
    }, 25_000)

    return () => {
      unmounted.current = true
      clearTimeout(reconnectTimer.current)
      clearInterval(ping)
      wsRef.current?.close()
    }
  }, [sessionId])

  // ── Public API ────────────────────────────────────────────────────────────

  sendRef.current = (text, fileIds = []) => {
    const ws = wsRef.current
    const s = useWorkspaceStore.getState()

    if (!ws || ws.readyState !== WebSocket.OPEN) {
      s.addMessage({ id: crypto.randomUUID(), role: 'error', text: 'Not connected. Retrying...', timestamp: Date.now() })
      return
    }

    if (s.restoring) {
      s.addMessage({ id: crypto.randomUUID(), role: 'error', text: 'Restoring session, please wait a moment...', timestamp: Date.now() })
      return
    }

    s.addMessage({ id: crypto.randomUUID(), role: 'user', text, fileIds, timestamp: Date.now() })

    const pendingId = crypto.randomUUID()
    s.setPendingMessageId(pendingId)
    s.addMessage({ id: pendingId, role: 'agent_pending', text: '', timestamp: Date.now() })
    s.setWsStatus('thinking')
    s.clearPendingFiles()

    const files = fileIds.map((fid) => {
      const doc = s.documents[fid]
      return { file_id: fid, filename: doc?.filename || fid }
    })

    ws.send(JSON.stringify({ type: 'user_message', text, files }))
  }

  sendCellEditRef.current = (sheet, cell, oldVal, newVal) => {
    const ws = wsRef.current
    if (ws?.readyState === WebSocket.OPEN) {
      ws.send(JSON.stringify({ type: 'cell_edit', sheet, cell, old: oldVal, new: newVal }))
    }
  }

  return {
    sendMessage:  (...args) => sendRef.current?.(...args),
    sendCellEdit: (...args) => sendCellEditRef.current?.(...args),
  }
}
