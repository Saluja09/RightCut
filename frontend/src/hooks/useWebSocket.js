/**
 * RightCut — WebSocket hook.
 * Simplified: single useEffect owns the full WS lifecycle.
 * No useCallback chains — avoids stale closure / infinite-reconnect bugs.
 */
import { useEffect, useRef } from 'react'
import useWorkspaceStore from '../stores/workspaceStore'

const RECONNECT_BASE_MS = 1500
const MAX_RECONNECT_ATTEMPTS = 10

export function useWebSocket() {
  const sessionId = useWorkspaceStore((s) => s.sessionId)
  const wsRef = useRef(null)
  const reconnectTimer = useRef(null)
  const attempts = useRef(0)
  const unmounted = useRef(false)

  // Stable send helpers exposed via refs so other hooks can call them
  const sendRef = useRef(null)
  const sendCellEditRef = useRef(null)

  useEffect(() => {
    if (!sessionId) return
    unmounted.current = false

    function dispatch(msg) {
      const s = useWorkspaceStore.getState()
      switch (msg.type) {
        case 'tool_call':
          if (s.pendingMessageId) s.addToolStep(s.pendingMessageId, msg.step)
          break
        case 'workbook_update':
          s.setWorkbookState(msg.state)
          s.updateSheetCount()
          break
        case 'agent_response': {
          const id = s.pendingMessageId || crypto.randomUUID()
          // Attach sheet references from current workbook so user can reopen them
          const sheets = s.workbookState?.sheets || []
          const sheetRefs = sheets.map((sh) => ({ name: sh.name }))
          s.addMessage({ id, role: 'agent', text: msg.text, timeline: msg.timeline || [], sheetRefs, timestamp: Date.now() })
          // Remove the pending placeholder
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
        case 'pong':
          break
        default:
          break
      }
    }

    function connect() {
      if (unmounted.current) return
      // Close any existing connection before opening a new one
      if (wsRef.current && wsRef.current.readyState < 2) {
        wsRef.current.onclose = null  // prevent reconnect loop
        wsRef.current.close()
      }

      const ws = new WebSocket(`ws://localhost:8000/ws/${sessionId}`)
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
        useWorkspaceStore.getState().setWsStatus('disconnected')
        if (attempts.current < MAX_RECONNECT_ATTEMPTS) {
          const delay = Math.min(RECONNECT_BASE_MS * Math.pow(2, attempts.current), 30_000)
          attempts.current++
          reconnectTimer.current = setTimeout(connect, delay)
        }
      }

      ws.onerror = () => ws.close()
    }

    connect()

    // Heartbeat
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
  }, [sessionId])  // reconnect only when sessionId changes

  // ── Public API (stable refs, safe to call from anywhere) ─────────────────

  sendRef.current = (text, fileIds = []) => {
    const ws = wsRef.current
    const s = useWorkspaceStore.getState()

    if (!ws || ws.readyState !== WebSocket.OPEN) {
      s.addMessage({ id: crypto.randomUUID(), role: 'error', text: 'Not connected. Retrying...', timestamp: Date.now() })
      return
    }

    // User message
    s.addMessage({ id: crypto.randomUUID(), role: 'user', text, fileIds, timestamp: Date.now() })

    // Pending agent placeholder
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
