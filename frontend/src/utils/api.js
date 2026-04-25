/**
 * RightCut — API URL helper.
 * In production, VITE_BACKEND_URL points to the Render deployment.
 * In local dev, requests go through Vite proxy (same origin).
 */
const BACKEND_URL = import.meta.env.VITE_BACKEND_URL || ''

export function apiUrl(path) {
  if (BACKEND_URL) return `${BACKEND_URL.replace(/\/$/, '')}${path}`
  return path
}

export function wsUrl(sessionId) {
  if (BACKEND_URL) {
    const base = BACKEND_URL.replace(/^http/, 'ws').replace(/\/$/, '')
    return `${base}/ws/${sessionId}`
  }
  const proto = location.protocol === 'https:' ? 'wss:' : 'ws:'
  return `${proto}//${location.host}/ws/${sessionId}`
}
