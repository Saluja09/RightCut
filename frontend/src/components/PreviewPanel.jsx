import { BarChart3, Download } from 'lucide-react'
import SheetTabs from './TabBar'
import SpreadsheetView from './SpreadsheetView'
import DocumentView from './DocumentView'
import ChartView from './ChartView'
import useWorkspaceStore from '../stores/workspaceStore'
import { useWorkbook } from '../hooks/useWorkbook'
import { apiUrl } from '../utils/api'

export default function PreviewPanel({ sessionId, onCellEdit }) {
  const activeTab     = useWorkspaceStore((s) => s.activeTab)
  const tabs          = useWorkspaceStore((s) => s.tabs)
  const workbookState = useWorkspaceStore((s) => s.workbookState)
  const { currentSheet } = useWorkbook()

  const activeTabInfo = tabs.find((t) => t.id === activeTab)
  const hasWorkbook   = workbookState?.sheets?.length > 0

  const handleDownload = async () => {
    if (!sessionId) return
    try {
      // Always restore the workbook to the backend before downloading.
      // The backend session may have expired (Render free tier spins down),
      // so we push the current workbook state to ensure it's there.
      if (workbookState?.sheets?.length) {
        const restoreRes = await fetch(apiUrl(`/restore/${sessionId}`), {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            workbook_state: workbookState,
            messages: [],
            role: useWorkspaceStore.getState().sessionRole || 'general',
          }),
        })
        if (!restoreRes.ok) {
          alert('Could not prepare workbook for download. The server may be starting up — please try again in a few seconds.')
          return
        }
      } else {
        alert('No workbook to download yet.')
        return
      }

      const res = await fetch(apiUrl(`/download/${sessionId}`))
      if (!res.ok) {
        const err = await res.json().catch(() => ({}))
        alert(err.detail || 'Download failed.')
        return
      }
      const blob = await res.blob()
      const url = URL.createObjectURL(blob)
      const a = document.createElement('a')
      a.href = url
      a.download = 'rightcut_analysis.xlsx'
      a.click()
      URL.revokeObjectURL(url)
    } catch {
      alert('Download failed. The server may be starting up — please try again in a few seconds.')
    }
  }

  return (
    <div className="preview-panel">
      {/* Excel-style toolbar — only when a workbook exists */}
      {hasWorkbook && (
        <div className="excel-toolbar">
          <button className="excel-toolbar-btn excel-toolbar-btn--download" onClick={handleDownload}>
            <Download size={12} />
            Download .xlsx
          </button>
        </div>
      )}

      {/* Main content area */}
      <div className="preview-content">
        {!activeTab ? (
          <WelcomeState />
        ) : activeTabInfo?.type === 'document' ? (
          <DocumentView fileId={activeTab} />
        ) : activeTabInfo?.type === 'chart' ? (
          <ChartView sheetName={activeTabInfo.sheetName} chartIndex={activeTabInfo.chartIndex} />
        ) : (
          <SpreadsheetView sheet={currentSheet} allSheets={workbookState?.sheets} onCellEdit={onCellEdit} />
        )}
      </div>

      {/* Excel-style sheet tabs at the bottom */}
      {tabs.length > 0 && <SheetTabs sessionId={sessionId} />}
    </div>
  )
}

function WelcomeState() {
  return (
    <div className="welcome-state">
      <div className="welcome-icon">
        <BarChart3 size={32} />
      </div>
      <h1 className="welcome-title">RightCut</h1>
      <p className="welcome-subtitle">AI Spreadsheet Agent for Private Markets</p>
      <p className="welcome-hint">Your model will appear here</p>
    </div>
  )
}
