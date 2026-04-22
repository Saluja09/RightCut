import { BarChart3, Download } from 'lucide-react'
import SheetTabs from './TabBar'
import SpreadsheetView from './SpreadsheetView'
import DocumentView from './DocumentView'
import useWorkspaceStore from '../stores/workspaceStore'
import { useWorkbook } from '../hooks/useWorkbook'

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
      const res = await fetch(`/download/${sessionId}`)
      if (!res.ok) {
        const err = await res.json().catch(() => ({}))
        alert(err.detail || 'Download failed. Please re-run your analysis to regenerate the workbook.')
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
      alert('Download failed. Please re-run your analysis to regenerate the workbook.')
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
