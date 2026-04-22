import { CheckCircle2, AlertTriangle, Sheet } from 'lucide-react'
import useWorkspaceStore from '../stores/workspaceStore'

const STATUS_CONF = {
  connected:    { color: '#22c55e', label: 'Connected' },
  thinking:     { color: '#f59e0b', label: 'Agent thinking…' },
  connecting:   { color: '#94a3b8', label: 'Connecting…' },
  disconnected: { color: '#ef4444', label: 'Disconnected' },
}

export default function StatusBar() {
  const wsStatus      = useWorkspaceStore((s) => s.wsStatus)
  const formulaCount  = useWorkspaceStore((s) => s.formulaCount)
  const hardcodedCount = useWorkspaceStore((s) => s.hardcodedCount)
  const workbookState = useWorkspaceStore((s) => s.workbookState)

  const conf = STATUS_CONF[wsStatus] || STATUS_CONF.disconnected
  const totalSheets = workbookState?.sheets?.length ?? 0

  return (
    <footer className="status-bar" role="status">
      <div className="status-left">
        <span className="status-dot" style={{ backgroundColor: conf.color }} />
        <span>{conf.label}</span>
      </div>

      <div className="status-center">
        {totalSheets > 0 && (
          <>
            <span className="status-item">
              <Sheet size={11} />
              {totalSheets} sheet{totalSheets !== 1 ? 's' : ''}
            </span>
            {formulaCount > 0 && (
              <span className="status-item status-good">
                <CheckCircle2 size={11} />
                {formulaCount} formula{formulaCount !== 1 ? 's' : ''}
              </span>
            )}
            {hardcodedCount > 0 && (
              <span className="status-item status-warn">
                <AlertTriangle size={11} />
                {hardcodedCount} hardcoded
              </span>
            )}
          </>
        )}
      </div>

      <div className="status-right">RightCut</div>
    </footer>
  )
}
