import { Sheet, FileText } from 'lucide-react'
import useWorkspaceStore from '../stores/workspaceStore'

function TabIcon({ type }) {
  if (type === 'document') return <FileText size={11} />
  return <Sheet size={11} />
}

export default function TabBar() {
  const tabs         = useWorkspaceStore((s) => s.tabs)
  const activeTab    = useWorkspaceStore((s) => s.activeTab)
  const setActiveTab = useWorkspaceStore((s) => s.setActiveTab)

  if (tabs.length === 0) return null

  return (
    <div className="sheet-tab-bar" role="tablist">
      {tabs.map((tab) => (
        <button
          key={tab.id}
          className={`sheet-tab ${activeTab === tab.id ? 'sheet-tab--active' : ''}`}
          onClick={() => setActiveTab(tab.id)}
          title={tab.name}
          role="tab"
          aria-selected={activeTab === tab.id}
        >
          <TabIcon type={tab.type} />
          <span className="sheet-tab-label">{tab.name}</span>
        </button>
      ))}
    </div>
  )
}
