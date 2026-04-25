/**
 * RightCut — Document preview tab.
 * Shows extracted text and tables from an uploaded document.
 */
import { FileText } from 'lucide-react'
import useWorkspaceStore from '../stores/workspaceStore'

export default function DocumentView({ fileId }) {
  const documents = useWorkspaceStore((s) => s.documents)
  const doc = documents[fileId]

  if (!doc) {
    return (
      <div className="doc-view doc-view--empty">
        <p>Document not found.</p>
      </div>
    )
  }

  return (
    <div className="doc-view">
      <div className="doc-header">
        <span className="doc-icon"><FileText size={16} /></span>
        <div className="doc-meta">
          <span className="doc-filename">{doc.filename}</span>
          {doc.page_count && (
            <span className="doc-pages">{doc.page_count} pages</span>
          )}
        </div>
      </div>

      {doc.tables?.length > 0 && (
        <div className="doc-tables">
          <div className="doc-section-label">Extracted Tables ({doc.tables.length})</div>
          {doc.tables.slice(0, 5).map((table, ti) => (
            <div key={ti} className="doc-table-wrapper">
              <div className="doc-table-title">Table {ti + 1}</div>
              <div className="doc-table-scroll">
                <table className="doc-table">
                  <thead>
                    <tr>
                      {(table[0] || []).map((cell, ci) => (
                        <th key={ci}>{cell}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {table.slice(1, 20).map((row, ri) => (
                      <tr key={ri}>
                        {row.map((cell, ci) => (
                          <td key={ci}>{cell}</td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          ))}
        </div>
      )}

      <div className="doc-content">
        <div className="doc-section-label">Extracted Text</div>
        <pre className="doc-text">{doc.content || 'No text extracted.'}</pre>
      </div>
    </div>
  )
}
