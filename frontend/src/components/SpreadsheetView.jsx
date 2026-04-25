/**
 * RightCut — Excel-style spreadsheet view.
 * Formula bar + row numbers + AG Grid with Excel-like styling.
 */
import { useMemo, useCallback, useState } from 'react'
import { AgGridReact } from 'ag-grid-react'
import {
  ModuleRegistry,
  ClientSideRowModelModule,
  CellStyleModule,
  TextEditorModule,
  RowSelectionModule,
  ValidationModule,
} from 'ag-grid-community'
import 'ag-grid-community/styles/ag-grid.css'
import 'ag-grid-community/styles/ag-theme-alpine.css'
import { BarChart2, Table2 } from 'lucide-react'
import { sheetStateToAgGrid, fieldRowToCell, columnIndexToLetter } from '../utils/sheetSync'

ModuleRegistry.registerModules([
  ClientSideRowModelModule,
  CellStyleModule,
  TextEditorModule,
  RowSelectionModule,
  ValidationModule,
])

export default function SpreadsheetView({ sheet, allSheets, onCellEdit }) {
  const [selectedCell, setSelectedCell] = useState(null)  // { col, row, value, colLetter }

  const { rowData, columnDefs } = useMemo(
    () => sheetStateToAgGrid(sheet, allSheets),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [sheet?.name, sheet?.rows, allSheets]
  )

  // Prepend row-number column
  const allColumnDefs = useMemo(() => {
    const rowNumCol = {
      field: '__rowNum',
      headerName: '',
      width: 46,
      minWidth: 46,
      maxWidth: 46,
      editable: false,
      sortable: false,
      resizable: false,
      suppressMovable: true,
      cellClass: 'excel-row-num',
      headerClass: 'excel-row-num-header',
      valueGetter: (params) => params.data.__rowIdx + 1,
      suppressNavigable: true,
    }
    return [rowNumCol, ...columnDefs]
  }, [columnDefs])

  const onCellValueChanged = useCallback(
    (params) => {
      if (!onCellEdit || !sheet) return
      const cell = fieldRowToCell(params.column.getColId(), params.data.__rowIdx)
      onCellEdit(sheet.name, cell, String(params.oldValue ?? ''), String(params.newValue ?? ''))
    },
    [onCellEdit, sheet]
  )

  const onCellFocused = useCallback(
    (params) => {
      if (!params.column || params.rowIndex == null) return
      const colId = params.column.getColId?.()
      if (!colId || colId === '__rowNum') return
      const colIdx = parseInt(colId.replace('col_', ''), 10)
      const rowIdx = params.rowIndex
      const colLetter = columnIndexToLetter(colIdx + 1)
      const excelRef = `${colLetter}${rowIdx + 1}`
      // show formula if cell has one, else evaluated value
      const formula = rowData[rowIdx]?.[`__formula_${colIdx}`]
      const value = formula ?? rowData[rowIdx]?.[colId] ?? ''
      setSelectedCell({ ref: excelRef, value })
    },
    [rowData]
  )

  const defaultColDef = useMemo(
    () => ({
      resizable: true,
      sortable: false,
      filter: false,
      suppressHeaderMenuButton: true,
    }),
    []
  )

  if (!sheet) {
    return (
      <div className="spreadsheet-empty">
        <div className="empty-icon-lg"><Table2 size={24} /></div>
        <div className="empty-title">No sheets yet</div>
        <div className="empty-subtitle">
          Ask the agent to build a deal sheet, comps table, or financial model.
        </div>
      </div>
    )
  }

  return (
    <div className="spreadsheet-wrapper">
      {/* Formula bar */}
      <div className="formula-bar">
        <div className="formula-bar-ref">
          {selectedCell?.ref || 'A1'}
        </div>
        <div className="formula-bar-divider" />
        <div className="formula-bar-fx">fx</div>
        <div className="formula-bar-value">
          {selectedCell?.value ?? ''}
        </div>
      </div>

      {/* Grid */}
      <div
        className="ag-theme-alpine spreadsheet-grid"
        style={{ width: '100%', height: '100%' }}
      >
        <AgGridReact
          theme="legacy"
          rowData={rowData}
          columnDefs={allColumnDefs}
          defaultColDef={defaultColDef}
          onCellValueChanged={onCellValueChanged}
          onCellFocused={onCellFocused}
          suppressMovableColumns
          enableCellTextSelection
        />
      </div>

      {sheet.charts?.length > 0 && (
        <div className="chart-placeholders">
          {sheet.charts.map((chart, i) => (
            <div key={chart.title || `chart_${i}`} className="chart-placeholder">
              <BarChart2 size={13} />
              <span className="chart-title">{chart.title || chart.chart_type} chart</span>
              <span className="chart-note">Visible in downloaded .xlsx</span>
            </div>
          ))}
        </div>
      )}
    </div>
  )
}
