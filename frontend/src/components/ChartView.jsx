/**
 * RightCut — Live chart view rendered from workbook state using Chart.js.
 * Renders in a dedicated tab alongside spreadsheet tabs.
 */
import { useEffect, useRef, useMemo } from 'react'
import {
  Chart,
  CategoryScale, LinearScale,
  BarElement, LineElement, PointElement, ArcElement,
  Title, Tooltip, Legend,
  BarController, LineController, PieController, ScatterController,
} from 'chart.js'
import { HyperFormula } from 'hyperformula'
import useWorkspaceStore from '../stores/workspaceStore'

// Register all needed chart.js components once
Chart.register(
  CategoryScale, LinearScale,
  BarElement, LineElement, PointElement, ArcElement,
  Title, Tooltip, Legend,
  BarController, LineController, PieController, ScatterController,
)

// Palette that matches RightCut green brand
const PALETTE = [
  '#217346', '#2d9e5e', '#1a5c38', '#4caf80', '#81c99f',
  '#145a32', '#52be80', '#0b7a3e', '#a9dfbf', '#1d6e40',
]

function resolveRef(ref, headers, rows) {
  // Parse a cell ref like "A1" or "B3" → { col: 0, row: 0 } (0-indexed)
  const match = ref.match(/^([A-Z]+)(\d+)$/)
  if (!match) return null
  const col = match[1].split('').reduce((acc, c) => acc * 26 + c.charCodeAt(0) - 64, 0) - 1
  const row = parseInt(match[2], 10) - 1  // 0-indexed; row 0 = headers
  return { col, row }
}

function parseRange(range) {
  // "A1:C5" → { startCol, startRow, endCol, endRow } (0-indexed)
  const parts = range.split(':')
  if (parts.length !== 2) return null
  const start = resolveRef(parts[0].trim(), null, null)
  const end   = resolveRef(parts[1].trim(), null, null)
  if (!start || !end) return null
  return { startCol: start.col, startRow: start.row, endCol: end.col, endRow: end.row }
}

/**
 * Build a HyperFormula-evaluated 2D grid from the sheet data.
 * Returns evaluatedGrid[row][col] where row 0 = headers.
 * Formulas are computed; raw values are preserved.
 */
function buildEvaluatedGrid(sheet, allSheets) {
  const { headers, rows } = sheet
  const headerRow = headers.map((h) => h || '')
  const dataRows = rows.map((row) =>
    headers.map((_, colIdx) => {
      const cell = (row || [])[colIdx]
      if (!cell) return ''
      if (cell.formula) return cell.formula
      if (cell.value !== null && cell.value !== undefined) {
        const num = Number(cell.value)
        return isNaN(num) ? String(cell.value) : num
      }
      return ''
    })
  )
  const hfData = [headerRow, ...dataRows]

  let hf = null
  try {
    if (allSheets && allSheets.length > 1) {
      const sheetsMap = {}
      allSheets.forEach((s) => {
        const hdr = (s.headers || []).map((h) => h || '')
        const drs = (s.rows || []).map((row) =>
          (s.headers || []).map((_, ci) => {
            const c = (row || [])[ci]
            if (!c) return ''
            if (c.formula) return c.formula
            if (c.value !== null && c.value !== undefined) {
              const n = Number(c.value)
              return isNaN(n) ? String(c.value) : n
            }
            return ''
          })
        )
        sheetsMap[s.name] = [hdr, ...drs]
      })
      hf = HyperFormula.buildFromSheets(sheetsMap, { licenseKey: 'gpl-v3' })
    } else {
      hf = HyperFormula.buildFromArray(hfData, { licenseKey: 'gpl-v3' })
    }
    const rawId = (allSheets && allSheets.length > 1) ? hf.getSheetId(sheet.name) : undefined
    const sheetIdx = (rawId != null && rawId !== undefined) ? rawId : 0

    const grid = [headerRow]  // row 0 = headers (not evaluated)
    for (let r = 0; r < rows.length; r++) {
      const evalRow = []
      for (let c = 0; c < headers.length; c++) {
        try {
          const v = hf.getCellValue({ sheet: sheetIdx, row: r + 1, col: c })
          if (v !== null && typeof v === 'object' && 'type' in v) {
            // HF error — use raw value
            const raw = hfData[r + 1]?.[c]
            evalRow.push(typeof raw === 'string' && raw.startsWith('=') ? '' : raw)
          } else {
            evalRow.push(v)
          }
        } catch {
          evalRow.push(hfData[r + 1]?.[c] ?? '')
        }
      }
      grid.push(evalRow)
    }
    return grid
  } catch {
    // Fallback: raw values, blank formulas
    const grid = [headerRow]
    for (let r = 0; r < rows.length; r++) {
      grid.push(headers.map((_, c) => {
        const raw = hfData[r + 1]?.[c]
        return (typeof raw === 'string' && raw.startsWith('=')) ? '' : (raw ?? '')
      }))
    }
    return grid
  } finally {
    hf?.destroy()
  }
}

function getCellVal(col, row, grid) {
  if (!grid || !grid[row]) return ''
  return grid[row][col] ?? ''
}

function buildChartData(chart, sheet, allSheets) {
  const grid = buildEvaluatedGrid(sheet, allSheets)
  const bounds = parseRange(chart.data_range)
  if (!bounds) return null

  const { startCol, startRow, endCol, endRow } = bounds
  const numCols = endCol - startCol + 1

  if (numCols < 2) return null  // need at least label col + one value col

  // First column = labels (categories)
  const labels = []
  for (let r = startRow + 1; r <= endRow; r++) {
    const v = getCellVal(startCol, r, grid)
    labels.push(String(v))
  }

  // Remaining columns = datasets
  const datasets = []
  for (let c = startCol + 1; c <= endCol; c++) {
    const label = String(getCellVal(c, startRow, grid) || `Series ${c - startCol}`)
    const data  = []
    for (let r = startRow + 1; r <= endRow; r++) {
      const v = getCellVal(c, r, grid)
      data.push(v === '' || v === null ? null : Number(v))
    }
    const color = PALETTE[(c - startCol - 1) % PALETTE.length]
    datasets.push({
      label,
      data,
      backgroundColor: chart.chart_type === 'pie'
        ? labels.map((_, i) => PALETTE[i % PALETTE.length])
        : color + 'cc',  // slight transparency for bars
      borderColor: chart.chart_type === 'line' ? color : undefined,
      borderWidth: chart.chart_type === 'line' ? 2 : 0,
      fill: false,
      pointRadius: chart.chart_type === 'line' ? 4 : 0,
      tension: 0.3,
    })
  }

  return { labels, datasets }
}

function buildChartConfig(chart, sheet, allSheets) {
  const data = buildChartData(chart, sheet, allSheets)
  if (!data) return null

  const typeMap = { bar: 'bar', line: 'line', pie: 'pie', scatter: 'scatter' }
  const type = typeMap[chart.chart_type] || 'bar'

  return {
    type,
    data,
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        title: {
          display: !!chart.title,
          text: chart.title,
          font: { size: 14, weight: '600', family: "'IBM Plex Sans', sans-serif" },
          color: getComputedStyle(document.documentElement).getPropertyValue('--text-primary').trim() || '#0f172a',
          padding: { bottom: 16 },
        },
        legend: {
          display: type !== 'bar' || data.datasets.length > 1,
          labels: {
            font: { family: "'IBM Plex Sans', sans-serif", size: 12 },
            color: getComputedStyle(document.documentElement).getPropertyValue('--text-muted').trim() || '#4a5568',
          },
        },
        tooltip: {
          bodyFont: { family: "'IBM Plex Sans', sans-serif" },
          titleFont: { family: "'IBM Plex Sans', sans-serif" },
        },
      },
      scales: type === 'pie' ? {} : {
        x: {
          grid: { color: getComputedStyle(document.documentElement).getPropertyValue('--border').trim() || 'rgba(0,0,0,0.05)' },
          ticks: { font: { family: "'IBM Plex Sans', sans-serif", size: 11 }, color: getComputedStyle(document.documentElement).getPropertyValue('--text-muted').trim() || '#4a5568' },
        },
        y: {
          grid: { color: getComputedStyle(document.documentElement).getPropertyValue('--border').trim() || 'rgba(0,0,0,0.07)' },
          ticks: { font: { family: "'IBM Plex Sans', sans-serif", size: 11 }, color: getComputedStyle(document.documentElement).getPropertyValue('--text-muted').trim() || '#4a5568' },
          beginAtZero: true,
        },
      },
    },
  }
}

export default function ChartView({ sheetName, chartIndex }) {
  const canvasRef  = useRef(null)
  const chartRef   = useRef(null)
  const workbookState = useWorkspaceStore((s) => s.workbookState)

  const allSheets = workbookState?.sheets || []

  const sheet = useMemo(
    () => allSheets.find((s) => s.name === sheetName),
    [allSheets, sheetName],
  )

  const chart = sheet?.charts?.[chartIndex]

  const config = useMemo(
    () => (chart && sheet ? buildChartConfig(chart, sheet, allSheets) : null),
    [chart, sheet, allSheets],
  )

  useEffect(() => {
    if (!canvasRef.current || !config) return

    // Destroy previous instance
    if (chartRef.current) {
      chartRef.current.destroy()
      chartRef.current = null
    }

    chartRef.current = new Chart(canvasRef.current, config)

    return () => {
      chartRef.current?.destroy()
      chartRef.current = null
    }
  }, [config])

  if (!sheet || !chart) {
    return (
      <div className="chart-view-empty">
        <p>Chart data not available.</p>
      </div>
    )
  }

  if (!config) {
    return (
      <div className="chart-view-empty">
        <p>Could not parse data range <code>{chart.data_range}</code>.</p>
      </div>
    )
  }

  return (
    <div className="chart-view">
      <div className="chart-view-inner">
        <canvas ref={canvasRef} />
      </div>
      <div className="chart-view-meta">
        Source: <strong>{sheetName}</strong> · Range: <code>{chart.data_range}</code>
      </div>
    </div>
  )
}
