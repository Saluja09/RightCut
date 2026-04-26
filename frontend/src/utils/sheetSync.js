/**
 * RightCut — Sheet sync utilities.
 * Coordinate conversion and AG Grid data transformation.
 * Formulas are evaluated via HyperFormula before display.
 */
import { HyperFormula } from 'hyperformula'

/**
 * Convert 1-indexed column number to Excel letter(s).
 * e.g. 1→'A', 26→'Z', 27→'AA'
 */
export function columnIndexToLetter(n) {
  let result = ''
  while (n > 0) {
    const remainder = (n - 1) % 26
    result = String.fromCharCode(65 + remainder) + result
    n = Math.floor((n - 1) / 26)
  }
  return result
}

/**
 * Format a numeric value for display.
 * Rounds floats to reasonable precision.
 */
function formatValue(val) {
  if (val === null || val === undefined) return ''
  if (typeof val === 'number') {
    if (Number.isNaN(val) || !Number.isFinite(val)) return ''
    // Round to 4 significant decimal places max
    return parseFloat(val.toPrecision(10)).toString()
  }
  return String(val)
}

/**
 * Convert SheetState (from backend) to AG Grid rowData + columnDefs.
 * Evaluates formulas via HyperFormula.
 * Returns { rowData: object[], columnDefs: ColDef[] }
 */
function sheetToHfData(sheet) {
  const headers = sheet.headers || []
  const rows = sheet.rows || []
  // Prepend header row so HF row index 0 = Excel row 1 (header).
  // This ensures absolute references like Assumptions!$B$3 resolve to the correct row.
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
  return [headerRow, ...dataRows]
}

/**
 * Convert SheetState to AG Grid rowData + columnDefs.
 * Pass allSheets (the full workbook sheets array) to enable cross-sheet formula evaluation.
 */
export function sheetStateToAgGrid(sheet, allSheets) {
  if (!sheet) return { rowData: [], columnDefs: [] }

  const headers = sheet.headers || []
  const rows = sheet.rows || []

  // Build raw 2D array for the current sheet (includes header row at index 0)
  const hfData = sheetToHfData(sheet)
  // hfData[0] = header row, hfData[1..] = data rows matching sheet.rows indices

  // Evaluate with HyperFormula — use multi-sheet if allSheets provided
  let hf = null
  let evaluated = null
  try {
    if (allSheets && allSheets.length > 1) {
      // Build named sheets map: { SheetName: [[...], ...] }
      const sheetsMap = {}
      allSheets.forEach((s) => { sheetsMap[s.name] = sheetToHfData(s) })
      hf = HyperFormula.buildFromSheets(sheetsMap, { licenseKey: 'gpl-v3' })
    } else {
      hf = HyperFormula.buildFromArray(hfData, { licenseKey: 'gpl-v3' })
    }
    const rawSheetId = (allSheets && allSheets.length > 1) ? hf.getSheetId(sheet.name) : undefined
    const sheetIdx = (rawSheetId != null && rawSheetId !== undefined) ? rawSheetId : 0
    let errorCount = 0
    let formulaCount = 0
    // Only evaluate data rows (skip header at index 0); map back to rows[] indices
    const raw = rows.map((_, rowIdx) => {
      const hfRow = rowIdx + 1  // offset by 1 because hfData[0] is the header
      return headers.map((_, c) => {
        const cell = hfData[hfRow]?.[c]
        const isFormula = typeof cell === 'string' && cell.startsWith('=')
        if (isFormula) formulaCount++
        try {
          const v = hf.getCellValue({ sheet: sheetIdx, row: hfRow, col: c })
          if (v !== null && typeof v === 'object' && 'type' in v) {
            if (isFormula) errorCount++
            return formatValue(cell)
          }
          return formatValue(v)
        } catch {
          return formatValue(cell)
        }
      })
    })
    // If >30% of formulas errored (likely circular refs), skip HF and show raw values
    if (formulaCount > 0 && errorCount / formulaCount > 0.3) {
      evaluated = rows.map((_, rowIdx) =>
        headers.map((_, c) => {
          const cell = hfData[rowIdx + 1]?.[c]
          if (typeof cell === 'string' && cell.startsWith('=')) return ''
          return formatValue(cell)
        })
      )
    } else {
      evaluated = raw
    }
  } catch {
    // Fallback: use raw values if HyperFormula fails
    evaluated = rows.map((_, rowIdx) =>
      headers.map((_, c) => {
        const cell = hfData[rowIdx + 1]?.[c]
        if (typeof cell === 'string' && cell.startsWith('=')) return ''
        return formatValue(cell)
      })
    )
  } finally {
    hf?.destroy()
  }

  // Build column definitions — compute min width from header AND data values
  const colMaxLengths = headers.map((header, colIdx) => {
    let maxLen = (header?.length || 6)
    rows.forEach((row) => {
      const cell = (row || [])[colIdx]
      const v = cell?.value
      if (v !== null && v !== undefined) {
        const len = String(v).length
        if (len > maxLen) maxLen = len
      }
    })
    return maxLen
  })

  const columnDefs = headers.map((header, colIdx) => {
    const field = `col_${colIdx}`
    return {
      field,
      headerName: header || `Col ${colIdx + 1}`,
      width: Math.min(Math.max(colMaxLengths[colIdx] * 8, 80), 320),
      editable: true,
      cellStyle: (params) => {
        const row = params.data
        if (!row) return null
        const cellMeta = row[`__meta_${colIdx}`]
        if (!cellMeta) return null
        const style = {}
        if (cellMeta.bold) style.fontWeight = 'bold'
        if (cellMeta.font_color && cellMeta.font_color !== '00000000') {
          style.color = `#${cellMeta.font_color.slice(-6)}`
        }
        if (cellMeta.bg_color && cellMeta.bg_color !== '00000000') {
          style.backgroundColor = `#${cellMeta.bg_color.slice(-6)}`
        }
        return style
      },
    }
  })

  // Build row data — each row is { col_0: val, col_1: val, __rowIdx: n, __meta_0: {...} }
  const rowData = rows.map((row, rowIdx) => {
    const rowObj = { __rowIdx: rowIdx }
    headers.forEach((_, colIdx) => {
      rowObj[`col_${colIdx}`] = evaluated?.[rowIdx]?.[colIdx] ?? ''
      const cell = (row || [])[colIdx]
      if (cell) {
        // Store formula for formula bar display
        if (cell.formula) rowObj[`__formula_${colIdx}`] = cell.formula
        // Always store cell meta so cellStyle can apply colors, bold, number formats
        if (cell.bold || cell.font_color || cell.bg_color || cell.number_format) {
          rowObj[`__meta_${colIdx}`] = cell
        }
      }
    })
    return rowObj
  })

  // Ensure at least 20 empty rows
  while (rowData.length < 20) {
    const empty = { __rowIdx: rowData.length }
    headers.forEach((_, i) => { empty[`col_${i}`] = '' })
    rowData.push(empty)
  }

  return { rowData, columnDefs }
}

/**
 * Given AG Grid column field name and row __rowIdx, return Excel cell ref.
 * e.g. field="col_2", rowIdx=0 → "C1"
 */
export function fieldRowToCell(field, rowIdx) {
  const colIdx = parseInt(field.replace('col_', ''), 10)
  return `${columnIndexToLetter(colIdx + 1)}${rowIdx + 1}`
}
