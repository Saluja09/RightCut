/**
 * RightCut — useWorkbook hook.
 * Derived state and helpers for workbook operations.
 */
import { useMemo } from 'react'
import useWorkspaceStore from '../stores/workspaceStore'

export function useWorkbook() {
  const workbookState = useWorkspaceStore((s) => s.workbookState)
  const activeSheet = useWorkspaceStore((s) => s.activeSheet)

  const currentSheet = useMemo(() => {
    if (!workbookState || !activeSheet) return null
    return workbookState.sheets?.find((s) => s.name === activeSheet) || null
  }, [workbookState, activeSheet])

  const sheetNames = useMemo(() => {
    return workbookState?.sheets?.map((s) => s.name) || []
  }, [workbookState])

  const isEmpty = !workbookState || !workbookState.sheets?.length

  return {
    workbookState,
    currentSheet,
    sheetNames,
    isEmpty,
  }
}
