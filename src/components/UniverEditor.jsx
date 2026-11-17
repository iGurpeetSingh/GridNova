import { useEffect, useRef, useState } from 'react'
import { createUniver, LocaleType, mergeLocales } from '@univerjs/presets'
import { UniverSheetsCorePreset } from '@univerjs/preset-sheets-core'
import '@univerjs/ui/facade'
import '@univerjs/sheets/facade'
import UniverPresetSheetsCoreEnUS from '@univerjs/preset-sheets-core/locales/en-US'
import '@univerjs/preset-sheets-core/lib/index.css'

const DEFAULT_ROW_COUNT = 1000
const DEFAULT_COLUMN_COUNT = 26

function createBlankWorkbookSnapshot(workbookId, title) {
  const sheetId = 'sheet1'
  return {
    id: workbookId,
    name: title || 'Untitled workbook',
    appVersion: '1.0.0',
    locale: 'enUS',
    styles: {},
    sheetOrder: [sheetId],
    sheets: {
      [sheetId]: {
        id: sheetId,
        name: 'Sheet1',
        tabColor: '',
        hidden: 0,
        rowCount: DEFAULT_ROW_COUNT,
        columnCount: DEFAULT_COLUMN_COUNT,
        zoomRatio: 1,
        freeze: { startRow: -1, startColumn: -1, ySplit: 0, xSplit: 0 },
        scrollTop: 0,
        scrollLeft: 0,
        defaultColumnWidth: 100,
        defaultRowHeight: 23,
        mergeData: [],
        cellData: {},
        rowData: {},
        columnData: {},
        rowHeader: { width: 46, hidden: 0 },
        columnHeader: { height: 20, hidden: 0 },
        showGridlines: 1,
        rightToLeft: 0,
      },
    },
  }
}

function ensureMinimumColumnCount(snapshot) {
  if (!snapshot || typeof snapshot !== 'object') return snapshot

  // Work on a copy so a loaded IndexedDB record is not mutated outside the editor.
  const normalized = JSON.parse(JSON.stringify(snapshot))
  for (const sheetId of normalized.sheetOrder || []) {
    const sheet = normalized.sheets?.[sheetId]
    if (!sheet) continue
    const currentCount = Number(sheet.columnCount) || 0
    sheet.columnCount = Math.max(currentCount, DEFAULT_COLUMN_COUNT)
  }
  return normalized
}

function addColumnHeaderDeleteAction(univerAPI) {
  // The core preset already provides Insert Before/After, Hide/Show, Width
  // and Autofit on the column-header context menu. Its public menu list does
  // not include Delete there, so add that one action through the UI Facade.
  const deleteMenu = univerAPI.createMenu({
    id: 'gridnova.context-menu.delete-selected-columns',
    title: 'Delete selected column(s)',
    action: () => {
      const sheet = univerAPI.getActiveWorkbook()?.getActiveSheet()
      const selectedRange = sheet?.getActiveRange()
      if (!sheet || !selectedRange) return

      const firstColumn = selectedRange.getColumn()
      const lastColumn = selectedRange.getLastColumn()
      const columnCount = Math.max(1, lastColumn - firstColumn + 1)
      if (typeof sheet.deleteColumns === 'function') {
        sheet.deleteColumns(firstColumn, columnCount)
      }
    },
  })

  // COL_HEADER is the dedicated right-click menu target for column letters.
  deleteMenu.appendTo(['contextMenu.colHeader', 'contextMenu.others'])
  return deleteMenu
}

export default function UniverEditor({ workbookId, title, initialSnapshot, onReady }) {
  const containerRef = useRef(null)
  const onReadyRef = useRef(onReady)
  const initialConfigRef = useRef({ workbookId, title, initialSnapshot })
  const [error, setError] = useState('')

  useEffect(() => {
    onReadyRef.current = onReady
  }, [onReady])

  useEffect(() => {
    if (!containerRef.current) return undefined

    let univerInstance
    let disposed = false

    try {
      const { univer, univerAPI } = createUniver({
        locale: LocaleType.EN_US,
        locales: {
          [LocaleType.EN_US]: mergeLocales(UniverPresetSheetsCoreEnUS),
        },
        presets: [
          UniverSheetsCorePreset({
            container: containerRef.current,
            header: true,
            toolbar: true,
            contextMenu: true,
            ribbonType: 'grid',
            footer: {
              addSheetButtonConfig: {
                show: true,
                defaultRowCount: DEFAULT_ROW_COUNT,
                defaultColumnCount: DEFAULT_COLUMN_COUNT,
              },
            },
          }),
        ],
      })

      univerInstance = univer
      const config = initialConfigRef.current
      const workbookData = config.initialSnapshot
        ? ensureMinimumColumnCount(config.initialSnapshot)
        : createBlankWorkbookSnapshot(config.workbookId, config.title)

      univerAPI.createWorkbook(workbookData)
      addColumnHeaderDeleteAction(univerAPI)

      const editorApi = {
        getSnapshot() {
          return univerAPI.getActiveWorkbook()?.save() || null
        },
        getActiveWorkbook() {
          return univerAPI.getActiveWorkbook()
        },
        getActiveSheet() {
          return univerAPI.getActiveWorkbook()?.getActiveSheet()
        },
        replaceSnapshot(snapshot) {
          if (!snapshot || !snapshot.sheetOrder?.length) {
            throw new Error('The imported workbook does not contain any worksheets.')
          }
          const normalizedSnapshot = ensureMinimumColumnCount(snapshot)
          const previousId = univerAPI.getActiveWorkbook()?.getId()
          const nextWorkbook = univerAPI.createWorkbook(normalizedSnapshot)
          if (previousId && previousId !== nextWorkbook.getId()) {
            univerAPI.disposeUnit(previousId)
          }
          return nextWorkbook
        },
      }

      if (!disposed) onReadyRef.current?.(editorApi)
    } catch (initializationError) {
      console.error('GridNova editor initialization failed:', initializationError)
      setError(initializationError instanceof Error ? initializationError.message : 'The spreadsheet editor could not be initialized.')
    }

    return () => {
      disposed = true
      onReadyRef.current?.(null)
      try {
        univerInstance?.dispose()
      } catch (disposeError) {
        console.error('GridNova editor cleanup failed:', disposeError)
      }
    }
  }, [workbookId, initialSnapshot])

  return (
    <div className="univer-shell">
      <div ref={containerRef} className="univer-mount" />
      {error && (
        <div className="editor-error" role="alert">
          <strong>GridNova could not start the spreadsheet editor.</strong>
          <span>{error}</span>
          <button type="button" onClick={() => window.location.reload()}>Reload editor</button>
        </div>
      )}
    </div>
  )
}
