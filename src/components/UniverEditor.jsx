import { useEffect, useRef, useState } from 'react'
import { ArrowDownloadRegular, ArrowRedoRegular, ArrowUndoRegular, ArrowUploadRegular, BookOpenRegular, BorderAllRegular, CalculatorRegular, ClipboardPasteRegular, CurrencyDollarRupeeRegular, CommentRegular, CopyRegular, CutRegular, DrawShapeRegular, FilterRegular, GridRegular, ImageRegular, LinkRegular, PaintBucketRegular, PaintBrushRegular, PrintRegular, QuestionCircleRegular, SaveRegular, SearchRegular, ShareRegular, TableRegular, TextAlignCenterRegular, TextAlignLeftRegular, TextAlignRightRegular, TextBoldRegular, TextColorRegular, TextItalicRegular, TextStrikethroughRegular, TextUnderlineRegular } from '@fluentui/react-icons'
import { createUniver, LocaleType } from '@univerjs/presets'
import { UniverSheetsCorePreset } from '@univerjs/preset-sheets-core'
import '@univerjs/ui/facade'
import '@univerjs/sheets/facade'
import UniverPresetSheetsCoreEnUS from '@univerjs/preset-sheets-core/locales/en-US'
import { UniverSheetsDrawingPreset } from '@univerjs/preset-sheets-drawing'
import UniverPresetSheetsDrawingEnUS from '@univerjs/preset-sheets-drawing/locales/en-US'
import { UniverSheetsTablePreset } from '@univerjs/preset-sheets-table'
import UniverPresetSheetsTableEnUS from '@univerjs/preset-sheets-table/locales/en-US'
import { UniverSheetsThreadCommentPreset } from '@univerjs/preset-sheets-thread-comment'
import UniverPresetSheetsThreadCommentEnUS from '@univerjs/preset-sheets-thread-comment/locales/en-US'
import { UniverSheetsNotePreset } from '@univerjs/preset-sheets-note'
import UniverPresetSheetsNoteEnUS from '@univerjs/preset-sheets-note/locales/en-US'
import { UniverSheetsDataValidationPreset } from '@univerjs/preset-sheets-data-validation'
import UniverPresetSheetsDataValidationEnUS from '@univerjs/preset-sheets-data-validation/locales/en-US'
import '@univerjs/preset-sheets-core/lib/index.css'
import '@univerjs/preset-sheets-drawing/lib/index.css'
import '@univerjs/preset-sheets-table/lib/index.css'
import '@univerjs/preset-sheets-thread-comment/lib/index.css'
import '@univerjs/preset-sheets-note/lib/index.css'
import '@univerjs/preset-sheets-data-validation/lib/index.css'

const DEFAULT_ROW_COUNT = 1000
const DEFAULT_COLUMN_COUNT = 26

function mergeLocaleMessages(...packs) {
  const mergeNode = (target, source) => {
    Object.entries(source || {}).forEach(([key, value]) => {
      if (value && typeof value === 'object' && !Array.isArray(value)) {
        target[key] = mergeNode(
          target[key] && typeof target[key] === 'object' && !Array.isArray(target[key]) ? target[key] : {},
          value,
        )
      } else {
        target[key] = value
      }
    })
    return target
  }

  return packs.reduce((result, pack) => mergeNode(result, pack), {})
}

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


const RIBBON_ICONS = {
  save: SaveRegular,
  import: ArrowUploadRegular,
  export: ArrowDownloadRegular,
  print: PrintRegular,
  share: ShareRegular,
  review: CommentRegular,
  help: QuestionCircleRegular,
  draw: DrawShapeRegular,
  layout: TableRegular,
  home: GridRegular,
  table: TableRegular,
  guide: BookOpenRegular,
  undo: ArrowUndoRegular,
  redo: ArrowRedoRegular,
  bold: TextBoldRegular,
  italic: TextItalicRegular,
  underline: TextUnderlineRegular,
  strike: TextStrikethroughRegular,
  alignLeft: TextAlignLeftRegular,
  alignCenter: TextAlignCenterRegular,
  alignRight: TextAlignRightRegular,
  textColor: TextColorRegular,
  fillColor: PaintBucketRegular,
  border: BorderAllRegular,
  paste: ClipboardPasteRegular,
  cut: CutRegular,
  copy: CopyRegular,
  filter: FilterRegular,
  sort: TableRegular,
  search: SearchRegular,
  calculator: CalculatorRegular,
  currency: CurrencyDollarRupeeRegular,
  percent: CalculatorRegular,
  image: ImageRegular,
  link: LinkRegular,
  paint: PaintBrushRegular,
}

const RIBBON_ICON_COLORS = {
  save: '#217346', share: '#107C41', review: '#7B2CBF', help: '#0078D4',
  draw: '#C65911', image: '#0078D4', table: '#217346', filter: '#0078D4',
  sort: '#0078D4', currency: '#107C41', percent: '#0078D4', paint: '#EAA300',
  fillColor: '#F2C811', textColor: '#C50F1F', paste: '#0078D4',
  cut: '#D83B01', copy: '#0078D4', border: '#64748B',
  alignLeft: '#0078D4', alignCenter: '#0078D4', alignRight: '#0078D4',
  calculator: '#0078D4', print: '#475569', layout: '#0078D4',
}

function registerRibbonIcons(univerAPI) {
  const disposables = []
  Object.entries(RIBBON_ICONS).forEach(([name, Icon]) => {
    const ColorIcon = (props) => <Icon {...props} style={{ ...props.style, color: RIBBON_ICON_COLORS[name] || '#475569' }} />
    disposables.push(univerAPI.registerComponent('gridnova-icon-' + name, ColorIcon))
  })
  return disposables
}

function addRibbonTab(univerAPI, tab, onRibbonAction) {
  const root = univerAPI.createSubmenu({
    // Separate localized tab keys from their group keys. LocaleService resolves
    // dotted paths, so a tab label cannot also be the parent of a group object.
    id: 'gridnova.ribbon.tabs.' + tab.id,
    title: tab.title,
    order: tab.order,
  })

  tab.groups.forEach((group) => {
    const groupMenu = univerAPI.createSubmenu({
      id: 'gridnova.ribbon.groups.' + tab.id + '.' + group.id,
      title: group.title,
    })

    group.items.forEach((item) => {
      const action = item.run
        ? () => item.run(univerAPI)
        : item.command
          ? () => univerAPI.executeCommand(item.command, item.params)
          : () => onRibbonAction?.(item.action)
      const menu = univerAPI.createMenu({
        id: 'gridnova.ribbon.commands.' + tab.id + '.' + item.id,
        title: item.title,
        tooltip: item.tooltip || item.title,
        icon: item.icon ? 'gridnova-icon-' + item.icon : undefined,
        gridLayout: item.gridLayout,
        action,
      })
      groupMenu.addSubmenu(menu)
    })

    root.addSubmenu(groupMenu)
  })

  root.appendTo('ribbon')
  return root
}



function GridNovaBarChart({ data }) {
  const labels = Array.isArray(data?.labels) ? data.labels : []
  const values = Array.isArray(data?.values) ? data.values : []
  const maxValue = Math.max(1, ...values.map((value) => Math.max(0, Number(value) || 0)))
  const chartWidth = 330
  const chartHeight = 150
  const barGap = 10
  const barWidth = Math.max(8, (chartWidth - Math.max(0, values.length - 1) * barGap) / Math.max(1, values.length))
  return (
    <div style={{ width: '100%', height: '100%', padding: 8, boxSizing: 'border-box', background: '#fff', border: '1px solid #cbd5e1', borderRadius: 8 }}>
      <svg viewBox="0 0 360 205" width="100%" height="100%" role="img" aria-label="GridNova bar chart">
        <line x1="20" y1="165" x2="348" y2="165" stroke="#94a3b8" strokeWidth="1" />
        {values.map((value, index) => {
          const numeric = Math.max(0, Number(value) || 0)
          const height = numeric / maxValue * chartHeight
          const x = 22 + index * (barWidth + barGap)
          const y = 165 - height
          return (
            <g key={index}>
              <rect x={x} y={y} width={barWidth} height={height} rx="3" fill="#217346" />
              <text x={x + barWidth / 2} y="181" textAnchor="middle" fontSize="9" fill="#475569">{String(labels[index] || '').slice(0, 9)}</text>
              <text x={x + barWidth / 2} y={Math.max(12, y - 5)} textAnchor="middle" fontSize="9" fill="#334155">{numeric}</text>
            </g>
          )
        })}
      </svg>
    </div>
  )
}

function insertBasicChart(univerAPI) {
  const sheet = univerAPI.getActiveWorkbook()?.getActiveSheet()
  const range = sheet?.getActiveRange()
  if (!sheet || !range || typeof sheet.addFloatDomToRange !== 'function') {
    univerAPI.showMessage({ content: 'Select a two-column data range to insert a chart.', type: 'warning' })
    return
  }
  const rows = range.getValues()
  const dataRows = rows.map((row, index) => ({
    label: String(row?.[0] ?? index + 1),
    value: Number(row?.[1]),
  })).filter((row) => Number.isFinite(row.value))
  if (!dataRows.length) {
    univerAPI.showMessage({ content: 'The selected range needs labels in the first column and numbers in the second column.', type: 'warning' })
    return
  }
  const id = 'gridnova-chart-' + (globalThis.crypto?.randomUUID?.() || Date.now())
  const chart = sheet.addFloatDomToRange(
    range,
    { componentKey: 'gridnova-bar-chart', data: { labels: dataRows.map((row) => row.label), values: dataRows.map((row) => row.value) }, allowTransform: true },
    { width: 360, height: 220 },
    id,
  )
  if (!chart) {
    univerAPI.showMessage({ content: 'The chart could not be inserted.', type: 'error' })
    return
  }
  univerAPI.showMessage({ content: 'Bar chart inserted over the selected data range.', type: 'success' })
}

function GridNovaRectangleShape() {
  return <svg viewBox="0 0 180 100" width="100%" height="100%" aria-label="Rectangle shape">
    <rect x="8" y="8" width="164" height="84" rx="5" fill="#dbeafe" fillOpacity=".55" stroke="#2563eb" strokeWidth="3" />
  </svg>
}

function GridNovaEllipseShape() {
  return <svg viewBox="0 0 180 100" width="100%" height="100%" aria-label="Ellipse shape">
    <ellipse cx="90" cy="50" rx="78" ry="38" fill="#dcfce7" fillOpacity=".55" stroke="#16834b" strokeWidth="3" />
  </svg>
}

function GridNovaLineShape() {
  return <svg viewBox="0 0 180 100" width="100%" height="100%" aria-label="Line shape">
    <line x1="8" y1="88" x2="172" y2="12" stroke="#d97706" strokeWidth="4" strokeLinecap="round" />
  </svg>
}

function registerDrawingComponents(univerAPI) {
  return [
    univerAPI.registerComponent('gridnova-draw-rectangle', GridNovaRectangleShape),
    univerAPI.registerComponent('gridnova-draw-ellipse', GridNovaEllipseShape),
    univerAPI.registerComponent('gridnova-draw-line', GridNovaLineShape),
    univerAPI.registerComponent('gridnova-bar-chart', GridNovaBarChart),
  ]
}

function insertDrawingShape(univerAPI, shapeType) {
  const sheet = univerAPI.getActiveWorkbook()?.getActiveSheet()
  if (!sheet || typeof sheet.addFloatDomToRange !== 'function') {
    univerAPI.showMessage({ content: 'Drawing overlays are not available in this editor build.', type: 'warning' })
    return
  }
  const range = sheet.getActiveRange() || sheet.getRange('A1')
  const componentKey = 'gridnova-draw-' + shapeType
  const id = 'gridnova-shape-' + (globalThis.crypto?.randomUUID?.() || Date.now())
  const shape = sheet.addFloatDomToRange(
    range,
    { componentKey, allowTransform: true },
    { width: 180, height: 100 },
    id,
  )
  if (!shape) {
    univerAPI.showMessage({ content: 'The shape could not be inserted.', type: 'error' })
    return
  }
  univerAPI.showMessage({ content: shapeType[0].toUpperCase() + shapeType.slice(1) + ' added over the selected range.', type: 'success' })
}

async function formatSelectionAsTable(univerAPI) {
  const sheet = univerAPI.getActiveWorkbook()?.getActiveSheet()
  const range = sheet?.getActiveRange()
  if (!sheet || !range) {
    univerAPI.showMessage({ content: 'Select a data range first.', type: 'warning' })
    return
  }
  const rowCount = range.getLastRow() - range.getRow() + 1
  const columnCount = range.getLastColumn() - range.getColumn() + 1
  if (rowCount < 2 || columnCount < 1) {
    univerAPI.showMessage({ content: 'Select a header row and at least one data row.', type: 'warning' })
    return
  }
  if (typeof sheet.addTable !== 'function') {
    univerAPI.showMessage({ content: 'Structured tables are unavailable in this editor build.', type: 'warning' })
    return
  }

  const tableId = 'gridnova-table-' + (globalThis.crypto?.randomUUID?.() || Date.now())
  const tableName = 'GridNovaTable' + Date.now()
  const success = await sheet.addTable(tableName, range.getRange(), tableId, {
    tableStyleId: 'table-default-4',
  })
  if (!success) {
    univerAPI.showMessage({ content: 'The selected range could not be converted to a structured table. Check that it does not overlap another table.', type: 'error' })
    return
  }
  univerAPI.showMessage({ content: 'Structured table created with filter controls.', type: 'success' })
}

function renameActiveSheet(univerAPI) {
  const sheet = univerAPI.getActiveWorkbook()?.getActiveSheet()
  if (!sheet) return
  const nextName = window.prompt('Enter a new worksheet name', sheet.getName())
  if (nextName == null || !nextName.trim()) return
  try { sheet.setName(nextName.trim()) }
  catch (error) { univerAPI.showMessage({ content: error instanceof Error ? error.message : 'Could not rename this worksheet.', type: 'error' }) }
}

function duplicateActiveSheet(univerAPI) {
  const workbook = univerAPI.getActiveWorkbook()
  if (!workbook?.getActiveSheet()) return
  try { workbook.duplicateActiveSheet() }
  catch (error) { univerAPI.showMessage({ content: error instanceof Error ? error.message : 'Could not duplicate this worksheet.', type: 'error' }) }
}

function deleteActiveSheet(univerAPI) {
  const workbook = univerAPI.getActiveWorkbook()
  const sheet = workbook?.getActiveSheet()
  if (!workbook || !sheet) return
  if (workbook.getSheets().length <= 1) {
    univerAPI.showMessage({ content: 'A workbook must keep at least one worksheet.', type: 'warning' })
    return
  }
  if (window.confirm('Delete worksheet "' + sheet.getName() + '"? This cannot be undone after saving.')) {
    workbook.deleteSheet(sheet)
  }
}

function hideActiveSheet(univerAPI) {
  const workbook = univerAPI.getActiveWorkbook()
  const sheet = workbook?.getActiveSheet()
  if (!workbook || !sheet) return
  if (workbook.getSheets().filter((item) => !item.isSheetHidden()).length <= 1) {
    univerAPI.showMessage({ content: 'At least one worksheet must remain visible.', type: 'warning' })
    return
  }
  sheet.hideSheet()
}

function unhideWorksheet(univerAPI) {
  const workbook = univerAPI.getActiveWorkbook()
  const hidden = workbook?.getSheets().filter((sheet) => sheet.isSheetHidden()) || []
  if (!hidden.length) {
    univerAPI.showMessage({ content: 'There are no hidden worksheets.', type: 'info' })
    return
  }
  const selected = window.prompt('Enter the worksheet to unhide:\n' + hidden.map((sheet, index) => (index + 1) + '. ' + sheet.getName()).join('\n'), hidden[0].getName())
  if (selected == null) return
  const sheet = hidden.find((item) => item.getName().toLowerCase() === selected.trim().toLowerCase())
  if (!sheet) {
    univerAPI.showMessage({ content: 'No hidden worksheet matches that name.', type: 'warning' })
    return
  }
  sheet.showSheet()
}

function setActiveSheetTabColor(univerAPI) {
  const sheet = univerAPI.getActiveWorkbook()?.getActiveSheet()
  if (!sheet) return
  const color = window.prompt('Enter a tab color (CSS color or hex). Leave blank to reset.', sheet.getTabColor?.() || '#16834b')
  if (color == null) return
  sheet.setTabColor(color.trim())
}

async function insertFormulaFunction(univerAPI, suggestedFunction = 'SUM') {
  const value = window.prompt('Enter a spreadsheet function name (without the leading =).', suggestedFunction)
  if (value == null) return
  const functionName = value.trim().replace(/^=/, '').replace(/\s*\(.*/, '').toUpperCase()
  if (!/^[A-Z][A-Z0-9_.]*$/.test(functionName)) {
    univerAPI.showMessage({ content: 'Enter a valid function name, such as SUM or VLOOKUP.', type: 'warning' })
    return
  }

  try {
    const result = await univerAPI.executeCommand('formula-ui.operation.insert-function', { value: functionName })
    if (result === false) {
      univerAPI.showMessage({ content: 'Select a cell or a supported range before inserting a function.', type: 'warning' })
    }
  } catch (error) {
    univerAPI.showMessage({ content: error instanceof Error ? error.message : 'The function could not be inserted.', type: 'error' })
  }
}

function showFeatureUnavailable(univerAPI, featureName, reason) {
  univerAPI.showMessage({
    content: featureName + ' is not available in the current GridNova editor integration. ' + reason,
    type: 'info',
  })
}

async function recalculateFormulas(univerAPI) {
  const formula = univerAPI.getFormula?.()
  if (!formula || typeof formula.executeCalculation !== 'function') {
    showFeatureUnavailable(univerAPI, 'Formula recalculation', 'The installed formula service does not expose recalculation.')
    return
  }

  try {
    await formula.executeCalculation()
    univerAPI.showMessage({ content: 'Formula recalculation completed.', type: 'success' })
  } catch (error) {
    univerAPI.showMessage({ content: error instanceof Error ? error.message : 'Formula recalculation failed.', type: 'error' })
  }
}

function setSheetZoom(univerAPI, requestedPercentage) {
  const sheet = univerAPI.getActiveWorkbook()?.getActiveSheet()
  if (!sheet) return
  let percentage = requestedPercentage
  if (percentage == null) {
    const input = window.prompt('Enter zoom percentage (50–200)', String(Math.round(sheet.getZoom() * 100)))
    if (input === null) return
    percentage = Number(input)
  }
  if (!Number.isFinite(percentage) || percentage < 50 || percentage > 200) {
    univerAPI.showMessage({ content: 'Enter a zoom value between 50 and 200.', type: 'warning' })
    return
  }
  sheet.zoom(percentage / 100)
}

async function toggleFullscreen(univerAPI) {
  try {
    if (document.fullscreenElement) await document.exitFullscreen()
    else await document.documentElement.requestFullscreen()
  } catch {
    showFeatureUnavailable(univerAPI, 'Fullscreen', 'The browser did not grant fullscreen access.')
  }
}

function unfreezePanes(univerAPI) {
  univerAPI.getActiveWorkbook()?.getActiveSheet()?.cancelFreeze()
}

function hideSelectedRows(univerAPI) {
  const sheet = univerAPI.getActiveWorkbook()?.getActiveSheet()
  const range = sheet?.getActiveRange()
  if (!sheet || !range) return
  sheet.hideRows(range.getRow(), range.getLastRow() - range.getRow() + 1)
}

function showSelectedRows(univerAPI) {
  const sheet = univerAPI.getActiveWorkbook()?.getActiveSheet()
  const range = sheet?.getActiveRange()
  if (!sheet || !range) return
  sheet.showRows(range.getRow(), range.getLastRow() - range.getRow() + 1)
}

function toggleGridlines(univerAPI) {
  const sheet = univerAPI.getActiveWorkbook()?.getActiveSheet()
  if (!sheet) return
  sheet.setHiddenGridlines(!sheet.hasHiddenGridLines())
}

function freezeFirstRow(univerAPI) {
  const sheet = univerAPI.getActiveWorkbook()?.getActiveSheet()
  if (!sheet) return
  sheet.setFrozenRows(1)
}

function freezeFirstColumn(univerAPI) {
  const sheet = univerAPI.getActiveWorkbook()?.getActiveSheet()
  if (!sheet) return
  sheet.setFrozenColumns(1)
}

function removeDuplicatesFromSelection(univerAPI) {
  const sheet = univerAPI.getActiveWorkbook()?.getActiveSheet()
  const range = sheet?.getActiveRange()
  if (!range) {
    univerAPI.showMessage({ content: 'Select a range before removing duplicates.', type: 'warning' })
    return
  }

  const values = range.getValues()
  if (!Array.isArray(values) || values.length < 2) {
    univerAPI.showMessage({ content: 'The selected range does not contain duplicate rows to remove.', type: 'info' })
    return
  }

  const hasHeader = window.confirm('Does the first selected row contain column headers?')
  const firstDataRow = hasHeader ? 1 : 0
  const seen = new Set()
  const duplicateRowIndexes = []

  values.forEach((row, index) => {
    if (index < firstDataRow) return
    const key = JSON.stringify(row)
    if (seen.has(key)) duplicateRowIndexes.push(range.getRow() + index)
    else seen.add(key)
  })

  if (!duplicateRowIndexes.length) {
    univerAPI.showMessage({ content: 'No duplicate rows were found in the selected range.', type: 'info' })
    return
  }

  if (typeof sheet.deleteRowsByPoints !== 'function') {
    showFeatureUnavailable(univerAPI, 'Remove Duplicates', 'This worksheet facade does not expose non-contiguous row deletion.')
    return
  }

  // The facade deletes the selected worksheet rows as one undoable command path;
  // row indexes are zero-based and the API handles reverse-order deletion.
  sheet.deleteRowsByPoints(duplicateRowIndexes)
  univerAPI.showMessage({ content: duplicateRowIndexes.length + ' duplicate row(s) removed from the worksheet.', type: 'success' })
}

function splitTextToColumns(univerAPI) {
  const sheet = univerAPI.getActiveWorkbook()?.getActiveSheet()
  const range = sheet?.getActiveRange()
  if (!sheet || !range) {
    univerAPI.showMessage({ content: 'Select one source column before using Text to Columns.', type: 'warning' })
    return
  }

  if (range.getColumn() !== range.getLastColumn()) {
    univerAPI.showMessage({ content: 'Text to Columns currently requires a single selected source column.', type: 'warning' })
    return
  }

  const delimiterInput = window.prompt('Choose a delimiter: comma, semicolon, pipe, tab, or enter a custom character.', 'comma')
  if (delimiterInput === null) return
  const delimiterNames = { comma: ',', semicolon: ';', pipe: '|', tab: '\t' }
  const delimiter = delimiterNames[delimiterInput.trim().toLowerCase()] || delimiterInput
  if (!delimiter) {
    univerAPI.showMessage({ content: 'Enter a delimiter to continue.', type: 'warning' })
    return
  }

  const values = range.getValues()
  const splitRows = values.map((row) => String(row?.[0] ?? '').split(delimiter))
  const outputColumnCount = Math.max(1, ...splitRows.map((row) => row.length))
  if (outputColumnCount > 1) {
    sheet.insertColumnsAfter(range.getColumn(), outputColumnCount - 1)
  }

  const output = splitRows.map((row) => {
    const normalized = row.slice(0, outputColumnCount)
    while (normalized.length < outputColumnCount) normalized.push('')
    return normalized
  })
  sheet.getRange(range.getRow(), range.getColumn(), output.length, outputColumnCount).setValues(output)
  sheet.autoResizeColumns(range.getColumn(), outputColumnCount)
  univerAPI.showMessage({ content: 'Split ' + output.length + ' row(s) into ' + outputColumnCount + ' column(s).', type: 'success' })
}

function addMicrosoftInspiredRibbonTabs(univerAPI, onRibbonAction) {
  const icons = registerRibbonIcons(univerAPI)
  const drawingComponents = registerDrawingComponents(univerAPI)
  const tabs = [
    {
      id: 'file', title: 'File', order: 0, groups: [
        { id: 'workbook', title: 'Workbook', items: [
          { id: 'new-workbook', title: 'New Workbook', icon: 'home', action: 'new-workbook' },
          { id: 'open-workbook', title: 'Browse Workbooks', icon: 'home', action: 'browse-workbooks' },
          { id: 'import-workbook', title: 'Import Excel File', icon: 'import', action: 'import' },
          { id: 'save', title: 'Save', icon: 'save', action: 'save' },
          { id: 'save-as', title: 'Save As', icon: 'copy', action: 'save-as' },
          { id: 'export-xlsx', title: 'Export Excel', icon: 'export', action: 'export-xlsx' },
        ] },
        { id: 'information', title: 'Workbook Information', items: [
          { id: 'copy-id', title: 'Copy Workbook ID', action: 'copy-workbook-id' },
        ] },
        { id: 'print', title: 'Print', items: [
          { id: 'print-workbook', title: 'Print', icon: 'print', action: 'print' },
        ] },
      ],
    },
    {
      id: 'share', title: 'Share', order: 4, groups: [
        { id: 'share-copy', title: 'Share & Access', items: [
          { id: 'share-file', title: 'Share Workbook File', icon: 'share', action: 'share-workbook' },
          { id: 'copy-link', title: 'Copy Workbook Link', icon: 'link', action: 'copy-workbook-link' },
          { id: 'manage-access', title: 'Manage Access (Unavailable)', action: 'manage-access-unavailable' },
        ] },
      ],
    },
    {
      id: 'page-layout', title: 'Page Layout', order: 5, groups: [
        { id: 'page-setup', title: 'Page Setup', items: [
          { id: 'print-layout', title: 'Page Setup', icon: 'layout', command: 'sheet.operation.open-print-grid-sidebar' },
          { id: 'print', title: 'Print', icon: 'print', action: 'print' },
          { id: 'orientation', title: 'Page Orientation (Unavailable)', action: 'page-layout-unavailable' },
          { id: 'paper-size', title: 'Paper Size (Unavailable)', action: 'page-layout-unavailable' },
          { id: 'margins', title: 'Margins (Unavailable)', action: 'page-layout-unavailable' },
          { id: 'print-area', title: 'Print Area (Unavailable)', action: 'page-layout-unavailable' },
          { id: 'page-breaks', title: 'Page Breaks (Unavailable)', action: 'page-layout-unavailable' },
        ] },
        { id: 'scale-to-fit', title: 'Scale to Fit', items: [
          { id: 'fit-width', title: 'Fit to Width (Unavailable)', action: 'page-layout-unavailable' },
          { id: 'fit-height', title: 'Fit to Height (Unavailable)', action: 'page-layout-unavailable' },
        ] },
        { id: 'sheet-background', title: 'Sheet Background', items: [
          { id: 'background-image', title: 'Background Image (Unavailable)', action: 'page-layout-unavailable' },
        ] },
      ],
    },
    {
      id: 'review', title: 'Review', order: 8, groups: [
        { id: 'comments', title: 'Comments & Notes', items: [
          { id: 'new-comment', title: 'New Comment', icon: 'review', command: 'sheet.operation.show-comment-modal' },
          { id: 'new-note', title: 'New Note', action: 'open-note' },
          { id: 'comment-panel', title: 'Comment Management', command: 'sheet.operation.toggle-comment-panel' },
          { id: 'spelling', title: 'Spelling (Unavailable)', action: 'review-tools-unavailable' },
        ] },
        { id: 'protection', title: 'Protection', items: [
          { id: 'protect-range', title: 'Protect Selected Range', command: 'sheet.command.add-range-protection-from-toolbar' },
          { id: 'protect-sheet-workbook-unavailable', title: 'Protect Sheet / Workbook (Unavailable)', action: 'sheet-protection-unavailable' },
          { id: 'show-changes', title: 'Show Changes (Unavailable)', action: 'review-tools-unavailable' },
        ] },
        { id: 'history', title: 'History', items: [
          { id: 'edit-history', title: 'Edit History', command: 'univer.operation.toggle-edit-history' },
        ] },
      ],
    },
    {
      id: 'help', title: 'Help', order: 10, groups: [
        { id: 'support', title: 'Help & Support', items: [
          { id: 'help-center', title: 'Help Center', icon: 'help', action: 'open-help' },
          { id: 'shortcuts', title: 'Keyboard Shortcuts', action: 'open-shortcuts' },
          { id: 'docs', title: 'Function Guide', icon: 'guide', action: 'open-help' },
        ] },
      ],
    },
    {
      id: 'draw', title: 'Draw', order: 11, groups: [
        { id: 'drawing', title: 'Drawing Tools', items: [
          { id: 'draw-tools-unavailable', title: 'Pen / Highlighter / Eraser (Unavailable)', action: 'draw-tools-unavailable' },
        ] },
      ],
    },
    {
      id: 'shape', title: 'Shape', order: 12, groups: [
        { id: 'insert-shape', title: 'Insert Shape', items: [
          { id: 'rectangle', title: 'Rectangle', icon: 'draw', run: (api) => insertDrawingShape(api, 'rectangle') },
          { id: 'ellipse', title: 'Ellipse', icon: 'draw', run: (api) => insertDrawingShape(api, 'ellipse') },
          { id: 'line', title: 'Line', icon: 'draw', run: (api) => insertDrawingShape(api, 'line') },
        ] },
        { id: 'shape-format', title: 'Shape Formatting', items: [
          { id: 'shape-format-unavailable', title: 'Fill / Outline / Properties (Unavailable)', action: 'shape-format-unavailable' },
        ] },
      ],
    },
  ]

  // Univer's Facade supports adding groups to built-in tabs and root tabs via appendTo('ribbon').
  const menus = tabs.map((tab) => addRibbonTab(univerAPI, tab, onRibbonAction))

  const clipboardGroup = univerAPI.createSubmenu({
    id: 'gridnova.ribbon.home.clipboard',
    title: 'Clipboard',
    order: 1,
  })
  ;[
    { id: 'paste', title: 'Paste', icon: 'paste', command: 'sheet.command.paste', gridLayout: { row: 1, column: 1, rowSpan: 2, width: 54, showLabel: true } },
    { id: 'cut', title: 'Cut', icon: 'cut', command: 'sheet.command.cut', gridLayout: { row: 1, column: 2 } },
    { id: 'copy', title: 'Copy', icon: 'copy', command: 'sheet.command.copy', gridLayout: { row: 2, column: 2 } },
    { id: 'format-painter', title: 'Format Painter', icon: 'paint', command: 'sheet.command.set-once-format-painter', gridLayout: { row: 1, column: 3 } },
    { id: 'paste-special', title: 'Paste Special (Unavailable)', icon: 'paste', action: 'paste-special-unavailable', gridLayout: { row: 2, column: 3 } },
  ].forEach((item) => {
    clipboardGroup.addSubmenu(univerAPI.createMenu({
      id: 'gridnova.ribbon.home.clipboard.' + item.id,
      title: item.title,
      tooltip: item.title,
      icon: item.icon ? 'gridnova-icon-' + item.icon : undefined,
      gridLayout: item.gridLayout,
      action: () => univerAPI.executeCommand(item.command),
    }))
  })
  clipboardGroup.appendTo('ribbon.start')

  const homeGroups = [
    {
      id: 'styles', title: 'Styles', order: 5, items: [
        { id: 'conditional-formatting', title: 'Conditional Formatting', icon: 'layout', gridLayout: { row: 1, column: 1 }, command: 'sheet.operation.open.conditional.formatting.panel' },
        { id: 'format-as-table', title: 'Format as Table', icon: 'table', gridLayout: { row: 1, column: 2 }, run: formatSelectionAsTable },
        { id: 'cell-styles', title: 'Cell Styles (Unavailable)', icon: 'paint', action: 'cell-styles-unavailable' },
        { id: 'format-cells', title: 'Format Cells (Unavailable)', icon: 'layout', action: 'format-cells-unavailable' },
      ],
    },
    {
      id: 'cells', title: 'Cells', order: 6, items: [
        { id: 'insert-row', title: 'Insert Row', icon: 'table', gridLayout: { row: 1, column: 1 }, command: 'sheet.command.insert-row-before' },
        { id: 'insert-column', title: 'Insert Column', icon: 'table', gridLayout: { row: 1, column: 2 }, command: 'sheet.command.insert-col-before' },
        { id: 'delete-row', title: 'Delete Row', icon: 'cut', gridLayout: { row: 1, column: 3 }, command: 'sheet.command.remove-row-confirm' },
        { id: 'delete-column', title: 'Delete Column', icon: 'cut', gridLayout: { row: 1, column: 4 }, command: 'sheet.command.remove-col-confirm' },
      ],
    },
    {
      id: 'editing', title: 'Editing', order: 7, items: [
        { id: 'sort-ascending', title: 'Sort A to Z', icon: 'sort', gridLayout: { row: 1, column: 1 }, command: 'sheet.command.sort-range-asc' },
        { id: 'sort-descending', title: 'Sort Z to A', icon: 'sort', gridLayout: { row: 1, column: 2 }, command: 'sheet.command.sort-range-desc' },
        { id: 'filter', title: 'Filter', icon: 'filter', gridLayout: { row: 1, column: 3 }, command: 'sheet.command.smart-toggle-filter' },
        { id: 'clear-contents', title: 'Clear Contents', icon: 'table', gridLayout: { row: 1, column: 4 }, command: 'sheet.command.clear-selection-content' },
        { id: 'clear-formats', title: 'Clear Formats', icon: 'paint', gridLayout: { row: 1, column: 5 }, command: 'sheet.command.clear-selection-format' },
        { id: 'clear-all', title: 'Clear All', icon: 'table', gridLayout: { row: 1, column: 6 }, command: 'sheet.command.clear-selection-all' },
        { id: 'find', title: 'Find & Replace', icon: 'search', gridLayout: { row: 1, column: 7 }, command: 'ui.operation.open-find-dialog' },
      ],
    },
    {
      id: 'sheet', title: 'Sheet', order: 8, items: [
        { id: 'rename-sheet', title: 'Rename Sheet', icon: 'table', gridLayout: { row: 1, column: 1 }, run: renameActiveSheet },
        { id: 'duplicate-sheet', title: 'Duplicate Sheet', icon: 'copy', gridLayout: { row: 1, column: 2 }, run: duplicateActiveSheet },
        { id: 'delete-sheet', title: 'Delete Sheet', icon: 'cut', gridLayout: { row: 1, column: 3 }, run: deleteActiveSheet },
        { id: 'hide-sheet', title: 'Hide Sheet', icon: 'layout', gridLayout: { row: 1, column: 4 }, run: hideActiveSheet },
        { id: 'unhide-sheet', title: 'Unhide Sheet', icon: 'layout', gridLayout: { row: 1, column: 5 }, run: unhideWorksheet },
        { id: 'change-tab-color', title: 'Sheet Tab Color', icon: 'fillColor', gridLayout: { row: 1, column: 6 }, run: setActiveSheetTabColor },
      ],
    },
  ]

  const nativeTabGroups = [
    {
      path: 'ribbon.insert', id: 'tables', title: 'Tables', order: 0, items: [
        { id: 'insert-table', title: 'Format as Table', icon: 'table', run: formatSelectionAsTable },
        { id: 'insert-pivot-summary', title: 'Pivot Summary Tools', action: 'insert-tools' },
      ],
    },
    {
      path: 'ribbon.insert', id: 'illustrations', title: 'Illustrations', order: 1, items: [
        { id: 'insert-picture', title: 'Pictures', icon: 'image', command: 'sheet.command.insert-float-image' },
      ],
    },
    {
      path: 'ribbon.insert', id: 'charts', title: 'Charts', order: 2, items: [
        { id: 'insert-chart', title: 'Quick Bar Chart Overlay', icon: 'table', run: insertBasicChart },
        { id: 'chart-pivot-tools', title: 'Chart & Pivot Summary Tools', icon: 'table', action: 'insert-tools' },
      ],
    },
    {
      path: 'ribbon.insert', id: 'links', title: 'Links', order: 3, items: [
        { id: 'insert-link', title: 'Hyperlink', icon: 'link', command: 'sheet.operation.insert-hyper-link-toolbar' },
      ],
    },
    {
      path: 'ribbon.insert', id: 'text-symbols', title: 'Text & Symbols', order: 4, items: [
        { id: 'text-icons-unavailable', title: 'Text Boxes / Icons (Unavailable)', action: 'insert-text-unavailable' },
      ],
    },
    {
      path: 'ribbon.formulas', id: 'function-library', title: 'Function Library', order: 0, items: [
        { id: 'insert-function', title: 'Insert Function', icon: 'calculator', run: (api) => insertFormulaFunction(api, 'SUM') },
        { id: 'auto-sum', title: 'AutoSum', icon: 'calculator', run: (api) => insertFormulaFunction(api, 'SUM') },
      ],
    },
    {
      path: 'ribbon.formulas', id: 'financial-functions', title: 'Financial', order: 1, items: [
        { id: 'function-pmt', title: 'PMT', run: (api) => insertFormulaFunction(api, 'PMT') },
        { id: 'function-fv', title: 'FV', run: (api) => insertFormulaFunction(api, 'FV') },
        { id: 'function-npv', title: 'NPV', run: (api) => insertFormulaFunction(api, 'NPV') },
      ],
    },
    {
      path: 'ribbon.formulas', id: 'logical-functions', title: 'Logical', order: 2, items: [
        { id: 'function-if', title: 'IF', run: (api) => insertFormulaFunction(api, 'IF') },
        { id: 'function-and', title: 'AND', run: (api) => insertFormulaFunction(api, 'AND') },
        { id: 'function-or', title: 'OR', run: (api) => insertFormulaFunction(api, 'OR') },
      ],
    },
    {
      path: 'ribbon.formulas', id: 'text-functions', title: 'Text', order: 3, items: [
        { id: 'function-left', title: 'LEFT', run: (api) => insertFormulaFunction(api, 'LEFT') },
        { id: 'function-right', title: 'RIGHT', run: (api) => insertFormulaFunction(api, 'RIGHT') },
        { id: 'function-len', title: 'LEN', run: (api) => insertFormulaFunction(api, 'LEN') },
        { id: 'function-concat', title: 'CONCAT', run: (api) => insertFormulaFunction(api, 'CONCAT') },
      ],
    },
    {
      path: 'ribbon.formulas', id: 'date-time-functions', title: 'Date & Time', order: 4, items: [
        { id: 'function-today', title: 'TODAY', run: (api) => insertFormulaFunction(api, 'TODAY') },
        { id: 'function-now', title: 'NOW', run: (api) => insertFormulaFunction(api, 'NOW') },
        { id: 'function-date', title: 'DATE', run: (api) => insertFormulaFunction(api, 'DATE') },
      ],
    },
    {
      path: 'ribbon.formulas', id: 'lookup-functions', title: 'Lookup & Reference', order: 5, items: [
        { id: 'function-vlookup', title: 'VLOOKUP', run: (api) => insertFormulaFunction(api, 'VLOOKUP') },
        { id: 'function-index', title: 'INDEX', run: (api) => insertFormulaFunction(api, 'INDEX') },
        { id: 'function-match', title: 'MATCH', run: (api) => insertFormulaFunction(api, 'MATCH') },
      ],
    },
    {
      path: 'ribbon.formulas', id: 'math-functions', title: 'Math & Trigonometry', order: 6, items: [
        { id: 'function-sum', title: 'SUM', run: (api) => insertFormulaFunction(api, 'SUM') },
        { id: 'function-round', title: 'ROUND', run: (api) => insertFormulaFunction(api, 'ROUND') },
        { id: 'function-abs', title: 'ABS', run: (api) => insertFormulaFunction(api, 'ABS') },
      ],
    },
    {
      path: 'ribbon.formulas', id: 'statistical-functions', title: 'Statistical', order: 7, items: [
        { id: 'function-average', title: 'AVERAGE', run: (api) => insertFormulaFunction(api, 'AVERAGE') },
        { id: 'function-count', title: 'COUNT', run: (api) => insertFormulaFunction(api, 'COUNT') },
        { id: 'function-max', title: 'MAX', run: (api) => insertFormulaFunction(api, 'MAX') },
        { id: 'function-min', title: 'MIN', run: (api) => insertFormulaFunction(api, 'MIN') },
      ],
    },
    {
      path: 'ribbon.formulas', id: 'calculation', title: 'Calculation', order: 8, items: [
        { id: 'recalculate', title: 'Recalculate Now', icon: 'calculator', run: recalculateFormulas },
      ],
    },
    {
      path: 'ribbon.data', id: 'sort-filter', title: 'Sort & Filter', order: 0, items: [
        { id: 'sort-ascending', title: 'Sort A to Z', icon: 'sort', command: 'sheet.command.sort-range-asc' },
        { id: 'sort-descending', title: 'Sort Z to A', icon: 'sort', command: 'sheet.command.sort-range-desc' },
        { id: 'custom-sort', title: 'Custom Sort', command: 'sheet.command.sort-range-custom' },
        { id: 'toggle-filter', title: 'Filter', icon: 'filter', command: 'sheet.command.smart-toggle-filter' },
        { id: 'clear-filter', title: 'Clear Filter', command: 'sheet.command.clear-filter-criteria' },
      ],
    },
    {
      path: 'ribbon.data', id: 'data-tools', title: 'Data Tools', order: 1, items: [
        { id: 'data-validation', title: 'Data Validation', command: 'data-validation.operation.open-validation-panel' },
        { id: 'remove-duplicates', title: 'Remove Duplicates', icon: 'table', run: removeDuplicatesFromSelection },
        { id: 'text-to-columns', title: 'Text to Columns', icon: 'table', run: splitTextToColumns },
      ],
    },
    {
      path: 'ribbon.data', id: 'outline', title: 'Row Visibility', order: 2, items: [
        { id: 'hide-selected-rows', title: 'Hide Selected Rows', run: hideSelectedRows },
        { id: 'show-selected-rows', title: 'Unhide Selected Rows', run: showSelectedRows },
      ],
    },
    {
      path: 'ribbon.data', id: 'group-subtotal', title: 'Grouping & Subtotals', order: 3, items: [
        { id: 'group-subtotal-unavailable', title: 'Group / Ungroup / Subtotal (Unavailable)', action: 'data-outline-unavailable' },
      ],
    },
    {
      path: 'ribbon.view', id: 'show', title: 'Show', order: 0, items: [
        { id: 'toggle-gridlines', title: 'Gridlines', icon: 'layout', run: toggleGridlines },
        { id: 'formula-bar-unavailable', title: 'Formula Bar (Unavailable)', action: 'formula-bar-unavailable' },
        { id: 'normal-view', title: 'Normal View (Current)', action: 'normal-view-current' },
        { id: 'page-layout-view', title: 'Page Layout View (Unavailable)', action: 'view-tools-unavailable' },
      ],
    },
    {
      path: 'ribbon.view', id: 'panes', title: 'Freeze Panes', order: 1, items: [
        { id: 'freeze-first-row', title: 'Freeze Top Row', run: freezeFirstRow },
        { id: 'freeze-first-column', title: 'Freeze First Column', run: freezeFirstColumn },
        { id: 'unfreeze-panes', title: 'Unfreeze Panes', run: unfreezePanes },
      ],
    },
    {
      path: 'ribbon.view', id: 'window', title: 'Zoom & Window', order: 2, items: [
        { id: 'zoom-out', title: 'Zoom Out', icon: 'calculator', run: (api) => setSheetZoom(api, Math.max(50, Math.round(api.getActiveSheet()?.getZoom?.() * 100 || 100) - 10)) },
        { id: 'zoom', title: 'Zoom Percentage', run: (api) => setSheetZoom(api) },
        { id: 'zoom-in', title: 'Zoom In', icon: 'calculator', run: (api) => setSheetZoom(api, Math.min(200, Math.round(api.getActiveSheet()?.getZoom?.() * 100 || 100) + 10)) },
        { id: 'zoom-100', title: '100% Zoom', run: (api) => setSheetZoom(api, 100) },
        { id: 'fullscreen', title: 'Fullscreen', run: toggleFullscreen },
      ],
    },
  ]

  const nativeTabGroupMenus = nativeTabGroups.map((group) => {
    const groupMenu = univerAPI.createSubmenu({
      id: 'gridnova.ribbon.' + group.path.replaceAll('.', '-') + '.' + group.id,
      title: group.title,
      order: group.order,
    })
    group.items.forEach((item) => {
      groupMenu.addSubmenu(univerAPI.createMenu({
        id: 'gridnova.ribbon.' + group.id + '.' + item.id,
        title: item.title,
        tooltip: item.title,
        icon: item.icon ? 'gridnova-icon-' + item.icon : undefined,
        gridLayout: item.gridLayout,
        action: () => item.run ? item.run(univerAPI) : item.action ? onRibbonAction?.(item.action) : univerAPI.executeCommand(item.command),
      }))
    })
    groupMenu.appendTo(group.path)
    return groupMenu
  })

  const homeGroupMenus = homeGroups.map((group) => {
    const groupMenu = univerAPI.createSubmenu({
      id: 'gridnova.ribbon.home.' + group.id,
      title: group.title,
      order: group.order,
    })
    group.items.forEach((item) => {
      groupMenu.addSubmenu(univerAPI.createMenu({
        id: 'gridnova.ribbon.home.' + item.id,
        title: item.title,
        tooltip: item.title,
        icon: item.icon ? 'gridnova-icon-' + item.icon : undefined,
        gridLayout: item.gridLayout,
        action: () => item.run ? item.run(univerAPI) : item.action ? onRibbonAction?.(item.action) : univerAPI.executeCommand(item.command),
      }))
    })
    groupMenu.appendTo('ribbon.start')
    return groupMenu
  })

  univerAPI.updateMenuConfig({
    'ribbon.start': { title: 'Home', order: 1 },
    'ribbon.insert': { title: 'Insert', order: 2, hidden: false },
    'ribbon.start.history': { hidden: true },
    'ribbon.start.format': { title: 'Font', order: 2 },
    'ribbon.start.layout': { title: 'Alignment', order: 3 },
    'ribbon.start.others': { title: 'More Editing', order: 8 },
    'univer.command.undo': { icon: 'gridnova-icon-undo' },
    'univer.command.redo': { icon: 'gridnova-icon-redo' },
    'ui.operation.activate-format-painter': { icon: 'gridnova-icon-paint' },
    'sheet.command.set-once-format-painter': { icon: 'gridnova-icon-paint' },
    'sheet.command.set-range-bold': { icon: 'gridnova-icon-bold', gridLayout: { row: 1, column: 3 } },
    'sheet.command.set-range-italic': { icon: 'gridnova-icon-italic', gridLayout: { row: 1, column: 4 } },
    // Keep Home controls in a single horizontal command row in grid (Single Line) mode.
    // Explicit gridLayout coordinates prevent controls from being collapsed into
    // unlabeled chevrons or pushed into unexpected empty grid cells.
    'sheet.command.set-range-font-family': { order: 0, gridLayout: { row: 1, column: 1, width: 150, showLabel: true } },
    'sheet.command.set-range-fontsize': { order: 1, gridLayout: { row: 1, column: 2, width: 58, showLabel: true } },
    'sheet.command.add-worksheet-merge': { icon: 'gridnova-icon-layout', gridLayout: { row: 1, column: 4 } },
    'sheet.command.add-worksheet-merge-all': { icon: 'gridnova-icon-layout', gridLayout: { row: 1, column: 5 } },
    'sheet.command.add-worksheet-merge-vertical': { icon: 'gridnova-icon-layout', gridLayout: { row: 1, column: 6 } },
    'sheet.command.add-worksheet-merge-horizontal': { icon: 'gridnova-icon-layout', gridLayout: { row: 1, column: 7 } },
    'sheet.command.remove-worksheet-merge': { icon: 'gridnova-icon-layout', gridLayout: { row: 1, column: 8 } },
    'sheet.command.set-text-rotation': { icon: 'gridnova-icon-layout', gridLayout: { row: 1, column: 9 } },
    'sheet.command.set-range-underline': { icon: 'gridnova-icon-underline', gridLayout: { row: 1, column: 5 } },
    'sheet.command.set-range-stroke': { icon: 'gridnova-icon-strike', gridLayout: { row: 1, column: 6 } },
    'sheet.command.set-range-text-color': { icon: 'gridnova-icon-textColor', gridLayout: { row: 1, column: 7 } },
    'sheet.command.set-background-color': { icon: 'gridnova-icon-fillColor', gridLayout: { row: 1, column: 8 } },
    'sheet.command.set-border-basic': { icon: 'gridnova-icon-border', gridLayout: { row: 1, column: 9 } },
    'sheet.command.set-horizontal-text-align': { icon: 'gridnova-icon-alignLeft', gridLayout: { row: 1, column: 1 } },
    'sheet.command.set-vertical-text-align': { icon: 'gridnova-icon-alignCenter', gridLayout: { row: 1, column: 2 } },
    'sheet.command.set-text-wrap': { icon: 'gridnova-icon-alignRight', gridLayout: { row: 1, column: 3 } },
    'sheet.command.copy': { icon: 'gridnova-icon-copy' },
    'sheet.command.cut': { icon: 'gridnova-icon-cut' },
    'sheet.command.paste': { icon: 'gridnova-icon-paste' },
    'sheet.command.smart-toggle-filter': { icon: 'gridnova-icon-filter' },
    'sheet.command.sort-range-asc': { icon: 'gridnova-icon-sort' },
    'ui.operation.open-find-dialog': { icon: 'gridnova-icon-search' },
    'sheet.command.numfmt.set.currency': { icon: 'gridnova-icon-currency', gridLayout: { row: 1, column: 2 } },
    'sheet.command.numfmt.set.percent': { icon: 'gridnova-icon-percent', gridLayout: { row: 1, column: 3 } },
    'sheet.command.numfmt.add.decimal.command': { icon: 'gridnova-icon-calculator', gridLayout: { row: 1, column: 4 } },
    'sheet.command.numfmt.subtract.decimal.command': { icon: 'gridnova-icon-calculator', gridLayout: { row: 1, column: 5 } },
    'sheet.operation.open.numfmt.panel': { icon: 'gridnova-icon-calculator', gridLayout: { row: 1, column: 1, width: 112, showLabel: true } },
    'sheet.command.insert-row-before': { icon: 'gridnova-icon-table' },
    'sheet.command.insert-col-before': { icon: 'gridnova-icon-table' },
    'sheet.command.remove-row-confirm': { icon: 'gridnova-icon-table' },
    'sheet.command.remove-col-confirm': { icon: 'gridnova-icon-table' },
    'sheet.menu.image': { icon: 'gridnova-icon-image' },
    'sheet.command.insert-float-image': { icon: 'gridnova-icon-image' },
    'sheet.operation.insert-hyper-link-toolbar': { icon: 'gridnova-icon-link' },
    'sheet.operation.print-open': { icon: 'gridnova-icon-print' },
    'sheet.command.add-range-protection-from-toolbar': { icon: 'gridnova-icon-review' },
  })

  return [...icons, ...drawingComponents, ...menus, clipboardGroup, ...homeGroupMenus, ...nativeTabGroupMenus]
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

export default function UniverEditor({ workbookId, title, initialSnapshot, ribbonType = 'classic', onReady, onRibbonAction }) {
  const containerRef = useRef(null)
  const onReadyRef = useRef(onReady)
  const onRibbonActionRef = useRef(onRibbonAction)
  const initialConfigRef = useRef({ workbookId, title, initialSnapshot, ribbonType })
  initialConfigRef.current = { workbookId, title, initialSnapshot, ribbonType }
  const [error, setError] = useState('')

  useEffect(() => {
    onReadyRef.current = onReady
    onRibbonActionRef.current = onRibbonAction
  }, [onReady, onRibbonAction])

  useEffect(() => {
    if (!containerRef.current) return undefined

    let univerInstance
    let ribbonDisposables = []
    let disposed = false

    try {
      const config = initialConfigRef.current
      const { univer, univerAPI } = createUniver({
        locale: LocaleType.EN_US,
        locales: {
          [LocaleType.EN_US]: mergeLocaleMessages(
            UniverPresetSheetsCoreEnUS,
            UniverPresetSheetsDrawingEnUS,
            UniverPresetSheetsTableEnUS,
            UniverPresetSheetsThreadCommentEnUS,
            UniverPresetSheetsNoteEnUS,
            UniverPresetSheetsDataValidationEnUS,
            {
              ui: {
                ribbon: {
                  start: 'Home',
                },
              },
              gridnova: {
                ribbon: {
                  tabs: {
                    file: 'File',
                    share: 'Share',
                    'page-layout': 'Page Layout',
                    review: 'Review',
                    help: 'Help',
                    draw: 'Draw',
                    shape: 'Shape',
                  },
                  groups: {
                    file: {
                      workbook: 'Workbook',
                      information: 'Workbook Information',
                      print: 'Print',
                    },
                    share: {
                      'share-copy': 'Share & Access',
                    },
                    'page-layout': {
                      'page-setup': 'Page Setup',
                      'scale-to-fit': 'Scale to Fit',
                      'sheet-background': 'Sheet Background',
                    },
                    review: {
                      comments: 'Comments & Notes',
                      protection: 'Protection',
                      history: 'History',
                    },
                    help: {
                      support: 'Help & Support',
                    },
                    draw: {
                      drawing: 'Drawing Tools',
                    },
                    shape: {
                      'insert-shape': 'Insert Shape',
                      'shape-format': 'Shape Formatting',
                    },
                  },
                  home: {
                    clipboard: 'Clipboard',
                    styles: 'Styles',
                    cells: 'Cells',
                    editing: 'Editing',
                    sheet: 'Sheet',
                  },
                  'ribbon-insert': {
                    tables: 'Tables',
                    illustrations: 'Illustrations',
                    charts: 'Charts',
                    links: 'Links',
                    'text-symbols': 'Text & Symbols',
                  },
                  'ribbon-formulas': {
                    'function-library': 'Function Library',
                    'financial-functions': 'Financial',
                    'logical-functions': 'Logical',
                    'text-functions': 'Text',
                    'date-time-functions': 'Date & Time',
                    'lookup-functions': 'Lookup & Reference',
                    'math-functions': 'Math & Trigonometry',
                    'statistical-functions': 'Statistical',
                    calculation: 'Calculation',
                  },
                  'ribbon-data': {
                    'sort-filter': 'Sort & Filter',
                    'data-tools': 'Data Tools',
                    outline: 'Row Visibility',
                    'group-subtotal': 'Grouping & Subtotals',
                  },
                  'ribbon-view': {
                    show: 'Show',
                    panes: 'Freeze Panes',
                    window: 'Zoom & Window',
                  },
                },
              },
            },
          )
        },
        presets: [
          UniverSheetsCorePreset({
            container: containerRef.current,
            header: true,
            toolbar: true,
            contextMenu: true,
            ribbonType: config.ribbonType,
            menu: {
              'ribbon.start': { title: 'Home', order: 1 },
              'ribbon.insert': { title: 'Insert', order: 2, hidden: false },
              'ribbon.formulas': { title: 'Formulas', order: 6 },
              'ribbon.data': { title: 'Data', order: 7 },
              'ribbon.view': { title: 'View', order: 9 },
            },
            footer: {
              addSheetButtonConfig: {
                show: true,
                defaultRowCount: DEFAULT_ROW_COUNT,
                defaultColumnCount: DEFAULT_COLUMN_COUNT,
              },
            },
          }),
          UniverSheetsDrawingPreset(),
          UniverSheetsTablePreset(),
          UniverSheetsThreadCommentPreset(),
          UniverSheetsNotePreset(),
          UniverSheetsDataValidationPreset(),
        ],
      })

      univerInstance = univer
      const workbookData = config.initialSnapshot
        ? ensureMinimumColumnCount(config.initialSnapshot)
        : createBlankWorkbookSnapshot(config.workbookId, config.title)

      univerAPI.createWorkbook(workbookData)
      addColumnHeaderDeleteAction(univerAPI)
      ribbonDisposables = addMicrosoftInspiredRibbonTabs(univerAPI, (action) => onRibbonActionRef.current?.(action))

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
        getWorkbook(id) {
          return univerAPI.getWorkbook(id)
        },
        executeCommand(commandId, params) {
          return univerAPI.executeCommand(commandId, params)
        },
        onCommandExecuted(callback) {
          return univerAPI.addEvent(univerAPI.Event.CommandExecuted, callback)
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
        ribbonDisposables?.forEach((item) => item?.dispose?.())
        univerInstance?.dispose()
      } catch (disposeError) {
        console.error('GridNova editor cleanup failed:', disposeError)
      }
    }
  }, [workbookId, initialSnapshot, ribbonType])

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
