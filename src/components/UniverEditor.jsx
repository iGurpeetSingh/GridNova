import { useEffect, useRef, useState } from 'react'
import { ArrowDownloadRegular, ArrowRedoRegular, ArrowUndoRegular, ArrowUploadRegular, BookOpenRegular, BorderAllRegular, CalculatorRegular, ClipboardPasteRegular, CommentRegular, CopyRegular, CurrencyDollarRegular, CutRegular, DrawShapeRegular, FilterRegular, GridRegular, ImageRegular, LinkRegular, PaintBucketRegular, PaintBrushRegular, PercentRegular, PrintRegular, QuestionCircleRegular, SaveRegular, SearchRegular, ShareRegular, SortAscendingRegular, TableRegular, TextAlignCenterRegular, TextAlignLeftRegular, TextAlignRightRegular, TextBoldRegular, TextColorRegular, TextItalicRegular, TextStrikethroughRegular, TextUnderlineRegular } from '@fluentui/react-icons'
import { createUniver, LocaleType, mergeLocales } from '@univerjs/presets'
import { UniverSheetsCorePreset } from '@univerjs/preset-sheets-core'
import '@univerjs/ui/facade'
import '@univerjs/sheets/facade'
import UniverPresetSheetsCoreEnUS from '@univerjs/preset-sheets-core/locales/en-US'
import { UniverSheetsDrawingPreset } from '@univerjs/preset-sheets-drawing'
import UniverPresetSheetsDrawingEnUS from '@univerjs/preset-sheets-drawing/locales/en-US'
import '@univerjs/preset-sheets-core/lib/index.css'
import '@univerjs/preset-sheets-drawing/lib/index.css'

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
  sort: SortAscendingRegular,
  search: SearchRegular,
  calculator: CalculatorRegular,
  currency: CurrencyDollarRegular,
  percent: PercentRegular,
  image: ImageRegular,
  link: LinkRegular,
  paint: PaintBrushRegular,
}

function registerRibbonIcons(univerAPI) {
  const disposables = []
  Object.entries(RIBBON_ICONS).forEach(([name, Icon]) => {
    disposables.push(univerAPI.registerComponent('gridnova-icon-' + name, Icon))
  })
  return disposables
}

function addRibbonTab(univerAPI, tab, onRibbonAction) {
  const root = univerAPI.createSubmenu({
    id: 'gridnova.ribbon.' + tab.id,
    title: tab.title,
    order: tab.order,
  })

  tab.groups.forEach((group) => {
    const groupMenu = univerAPI.createSubmenu({
      id: 'gridnova.ribbon.' + tab.id + '.' + group.id,
      title: group.title,
    })

    group.items.forEach((item) => {
      const action = item.run
        ? () => item.run(univerAPI)
        : item.command
          ? () => univerAPI.executeCommand(item.command, item.params)
          : () => onRibbonAction?.(item.action)
      const menu = univerAPI.createMenu({
        id: 'gridnova.ribbon.' + tab.id + '.' + item.id,
        title: item.title,
        tooltip: item.tooltip || item.title,
        icon: item.icon ? 'gridnova-icon-' + item.icon : undefined,
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
    { componentKey: 'gridnova-bar-chart', data: { labels: dataRows.map((row) => row.label), values: dataRows.map((row) => row.value) } },
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
    { componentKey },
    { width: 180, height: 100 },
    id,
  )
  if (!shape) {
    univerAPI.showMessage({ content: 'The shape could not be inserted.', type: 'error' })
    return
  }
  univerAPI.showMessage({ content: shapeType[0].toUpperCase() + shapeType.slice(1) + ' added over the selected range.', type: 'success' })
}

function formatSelectionAsTable(univerAPI) {
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
  sheet.getRange(range.getRow(), range.getColumn(), 1, columnCount)
    .setFontWeight('bold')
    .setBackgroundColor('#DCEAF7')
  for (let row = 1; row < rowCount; row += 2) {
    sheet.getRange(range.getRow() + row, range.getColumn(), 1, columnCount)
      .setBackgroundColor('#F4F8FC')
  }
  univerAPI.executeCommand('sheet.command.smart-toggle-filter')
  univerAPI.showMessage({ content: 'Table formatting and filter controls applied to the selected range.', type: 'success' })
}

function recalculateFormulas(univerAPI) {
  univerAPI.getFormula()?.executeCalculation()
  univerAPI.showMessage({ content: 'Formula recalculation started.', type: 'success' })
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
  const headerRows = hasHeader ? [values[0]] : []
  const seen = new Set()
  const uniqueRows = []

  values.slice(firstDataRow).forEach((row) => {
    const key = JSON.stringify(row)
    if (!seen.has(key)) {
      seen.add(key)
      uniqueRows.push(row)
    }
  })

  const removedCount = values.length - headerRows.length - uniqueRows.length
  if (removedCount <= 0) {
    univerAPI.showMessage({ content: 'No duplicate rows were found in the selected range.', type: 'info' })
    return
  }

  const columnCount = Math.max(1, ...values.map((row) => row.length))
  const output = [...headerRows, ...uniqueRows]
  while (output.length < values.length) output.push(Array(columnCount).fill(null))
  range.setValues(output)
  univerAPI.showMessage({ content: removedCount + ' duplicate row(s) cleared from the selected range.', type: 'success' })
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
  const delimiterNames = { comma: ',', semicolon: ';', pipe: '|', tab: '\\t' }
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
          { id: 'save', title: 'Save', icon: 'save', action: 'save' },
          { id: 'import', title: 'Import Excel', icon: 'import', action: 'import' },
          { id: 'export-xlsx', title: 'Export Excel', icon: 'export', action: 'export-xlsx' },
          { id: 'export-csv', title: 'Export CSV', action: 'export-csv' },
        ] },
        { id: 'print', title: 'Print', items: [
          { id: 'print-workbook', title: 'Print', icon: 'print', action: 'print' },
        ] },
      ],
    },
    {
      id: 'share', title: 'Share', order: 4, groups: [
        { id: 'share-copy', title: 'Share a Copy', items: [
          { id: 'share-file', title: 'Share Workbook File', icon: 'share', action: 'share-workbook' },
          { id: 'share-export', title: 'Export Workbook Copy', icon: 'export', action: 'export-xlsx' },
          { id: 'copy-id', title: 'Copy Workbook ID', action: 'copy-workbook-id' },
        ] },
      ],
    },
    {
      id: 'page-layout', title: 'Page Layout', order: 5, groups: [
        { id: 'page-setup', title: 'Page Setup', items: [
          { id: 'print-layout', title: 'Print Layout', icon: 'layout', action: 'print' },
          { id: 'print', title: 'Print', icon: 'print', action: 'print' },
        ] },
        { id: 'sheet-options', title: 'Sheet Options', items: [
          { id: 'gridlines', title: 'Toggle Gridlines', icon: 'layout', run: toggleGridlines },
          { id: 'freeze-row', title: 'Freeze First Row', run: freezeFirstRow },
          { id: 'freeze-column', title: 'Freeze First Column', run: freezeFirstColumn },
        ] },
      ],
    },
    {
      id: 'review', title: 'Review', order: 8, groups: [
        { id: 'comments', title: 'Comments & Notes', items: [
          { id: 'new-comment', title: 'New Comment', icon: 'review', action: 'open-comments' },
          { id: 'new-note', title: 'New Note', action: 'open-note' },
          { id: 'comment-panel', title: 'Comment Management', action: 'open-comments' },
        ] },
        { id: 'protection', title: 'Protection', items: [
          { id: 'protect-range', title: 'Protect Selected Range', command: 'sheet.command.add-range-protection-from-toolbar' },
        ] },
        { id: 'history', title: 'History', items: [
          { id: 'edit-history', title: 'Edit History', action: 'open-catch-up' },
        ] },
      ],
    },
    {
      id: 'help', title: 'Help', order: 10, groups: [
        { id: 'support', title: 'Help & Support', items: [
          { id: 'feature-search', title: 'Search Help', icon: 'help', action: 'open-help' },
          { id: 'shortcuts', title: 'Keyboard Shortcuts Guide', action: 'open-help' },
          { id: 'docs', title: 'GridNova Guide', icon: 'guide', action: 'open-help' },
        ] },
      ],
    },
    {
      id: 'draw', title: 'Draw', order: 11, groups: [
        { id: 'drawing', title: 'Drawing', items: [
          { id: 'rectangle', title: 'Rectangle', icon: 'draw', run: (api) => insertDrawingShape(api, 'rectangle') },
          { id: 'ellipse', title: 'Ellipse', icon: 'draw', run: (api) => insertDrawingShape(api, 'ellipse') },
          { id: 'line', title: 'Line', icon: 'draw', run: (api) => insertDrawingShape(api, 'line') },
        ] },
      ],
    },
  ]

  const menus = tabs.map((tab) => addRibbonTab(univerAPI, tab, onRibbonAction))

  const clipboardGroup = univerAPI.createSubmenu({
    id: 'gridnova.ribbon.home.clipboard',
    title: 'Clipboard',
    order: 1,
  })
  ;[
    { id: 'paste', title: 'Paste', icon: 'table', command: 'sheet.command.paste' },
    { id: 'cut', title: 'Cut', command: 'sheet.command.cut' },
    { id: 'copy', title: 'Copy', command: 'sheet.command.copy' },
    { id: 'format-painter', title: 'Format Painter', icon: 'layout', command: 'ui.operation.activate-format-painter' },
  ].forEach((item) => {
    clipboardGroup.addSubmenu(univerAPI.createMenu({
      id: 'gridnova.ribbon.home.clipboard.' + item.id,
      title: item.title,
      tooltip: item.title,
      icon: item.icon ? 'gridnova-icon-' + item.icon : undefined,
      action: () => univerAPI.executeCommand(item.command),
    }))
  })
  clipboardGroup.appendTo('ribbon.start')

  const homeGroups = [
    {
      id: 'number', title: 'Number', order: 4, items: [
        { id: 'number-format', title: 'Number Format', icon: 'table', command: 'sheet.operation.open.numfmt.panel' },
        { id: 'currency', title: 'Currency', command: 'sheet.command.numfmt.set.currency' },
        { id: 'percent', title: 'Percentage', command: 'sheet.command.numfmt.set.percent' },
        { id: 'increase-decimal', title: 'Increase Decimal', command: 'sheet.command.numfmt.add.decimal.command' },
        { id: 'decrease-decimal', title: 'Decrease Decimal', command: 'sheet.command.numfmt.subtract.decimal.command' },
      ],
    },
    {
      id: 'styles', title: 'Styles', order: 5, items: [
        { id: 'conditional-formatting', title: 'Conditional Formatting', icon: 'layout', command: 'sheet.operation.open.conditional.formatting.panel' },
        { id: 'data-bars', title: 'Data Bars', command: 'sheet.command.add-data-bar-conditional-rule' },
        { id: 'icon-sets', title: 'Icon Sets', command: 'sheet.command.add-icon-set-conditional-rule' },
      ],
    },
    {
      id: 'cells', title: 'Cells', order: 6, items: [
        { id: 'insert-row', title: 'Insert Row', command: 'sheet.command.insert-row-before' },
        { id: 'insert-column', title: 'Insert Column', command: 'sheet.command.insert-col-before' },
        { id: 'delete-row', title: 'Delete Row', command: 'sheet.command.remove-row-confirm' },
        { id: 'delete-column', title: 'Delete Column', command: 'sheet.command.remove-col-confirm' },
      ],
    },
    {
      id: 'editing', title: 'Editing', order: 7, items: [
        { id: 'sort-ascending', title: 'Sort A to Z', command: 'sheet.command.sort-range-asc' },
        { id: 'sort-descending', title: 'Sort Z to A', command: 'sheet.command.sort-range-desc' },
        { id: 'filter', title: 'Filter', icon: 'table', command: 'sheet.command.smart-toggle-filter' },
        { id: 'clear-contents', title: 'Clear Contents', command: 'sheet.command.clear-selection-content' },
        { id: 'find', title: 'Find & Replace', icon: 'help', command: 'ui.operation.open-find-dialog' },
      ],
    },
  ]

  const nativeTabGroups = [
    {
      path: 'ribbon.insert', id: 'tables', title: 'Tables', order: 0, items: [
        { id: 'insert-table', title: 'Format as Table', icon: 'table', run: formatSelectionAsTable },
        { id: 'insert-pivot-table', title: 'Pivot Summary Tools', action: 'insert-tools' },
      ],
    },
    {
      path: 'ribbon.insert', id: 'illustrations', title: 'Illustrations', order: 1, items: [
        { id: 'insert-picture', title: 'Pictures', command: 'sheet.command.insert-float-image' },
        { id: 'insert-shape', title: 'Rectangle', icon: 'draw', run: (api) => insertDrawingShape(api, 'rectangle') },
      ],
    },
    {
      path: 'ribbon.insert', id: 'charts', title: 'Charts', order: 2, items: [
        { id: 'insert-chart', title: 'Insert Bar Chart', icon: 'table', run: insertBasicChart },
        { id: 'chart-pivot-tools', title: 'Chart & Pivot Tools', icon: 'table', action: 'insert-tools' },
        { id: 'chart-pivot-tools', title: 'Chart & Pivot Tools', icon: 'table', action: 'insert-tools' },
      ],
    },
    {
      path: 'ribbon.insert', id: 'links', title: 'Links', order: 3, items: [
        { id: 'insert-link', title: 'Hyperlink', icon: 'share', command: 'sheet.operation.insert-hyper-link-toolbar' },
      ],
    },
    {
      path: 'ribbon.formulas', id: 'function-library', title: 'Function Library', order: 0, items: [
        { id: 'insert-function', title: 'Insert Function', command: 'formula-ui.operation.insert-function' },
        { id: 'more-functions', title: 'More Functions', command: 'formula-ui.operation.more-functions' },
      ],
    },
    {
      path: 'ribbon.formulas', id: 'calculation', title: 'Calculation', order: 1, items: [
        { id: 'recalculate', title: 'Recalculate Now', icon: 'table', run: recalculateFormulas },
      ],
    },
    {
      path: 'ribbon.data', id: 'sort-filter', title: 'Sort & Filter', order: 0, items: [
        { id: 'sort-ascending', title: 'Sort A to Z', command: 'sheet.command.sort-range-asc' },
        { id: 'sort-descending', title: 'Sort Z to A', command: 'sheet.command.sort-range-desc' },
        { id: 'custom-sort', title: 'Custom Sort', command: 'sheet.command.sort-range-custom' },
        { id: 'toggle-filter', title: 'Filter', icon: 'table', command: 'sheet.command.smart-toggle-filter' },
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
      path: 'ribbon.view', id: 'show', title: 'Show', order: 0, items: [
        { id: 'toggle-gridlines', title: 'Gridlines', icon: 'layout', run: toggleGridlines },
      ],
    },
    {
      path: 'ribbon.view', id: 'window', title: 'Window', order: 1, items: [
        { id: 'freeze-first-row', title: 'Freeze First Row', run: freezeFirstRow },
        { id: 'freeze-first-column', title: 'Freeze First Column', run: freezeFirstColumn },
        { id: 'unfreeze-panes', title: 'Unfreeze Panes', run: unfreezePanes },
        { id: 'zoom', title: 'Zoom', run: (api) => setSheetZoom(api) },
        { id: 'zoom-100', title: '100% Zoom', run: (api) => setSheetZoom(api, 100) },
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
        action: () => univerAPI.executeCommand(item.command),
      }))
    })
    groupMenu.appendTo('ribbon.start')
    return groupMenu
  })

  univerAPI.updateMenuConfig({
    'ribbon.start.history': { title: 'History', order: 0 },
    'ribbon.start.format': { title: 'Font', order: 2 },
    'ribbon.start.layout': { title: 'Alignment', order: 3 },
    'ribbon.start.others': { title: 'More Editing', order: 8 },
    'univer.command.undo': { icon: 'gridnova-icon-undo' },
    'univer.command.redo': { icon: 'gridnova-icon-redo' },
    'sheet.command.set-once-format-painter': { icon: 'gridnova-icon-paint' },
    'sheet.command.set-range-bold': { icon: 'gridnova-icon-bold' },
    'sheet.command.set-range-italic': { icon: 'gridnova-icon-italic' },
    'sheet.command.set-range-underline': { icon: 'gridnova-icon-underline' },
    'sheet.command.set-range-stroke': { icon: 'gridnova-icon-strike' },
    'sheet.command.set-range-text-color': { icon: 'gridnova-icon-textColor' },
    'sheet.command.set-background-color': { icon: 'gridnova-icon-fillColor' },
    'sheet.command.set-border-basic': { icon: 'gridnova-icon-border' },
    'sheet.command.set-horizontal-text-align': { icon: 'gridnova-icon-alignLeft' },
    'sheet.command.set-vertical-text-align': { icon: 'gridnova-icon-alignCenter' },
    'sheet.command.set-text-wrap': { icon: 'gridnova-icon-alignRight' },
    'sheet.command.copy': { icon: 'gridnova-icon-copy' },
    'sheet.command.cut': { icon: 'gridnova-icon-cut' },
    'sheet.command.paste': { icon: 'gridnova-icon-paste' },
    'sheet.command.smart-toggle-filter': { icon: 'gridnova-icon-filter' },
    'sheet.command.sort-range-asc': { icon: 'gridnova-icon-sort' },
    'ui.operation.open-find-dialog': { icon: 'gridnova-icon-search' },
    'sheet.command.numfmt.set.currency': { icon: 'gridnova-icon-currency' },
    'sheet.command.numfmt.set.percent': { icon: 'gridnova-icon-percent' },
    'sheet.operation.open.numfmt.panel': { icon: 'gridnova-icon-calculator' },
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
          [LocaleType.EN_US]: mergeLocales(UniverPresetSheetsCoreEnUS, UniverPresetSheetsDrawingEnUS),
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
              'ribbon.insert': { title: 'Insert', order: 2 },
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
