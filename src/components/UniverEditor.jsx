import { useEffect, useRef, useState } from 'react'
import { ArrowDownloadRegular, ArrowUploadRegular, BookOpenRegular, CommentRegular, DrawShapeRegular, GridRegular, PrintRegular, QuestionCircleRegular, SaveRegular, ShareRegular, TableRegular } from '@fluentui/react-icons'
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
      const action = item.command
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

function addMicrosoftInspiredRibbonTabs(univerAPI, onRibbonAction) {
  const icons = registerRibbonIcons(univerAPI)
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
          { id: 'share-export', title: 'Export Workbook Copy', icon: 'share', action: 'export-xlsx' },
          { id: 'copy-id', title: 'Copy Workbook ID', action: 'copy-workbook-id' },
        ] },
      ],
    },
    {
      id: 'page-layout', title: 'Page Layout', order: 5, groups: [
        { id: 'page-setup', title: 'Page Setup', items: [
          { id: 'print-layout', title: 'Print Layout', icon: 'layout', command: 'sheet.operation.open-print-grid-sidebar' },
          { id: 'print', title: 'Print', icon: 'print', action: 'print' },
        ] },
        { id: 'sheet-options', title: 'Sheet Options', items: [
          { id: 'gridlines', title: 'Toggle Gridlines', command: 'sheet.command.toggle-gridlines' },
          { id: 'freeze-row', title: 'Freeze First Row', command: 'sheet.command.set-first-row-frozen' },
          { id: 'freeze-column', title: 'Freeze First Column', command: 'sheet.command.set-first-column-frozen' },
        ] },
      ],
    },
    {
      id: 'review', title: 'Review', order: 8, groups: [
        { id: 'comments', title: 'Comments & Notes', items: [
          { id: 'new-comment', title: 'New Comment', icon: 'review', command: 'sheet.operation.show-comment-modal' },
          { id: 'new-note', title: 'New Note', command: 'sheet.operation.add-note-popup' },
          { id: 'comment-panel', title: 'Comment Management', command: 'sheet.operation.toggle-comment-panel' },
        ] },
        { id: 'protection', title: 'Protection', items: [
          { id: 'protect-range', title: 'Protect Selected Range', command: 'sheet.command.add-range-protection-from-toolbar' },
        ] },
        { id: 'history', title: 'History', items: [
          { id: 'edit-history', title: 'Edit History', command: 'univer.operation.toggle-edit-history' },
        ] },
      ],
    },
    {
      id: 'help', title: 'Help', order: 10, groups: [
        { id: 'support', title: 'Help & Support', items: [
          { id: 'feature-search', title: 'Search Tools', icon: 'help', command: 'ui.operation.open-feature-search' },
          { id: 'shortcuts', title: 'Keyboard Shortcuts', command: 'base-ui.operation.toggle-shortcut-panel' },
          { id: 'docs', title: 'GridNova Guide', icon: 'guide', action: 'open-help' },
        ] },
      ],
    },
    {
      id: 'draw', title: 'Draw', order: 11, groups: [
        { id: 'drawing', title: 'Drawing', items: [
          { id: 'insert-shape', title: 'Insert Shape', icon: 'draw', command: 'sheet.command.menu-insert-shape' },
          { id: 'layers', title: 'Selection Pane', command: 'sheet.operation.open-shape-object-list-panel' },
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
        { id: 'insert-table', title: 'Table', icon: 'table', command: 'sheet.operation.open-table-selector' },
        { id: 'insert-pivot-table', title: 'PivotTable', command: 'sheet.operation.open-pivot-table-range-selector-panel' },
      ],
    },
    {
      path: 'ribbon.insert', id: 'illustrations', title: 'Illustrations', order: 1, items: [
        { id: 'insert-picture', title: 'Pictures', command: 'sheet.command.insert-float-image' },
        { id: 'insert-shape', title: 'Shapes', icon: 'draw', command: 'sheet.command.menu-insert-shape' },
      ],
    },
    {
      path: 'ribbon.insert', id: 'charts', title: 'Charts', order: 2, items: [
        { id: 'insert-chart', title: 'Recommended Charts', icon: 'table', command: 'sheet.command.menu-insert-chart' },
        { id: 'insert-sparkline', title: 'Sparkline', command: 'sheet.operation.open-sparkline-selector' },
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
        { id: 'insert-function', title: 'Insert Function', command: 'formula-ui.operation.insert-function.common' },
        { id: 'more-functions', title: 'More Functions', command: 'formula-ui.operation.more-functions' },
        { id: 'financial-functions', title: 'Financial', command: 'formula-ui.operation.insert-function.financial' },
        { id: 'logical-functions', title: 'Logical', command: 'formula-ui.operation.insert-function.logical' },
        { id: 'text-functions', title: 'Text', command: 'formula-ui.operation.insert-function.text' },
        { id: 'date-functions', title: 'Date & Time', command: 'formula-ui.operation.insert-function.date' },
        { id: 'lookup-functions', title: 'Lookup & Reference', command: 'formula-ui.operation.insert-function.lookup' },
        { id: 'math-functions', title: 'Math & Trig', command: 'formula-ui.operation.insert-function.math' },
        { id: 'statistical-functions', title: 'Statistical', command: 'formula-ui.operation.insert-function.statistical' },
      ],
    },
    {
      path: 'ribbon.formulas', id: 'calculation', title: 'Calculation', order: 1, items: [
        { id: 'calculation-options', title: 'Calculation Options', command: 'sheets-formula-ui.operation.calculation-options' },
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
      ],
    },
    {
      path: 'ribbon.data', id: 'outline', title: 'Outline', order: 2, items: [
        { id: 'group-rows-columns', title: 'Group', command: 'sheet.toolbar.dimension-outline.group' },
      ],
    },
    {
      path: 'ribbon.view', id: 'show', title: 'Show', order: 0, items: [
        { id: 'toggle-gridlines', title: 'Gridlines', command: 'sheet.command.toggle-gridlines' },
      ],
    },
    {
      path: 'ribbon.view', id: 'window', title: 'Window', order: 1, items: [
        { id: 'freeze-panes', title: 'Freeze Panes', command: 'sheet.toolbar.sheet-frozen' },
        { id: 'zoom', title: 'Zoom', command: 'sheet.menu.zoom-ratio' },
        { id: 'fullscreen', title: 'Full Screen', command: 'base-ui.operation.toggle-fullscreen' },
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
    'univer.command.undo': { icon: 'gridnova-icon-home' },
    'univer.command.redo': { icon: 'gridnova-icon-home' },
    'ui.operation.activate-format-painter': { icon: 'gridnova-icon-layout' },
    'sheet.command.menu-insert-chart': { icon: 'gridnova-icon-table' },
    'sheet.operation.open.conditional.formatting.panel': { icon: 'gridnova-icon-layout' },
    'sheet.command.smart-toggle-filter': { icon: 'gridnova-icon-table' },
    'sheet.operation.open.numfmt.panel': { icon: 'gridnova-icon-table' },
    'sheet.operation.open-pivot-table-range-selector-panel': { icon: 'gridnova-icon-table' },
    'sheet.operation.open-sparkline-selector': { icon: 'gridnova-icon-table' },
    'sheet.operation.insert-hyper-link-toolbar': { icon: 'gridnova-icon-share' },
    'ui.operation.open-find-dialog': { icon: 'gridnova-icon-help' },
    'sheet.toolbar.sheet-frozen': { icon: 'gridnova-icon-layout' },
    'sheet.command.toggle-gridlines': { icon: 'gridnova-icon-layout' },
    'sheet.menu.image': { icon: 'gridnova-icon-table' },
    'sheet.operation.print-open': { icon: 'gridnova-icon-print' },
    'sheet.operation.open-print-grid-sidebar': { icon: 'gridnova-icon-print' },
    'sheet.operation.toggle-comment-panel': { icon: 'gridnova-icon-review' },
    'sheet.operation.show-comment-modal': { icon: 'gridnova-icon-review' },
    'sheet.command.add-range-protection-from-toolbar': { icon: 'gridnova-icon-review' },
    'sheet.command.menu-insert-shape': { icon: 'gridnova-icon-draw' },
  })

  return [...icons, ...menus, clipboardGroup, ...homeGroupMenus, ...nativeTabGroupMenus]
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
    let disposed = false

    try {
      const config = initialConfigRef.current
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
        ],
      })

      univerInstance = univer
      const workbookData = config.initialSnapshot
        ? ensureMinimumColumnCount(config.initialSnapshot)
        : createBlankWorkbookSnapshot(config.workbookId, config.title)

      univerAPI.createWorkbook(workbookData)
      addColumnHeaderDeleteAction(univerAPI)
      const ribbonDisposables = addMicrosoftInspiredRibbonTabs(univerAPI, (action) => onRibbonActionRef.current?.(action))

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
