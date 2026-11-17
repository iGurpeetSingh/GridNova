import { useEffect, useRef, useState } from 'react'
import { ArrowDownloadRegular, ArrowUploadRegular, BookOpenRegular, CommentRegular, DrawShapeRegular, GridRegular, LayoutRegular, PrintRegular, QuestionCircleRegular, SaveRegular, ShareRegular, TableRegular } from '@fluentui/react-icons'
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
  layout: LayoutRegular,
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

  const homeGroups = [
    {
      id: 'number', title: 'Number', items: [
        { id: 'number-format', title: 'Number Format', icon: 'table', command: 'sheet.operation.open.numfmt.panel' },
        { id: 'currency', title: 'Currency', command: 'sheet.command.numfmt.set.currency' },
        { id: 'percent', title: 'Percentage', command: 'sheet.command.numfmt.set.percent' },
        { id: 'increase-decimal', title: 'Increase Decimal', command: 'sheet.command.numfmt.add.decimal.command' },
        { id: 'decrease-decimal', title: 'Decrease Decimal', command: 'sheet.command.numfmt.subtract.decimal.command' },
      ],
    },
    {
      id: 'styles', title: 'Styles', items: [
        { id: 'conditional-formatting', title: 'Conditional Formatting', icon: 'layout', command: 'sheet.operation.open.conditional.formatting.panel' },
        { id: 'data-bars', title: 'Data Bars', command: 'sheet.command.add-data-bar-conditional-rule' },
        { id: 'icon-sets', title: 'Icon Sets', command: 'sheet.command.add-icon-set-conditional-rule' },
      ],
    },
    {
      id: 'cells', title: 'Cells', items: [
        { id: 'insert-row', title: 'Insert Row', command: 'sheet.command.insert-row-before' },
        { id: 'insert-column', title: 'Insert Column', command: 'sheet.command.insert-col-before' },
        { id: 'delete-row', title: 'Delete Row', command: 'sheet.command.remove-row-confirm' },
        { id: 'delete-column', title: 'Delete Column', command: 'sheet.command.remove-col-confirm' },
      ],
    },
    {
      id: 'editing', title: 'Editing', items: [
        { id: 'sort-ascending', title: 'Sort A to Z', command: 'sheet.command.sort-range-asc' },
        { id: 'sort-descending', title: 'Sort Z to A', command: 'sheet.command.sort-range-desc' },
        { id: 'filter', title: 'Filter', icon: 'table', command: 'sheet.command.smart-toggle-filter' },
        { id: 'clear-contents', title: 'Clear Contents', command: 'sheet.command.clear-selection-content' },
        { id: 'find', title: 'Find & Replace', icon: 'help', command: 'ui.operation.open-find-dialog' },
      ],
    },
  ]

  const homeGroupMenus = homeGroups.map((group) => {
    const groupMenu = univerAPI.createSubmenu({
      id: 'gridnova.ribbon.home.' + group.id,
      title: group.title,
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
    'ribbon.start.history': { title: 'Clipboard', order: 0 },
    'ribbon.start.format': { title: 'Font', order: 1 },
    'ribbon.start.layout': { title: 'Alignment', order: 2 },
    'ribbon.start.others': { title: 'Editing', order: 6 },
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

  return [...icons, ...menus, ...homeGroupMenus]
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
