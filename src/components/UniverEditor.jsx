import { useEffect, useRef, useState } from 'react'
import { ArrowDownloadRegular, ArrowRedoRegular, ArrowUndoRegular, ArrowUploadRegular, BookOpenRegular, BorderAllRegular, CalculatorRegular, ClipboardPasteRegular, CurrencyDollarRupeeRegular, CommentRegular, CopyRegular, CutRegular, DrawShapeRegular, FilterRegular, GridRegular, ImageRegular, LinkRegular, PaintBucketRegular, PaintBrushRegular, PrintRegular, QuestionCircleRegular, SaveRegular, SearchRegular, ShareRegular, TableRegular, TextAlignCenterRegular, TextAlignLeftRegular, TextAlignRightRegular, TextBoldRegular, TextColorRegular, TextItalicRegular, TextStrikethroughRegular, TextUnderlineRegular } from '@fluentui/react-icons'
import { createUniver, LocaleType } from '@univerjs/presets'
import { UniverSheetsCorePreset } from '@univerjs/preset-sheets-core'
import '@univerjs/ui/facade'
import '@univerjs/sheets/facade'
import { withRibbonGridLayouts } from '../utils/ribbonLayout.js'
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

async function executeRibbonItem(univerAPI, item, onRibbonAction) {
  try {
    let result
    if (typeof item.run === 'function') {
      result = await item.run(univerAPI)
    } else if (item.action) {
      result = await onRibbonAction?.(item.action)
    } else if (item.command) {
      result = await univerAPI.executeCommand(item.command, item.params)
    } else {
      return false
    }

    if (result === false) {
      univerAPI.showMessage({
        content: item.title + ' could not run. Check the active selection or required workbook context.',
        type: 'warning',
      })
    }
    return result
  } catch (error) {
    univerAPI.showMessage({
      content: item.title + ' failed: ' + (error instanceof Error ? error.message : 'Unknown command error.'),
      type: 'error',
    })
    return false
  }
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

    withRibbonGridLayouts(group.items).forEach((item) => {
      const action = () => executeRibbonItem(univerAPI, item, onRibbonAction)
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



function GridNovaDataChart({ data }) {
  const labels = Array.isArray(data?.labels) ? data.labels : []
  const values = Array.isArray(data?.values) ? data.values.map((value) => Number(value) || 0) : []
  const type = ['bar', 'column', 'line', 'pie', 'area'].includes(data?.type) ? data.type : 'bar'
  const colors = ['#217346', '#4472c4', '#ed7d31', '#70ad47', '#a5a5a5', '#ffc000', '#5b9bd5', '#8064a2', '#c0504d', '#4bacc6']
  const positiveValues = values.map((value) => Math.max(0, value))
  const maxValue = Math.max(1, ...positiveValues)
  const total = positiveValues.reduce((sum, value) => sum + value, 0)
  const chartTitle = String(data?.title || (type[0].toUpperCase() + type.slice(1) + ' Chart'))
  const plotLeft = 34
  const plotTop = 24
  const plotWidth = 306
  const plotHeight = 132
  const chartBottom = plotTop + plotHeight
  const points = positiveValues.map((value, index) => ({
    x: positiveValues.length <= 1 ? plotLeft + plotWidth / 2 : plotLeft + index * plotWidth / (positiveValues.length - 1),
    y: chartBottom - value / maxValue * (plotHeight - 8),
    value,
    label: String(labels[index] || index + 1),
  }))
  const pieRadius = 58
  const pieCircumference = 2 * Math.PI * pieRadius
  let pieOffset = 0

  return (
    <div style={{ width: '100%', height: '100%', padding: 8, boxSizing: 'border-box', background: '#fff', border: '1px solid #cbd5e1', borderRadius: 8, fontFamily: 'Arial,sans-serif' }}>
      <div style={{ height: 18, overflow: 'hidden', whiteSpace: 'nowrap', textOverflow: 'ellipsis', fontSize: 12, fontWeight: 700, color: '#243342' }}>{chartTitle}</div>
      <svg viewBox="0 0 360 205" width="100%" height="calc(100% - 18px)" role="img" aria-label={'GridNova ' + type + ' chart'}>
        {(type === 'bar') && values.map((value, index) => {
          const numeric = positiveValues[index]
          const y = 26 + index * Math.max(15, 145 / Math.max(1, values.length))
          const height = Math.max(8, Math.min(22, 140 / Math.max(1, values.length) - 3))
          return (
            <g key={index}>
              <text x="4" y={y + height * .78} fontSize="9" fill="#475569">{String(labels[index] || index + 1).slice(0, 7)}</text>
              <rect x="72" y={y} width={numeric / maxValue * 250} height={height} rx="3" fill={colors[index % colors.length]} />
              <text x={Math.min(336, 76 + numeric / maxValue * 250)} y={y + height * .78} fontSize="9" fill="#334155">{numeric}</text>
            </g>
          )
        })}
        {type === 'column' && <>
          <line x1={plotLeft} y1={chartBottom} x2="348" y2={chartBottom} stroke="#94a3b8" strokeWidth="1" />
          {values.map((value, index) => {
            const numeric = positiveValues[index]
            const width = Math.max(7, Math.min(32, plotWidth / Math.max(1, values.length) - 5))
            const x = plotLeft + index * (plotWidth / Math.max(1, values.length)) + 2
            const height = numeric / maxValue * (plotHeight - 8)
            const y = chartBottom - height
            return <g key={index}>
              <rect x={x} y={y} width={width} height={height} rx="2" fill={colors[index % colors.length]} />
              <text x={x + width / 2} y="177" textAnchor="middle" fontSize="8" fill="#475569">{String(labels[index] || '').slice(0, 6)}</text>
              <text x={x + width / 2} y={Math.max(18, y - 4)} textAnchor="middle" fontSize="8" fill="#334155">{numeric}</text>
            </g>
          })}
        </>}
        {type === 'line' && <>
          <line x1={plotLeft} y1={chartBottom} x2="348" y2={chartBottom} stroke="#94a3b8" strokeWidth="1" />
          <polyline points={points.map((point) => point.x + ',' + point.y).join(' ')} fill="none" stroke="#217346" strokeWidth="3" strokeLinejoin="round" strokeLinecap="round" />
          {points.map((point, index) => <g key={index}>
            <circle cx={point.x} cy={point.y} r="4" fill="#217346" />
            <text x={point.x} y="177" textAnchor="middle" fontSize="8" fill="#475569">{point.label.slice(0, 6)}</text>
          </g>)}
        </>}
        {type === 'area' && <>
          <line x1={plotLeft} y1={chartBottom} x2="348" y2={chartBottom} stroke="#94a3b8" strokeWidth="1" />
          <polygon points={[plotLeft + ',' + chartBottom, ...points.map((point) => point.x + ',' + point.y), points.at(-1)?.x + ',' + chartBottom].join(' ')} fill="#70ad47" fillOpacity=".28" stroke="none" />
          <polyline points={points.map((point) => point.x + ',' + point.y).join(' ')} fill="none" stroke="#217346" strokeWidth="2.5" strokeLinejoin="round" strokeLinecap="round" />
          {points.map((point, index) => <text key={index} x={point.x} y="177" textAnchor="middle" fontSize="8" fill="#475569">{point.label.slice(0, 6)}</text>)}
        </>}
        {type === 'pie' && <>
          {total > 0 && values.map((value, index) => {
            const fraction = positiveValues[index] / total
            const length = fraction * pieCircumference
            const offset = pieOffset
            pieOffset += length
            return <circle
              key={index}
              cx="148"
              cy="96"
              r={pieRadius}
              fill="none"
              stroke={colors[index % colors.length]}
              strokeWidth="42"
              strokeDasharray={length + ' ' + (pieCircumference - length)}
              strokeDashoffset={-offset}
              transform="rotate(-90 148 96)"
            />
          })}
          {values.slice(0, 8).map((value, index) => <g key={index}>
            <rect x="236" y={28 + index * 19} width="9" height="9" fill={colors[index % colors.length]} />
            <text x="250" y={36 + index * 19} fontSize="9" fill="#334155">{String(labels[index] || index + 1).slice(0, 10)}: {positiveValues[index]}</text>
          </g>)}
        </>}
      </svg>
    </div>
  )
}

function GridNovaTextBox({ data }) {
  const text = String(data?.text ?? '')
  const fontSize = Number.isFinite(Number(data?.fontSize)) ? Math.min(72, Math.max(10, Number(data.fontSize))) : 18
  const color = typeof data?.color === 'string' ? data.color : '#243342'
  const align = ['left', 'center', 'right'].includes(data?.align) ? data.align : 'left'
  return (
    <div
      role="note"
      aria-label="GridNova text box"
      style={{
        width: '100%',
        height: '100%',
        boxSizing: 'border-box',
        display: 'flex',
        alignItems: 'center',
        justifyContent: align === 'left' ? 'flex-start' : align === 'right' ? 'flex-end' : 'center',
        padding: 10,
        overflow: 'hidden',
        whiteSpace: 'pre-wrap',
        overflowWrap: 'anywhere',
        background: data?.background || 'rgba(255,255,255,.96)',
        border: data?.border || '1px solid #93a4b5',
        borderRadius: 6,
        color,
        fontSize,
        fontWeight: data?.bold ? 700 : 400,
        textAlign: align,
        fontFamily: 'Arial, sans-serif',
      }}
    >
      {text}
    </div>
  )
}

function GridNovaSymbol({ data }) {
  const symbol = String(data?.text ?? '★').slice(0, 8)
  const color = typeof data?.color === 'string' ? data.color : '#217346'
  return (
    <div
      role="img"
      aria-label={'GridNova symbol ' + symbol}
      style={{
        width: '100%',
        height: '100%',
        display: 'grid',
        placeItems: 'center',
        boxSizing: 'border-box',
        color,
        fontSize: Math.min(96, Math.max(18, Number(data?.fontSize) || 56)),
        lineHeight: 1,
        background: 'transparent',
        userSelect: 'none',
      }}
    >
      {symbol}
    </div>
  )
}

function GridNovaInkCanvas({ data, univerAPI }) {
  const [strokes, setStrokes] = useState(() => Array.isArray(data?.strokes) ? data.strokes : [])
  const [draft, setDraft] = useState(null)
  const draftRef = useRef(null)
  const pointerIdRef = useRef(null)
  const svgRef = useRef(null)

  useEffect(() => {
    setStrokes(Array.isArray(data?.strokes) ? data.strokes : [])
  }, [data?.strokes])

  const mode = data?.mode === 'highlighter' ? 'highlighter' : 'pen'
  const defaultColor = mode === 'highlighter' ? '#ffe66d' : '#1d4ed8'
  const defaultWidth = mode === 'highlighter' ? 14 : 3

  function getPoint(event) {
    const bounds = svgRef.current?.getBoundingClientRect()
    if (!bounds || !bounds.width || !bounds.height) return null
    return {
      x: Math.max(0, Math.min(360, (event.clientX - bounds.left) / bounds.width * 360)),
      y: Math.max(0, Math.min(200, (event.clientY - bounds.top) / bounds.height * 200)),
    }
  }

  function persistStrokes(nextStrokes) {
    setStrokes(nextStrokes)
    const sheet = univerAPI?.getActiveWorkbook?.()?.getActiveSheet?.()
    if (!sheet || !data?.id || typeof sheet.updateFloatDom !== 'function') return
    try {
      sheet.updateFloatDom(data.id, {
        data: { ...data, strokes: nextStrokes },
      })
    } catch (error) {
      univerAPI?.showMessage?.({
        content: error instanceof Error ? error.message : 'Could not save the drawing stroke.',
        type: 'error',
      })
    }
  }

  function handlePointerDown(event) {
    if (event.button !== undefined && event.button !== 0) return
    const point = getPoint(event)
    if (!point) return
    event.preventDefault()
    pointerIdRef.current = event.pointerId
    event.currentTarget.setPointerCapture?.(event.pointerId)
    const nextDraft = {
      points: [point],
      color: mode === 'highlighter' ? '#f2c94c' : defaultColor,
      width: defaultWidth,
      opacity: mode === 'highlighter' ? 0.48 : 1,
    }
    draftRef.current = nextDraft
    setDraft(nextDraft)
  }

  function handlePointerMove(event) {
    if (pointerIdRef.current !== event.pointerId) return
    const point = getPoint(event)
    if (!point) return
    const previous = draftRef.current
    if (!previous) return
    const nextDraft = { ...previous, points: [...previous.points, point] }
    draftRef.current = nextDraft
    setDraft(nextDraft)
  }

  function finishStroke(event) {
    if (pointerIdRef.current !== event.pointerId) return
    pointerIdRef.current = null
    try { event.currentTarget.releasePointerCapture?.(event.pointerId) } catch { /* pointer already released */ }
    const finished = draftRef.current
    draftRef.current = null
    setDraft(null)
    if (finished?.points?.length) persistStrokes([...strokes, finished])
  }

  const visibleStrokes = draft ? [...strokes, draft] : strokes
  return (
    <div style={{ position: 'relative', width: '100%', height: '100%', boxSizing: 'border-box', background: 'rgba(255,255,255,.88)', border: '1px solid #94a3b8', borderRadius: 8, overflow: 'hidden' }}>
      <div style={{ position: 'absolute', zIndex: 1, top: 5, left: 8, pointerEvents: 'none', padding: '2px 6px', borderRadius: 4, background: 'rgba(255,255,255,.82)', color: '#475569', fontSize: 10, fontFamily: 'Arial,sans-serif' }}>
        {mode === 'highlighter' ? 'Highlighter · drag to mark' : 'Pen · drag to draw'}
      </div>
      <svg
        ref={svgRef}
        viewBox="0 0 360 200"
        preserveAspectRatio="none"
        width="100%"
        height="100%"
        style={{ display: 'block', touchAction: 'none', cursor: 'crosshair' }}
        onPointerDown={handlePointerDown}
        onPointerMove={handlePointerMove}
        onPointerUp={finishStroke}
        onPointerCancel={finishStroke}
      >
        {visibleStrokes.map((stroke, index) => (
          <polyline
            key={index}
            points={(stroke.points || []).map((point) => point.x + ',' + point.y).join(' ')}
            fill="none"
            stroke={stroke.color || defaultColor}
            strokeWidth={stroke.width || defaultWidth}
            strokeOpacity={stroke.opacity ?? 1}
            strokeLinecap="round"
            strokeLinejoin="round"
          />
        ))}
      </svg>
    </div>
  )
}

function applyCellStylePreset(univerAPI, presetName) {
  const sheet = univerAPI.getActiveWorkbook()?.getActiveSheet()
  const range = sheet?.getActiveRange()
  if (!range) {
    univerAPI.showMessage({ content: 'Select a cell or range before applying a Cell Style.', type: 'warning' })
    return false
  }

  const presets = {
    title: { background: '#217346', color: '#ffffff', weight: 'bold', size: 14, horizontal: 'center', vertical: 'middle' },
    heading: { background: '#d9ead3', color: '#274e13', weight: 'bold', size: 11 },
    input: { background: '#fff2cc', color: '#7f6000', weight: 'normal', size: 11 },
    good: { background: '#d9ead3', color: '#38761d', weight: 'bold', size: 11 },
    bad: { background: '#f4cccc', color: '#990000', weight: 'bold', size: 11 },
    neutral: { background: '#eeeeee', color: '#444444', weight: 'normal', size: 11 },
  }
  const preset = presets[presetName]
  if (!preset) return false

  range.setBackgroundColor(preset.background)
  range.setFontColor(preset.color)
  range.setFontWeight(preset.weight)
  range.setFontSize(preset.size)
  if (preset.horizontal) range.setHorizontalAlignment(preset.horizontal)
  if (preset.vertical) range.setVerticalAlignment(preset.vertical)

  univerAPI.showMessage({ content: presetName[0].toUpperCase() + presetName.slice(1) + ' style applied to the selected range.', type: 'success' })
  return true
}

function openFormatCellsDialog(univerAPI) {
  const sheet = univerAPI.getActiveWorkbook()?.getActiveSheet()
  const range = sheet?.getActiveRange()
  if (!sheet || !range) {
    univerAPI.showMessage({ content: 'Select a cell or range before opening Format Cells.', type: 'warning' })
    return false
  }

  const currentFormat = range.getNumberFormat?.() || 'General'
  const choice = window.prompt(
    'FORMAT CELLS — choose a format:\n1 General\n2 Number (2 decimals)\n3 Integer\n4 Percentage\n5 Currency (₹)\n6 Date\n7 Text\n8 Custom format pattern',
    '1',
  )
  if (choice == null) return false

  const presets = {
    '1': 'General',
    '2': '#,##0.00',
    '3': '#,##0',
    '4': '0.00%',
    '5': '₹#,##0.00',
    '6': 'dd-mmm-yyyy',
    '7': '@',
  }
  let pattern = presets[choice.trim()]
  if (choice.trim() === '8') {
    pattern = window.prompt('Enter a Univer/Excel number format pattern.', currentFormat)
    if (pattern == null) return false
  }
  if (!pattern) {
    univerAPI.showMessage({ content: 'Choose a listed format or enter option 8 for a custom pattern.', type: 'warning' })
    return false
  }

  try {
    if (typeof range.setNumberFormat !== 'function') {
      showFeatureUnavailable(univerAPI, 'Format Cells number format', 'The installed Number Format Facade is not available in this preset.')
      return false
    }
    range.setNumberFormat(pattern)
    univerAPI.showMessage({ content: 'Number format applied to the selected range: ' + pattern, type: 'success' })
    return true
  } catch (error) {
    univerAPI.showMessage({ content: error instanceof Error ? error.message : 'Could not apply the number format.', type: 'error' })
    return false
  }
}

function insertChartOverlay(univerAPI, chartType = 'column') {
  const sheet = univerAPI.getActiveWorkbook()?.getActiveSheet()
  const range = sheet?.getActiveRange()
  if (!sheet || !range || typeof sheet.addFloatDomToRange !== 'function') {
    univerAPI.showMessage({ content: 'Select a two-column data range to insert a chart overlay.', type: 'warning' })
    return false
  }
  const rows = range.getValues()
  const dataRows = rows.map((row, index) => ({
    label: String(row?.[0] ?? index + 1),
    value: Number(row?.[1]),
  })).filter((row) => Number.isFinite(row.value))
  if (!dataRows.length) {
    univerAPI.showMessage({ content: 'The selected range needs labels in the first column and numbers in the second column.', type: 'warning' })
    return false
  }
  if (chartType === 'pie' && dataRows.some((row) => row.value < 0)) {
    univerAPI.showMessage({ content: 'Pie chart overlays require non-negative values.', type: 'warning' })
    return false
  }
  if (chartType === 'pie' && dataRows.every((row) => row.value === 0)) {
    univerAPI.showMessage({ content: 'Pie charts require at least one value greater than zero.', type: 'warning' })
    return false
  }
  const id = 'gridnova-chart-' + (globalThis.crypto?.randomUUID?.() || Date.now())
  const chart = sheet.addFloatDomToRange(
    range,
    {
      componentKey: 'gridnova-data-chart',
      data: {
        type: chartType,
        title: chartType[0].toUpperCase() + chartType.slice(1) + ' Chart',
        labels: dataRows.map((row) => row.label),
        values: dataRows.map((row) => row.value),
      },
      allowTransform: true,
    },
    { width: 420, height: 260 },
    id,
  )
  if (!chart) {
    univerAPI.showMessage({ content: 'The chart overlay could not be inserted.', type: 'error' })
    return false
  }
  univerAPI.showMessage({
    content: chartType[0].toUpperCase() + chartType.slice(1) + ' chart overlay inserted. It is a GridNova drawing object, not a native editable Excel chart.',
    type: 'success',
  })
  return true
}

function insertBasicChart(univerAPI) {
  return insertChartOverlay(univerAPI, 'bar')
}

function GridNovaRectangleShape({ data }) {
  return <svg viewBox="0 0 180 100" width="100%" height="100%" aria-label="Rectangle shape">
    <rect x="8" y="8" width="164" height="84" rx="5" fill={data?.fill || '#dbeafe'} fillOpacity={data?.fillOpacity ?? '.55'} stroke={data?.stroke || '#2563eb'} strokeWidth={Number(data?.strokeWidth) || 3} />
  </svg>
}

function GridNovaEllipseShape({ data }) {
  return <svg viewBox="0 0 180 100" width="100%" height="100%" aria-label="Ellipse shape">
    <ellipse cx="90" cy="50" rx="78" ry="38" fill={data?.fill || '#dcfce7'} fillOpacity={data?.fillOpacity ?? '.55'} stroke={data?.stroke || '#16834b'} strokeWidth={Number(data?.strokeWidth) || 3} />
  </svg>
}

function GridNovaLineShape({ data }) {
  return <svg viewBox="0 0 180 100" width="100%" height="100%" aria-label="Line shape">
    <line x1="8" y1="88" x2="172" y2="12" stroke={data?.stroke || '#d97706'} strokeWidth={Number(data?.strokeWidth) || 4} strokeLinecap="round" />
  </svg>
}

function registerDrawingComponents(univerAPI) {
  return [
    univerAPI.registerComponent('gridnova-draw-rectangle', GridNovaRectangleShape),
    univerAPI.registerComponent('gridnova-draw-ellipse', GridNovaEllipseShape),
    univerAPI.registerComponent('gridnova-draw-line', GridNovaLineShape),
    univerAPI.registerComponent('gridnova-bar-chart', GridNovaDataChart),
    univerAPI.registerComponent('gridnova-data-chart', GridNovaDataChart),
    univerAPI.registerComponent('gridnova-text-box', GridNovaTextBox),
    univerAPI.registerComponent('gridnova-symbol', GridNovaSymbol),
    univerAPI.registerComponent('gridnova-ink-canvas', (props) => <GridNovaInkCanvas {...props} univerAPI={univerAPI} />),
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
  const shapeDefaults = {
    rectangle: { fill: '#dbeafe', stroke: '#2563eb', fillOpacity: '.55', strokeWidth: 3 },
    ellipse: { fill: '#dcfce7', stroke: '#16834b', fillOpacity: '.55', strokeWidth: 3 },
    line: { stroke: '#d97706', strokeWidth: 4 },
  }
  const shape = sheet.addFloatDomToRange(
    range,
    { componentKey, data: shapeDefaults[shapeType] || {}, allowTransform: true },
    { width: 180, height: 100 },
    id,
  )
  if (!shape) {
    univerAPI.showMessage({ content: 'The shape could not be inserted.', type: 'error' })
    return
  }
  univerAPI.showMessage({ content: shapeType[0].toUpperCase() + shapeType.slice(1) + ' added over the selected range.', type: 'success' })
}

function getSelectedGridNovaFloatDom(univerAPI, allowedPrefixes = ['gridnova-draw-', 'gridnova-text-box', 'gridnova-symbol']) {
  const sheet = univerAPI.getActiveWorkbook()?.getActiveSheet()
  if (!sheet || typeof sheet.getAllFloatDoms !== 'function') {
    univerAPI.showMessage({ content: 'Shape formatting requires the installed Drawing Facade.', type: 'warning' })
    return null
  }

  const items = sheet.getAllFloatDoms().filter((item) =>
    allowedPrefixes.some((prefix) => item.componentKey?.startsWith(prefix))
  )
  if (!items.length) {
    univerAPI.showMessage({ content: 'Insert a GridNova shape, text box or symbol first.', type: 'warning' })
    return null
  }

  const options = items.map((item, index) => {
    const kind = item.componentKey?.replace('gridnova-draw-', '').replace('gridnova-', '') || 'shape'
    return (index + 1) + '. ' + kind + ' · ' + item.id.slice(-8)
  })
  const answer = window.prompt('Choose the GridNova object to format:\n' + options.join('\n'), String(items.length))
  if (answer == null) return null
  const index = Number(answer) - 1
  if (!Number.isInteger(index) || index < 0 || index >= items.length) {
    univerAPI.showMessage({ content: 'Enter one of the listed object numbers.', type: 'warning' })
    return null
  }
  return { sheet, item: items[index] }
}

function setGridNovaShapeFill(univerAPI) {
  const selected = getSelectedGridNovaFloatDom(univerAPI)
  if (!selected) return false
  const color = window.prompt('Enter a Fill Color (CSS color or hex).', '#dbeafe')
  if (color == null) return false
  const value = color.trim()
  if (!value || value.length > 40) {
    univerAPI.showMessage({ content: 'Enter a valid CSS color.', type: 'warning' })
    return false
  }
  const current = selected.item.data || {}
  const componentKey = selected.item.componentKey || ''
  const data = componentKey === 'gridnova-text-box'
    ? { ...current, background: value }
    : componentKey === 'gridnova-symbol'
      ? { ...current, color: value }
      : { ...current, fill: value }
  selected.sheet.updateFloatDom(selected.item.id, { data })
  univerAPI.showMessage({ content: 'Shape fill updated.', type: 'success' })
  return true
}

function setGridNovaShapeOutline(univerAPI) {
  const selected = getSelectedGridNovaFloatDom(univerAPI)
  if (!selected) return false
  const color = window.prompt('Enter an Outline Color (CSS color or hex).', '#2563eb')
  if (color == null) return false
  const value = color.trim()
  if (!value || value.length > 40) {
    univerAPI.showMessage({ content: 'Enter a valid CSS color.', type: 'warning' })
    return false
  }
  const current = selected.item.data || {}
  const componentKey = selected.item.componentKey || ''
  if (componentKey === 'gridnova-symbol') {
    univerAPI.showMessage({ content: 'Symbols support Color, but not a separate Outline.', type: 'info' })
    return false
  }
  const data = componentKey === 'gridnova-text-box'
    ? { ...current, border: '2px solid ' + value }
    : { ...current, stroke: value }
  selected.sheet.updateFloatDom(selected.item.id, { data })
  univerAPI.showMessage({ content: 'Shape outline updated.', type: 'success' })
  return true
}

function setGridNovaShapeProperties(univerAPI) {
  const selected = getSelectedGridNovaFloatDom(univerAPI)
  if (!selected) return false
  const current = selected.item.position || {}
  const width = window.prompt('Object width in pixels.', String(Math.round(current.width || 180)))
  if (width == null) return false
  const height = window.prompt('Object height in pixels.', String(Math.round(current.height || 100)))
  if (height == null) return false
  const angle = window.prompt('Object rotation in degrees.', String(Math.round(current.angle || 0)))
  if (angle == null) return false
  const nextWidth = Number(width)
  const nextHeight = Number(height)
  const nextAngle = Number(angle)
  if (!Number.isFinite(nextWidth) || nextWidth < 20 || nextWidth > 3000
    || !Number.isFinite(nextHeight) || nextHeight < 20 || nextHeight > 3000
    || !Number.isFinite(nextAngle) || Math.abs(nextAngle) > 360) {
    univerAPI.showMessage({ content: 'Use width/height 20–3000 px and rotation -360° to 360°.', type: 'warning' })
    return false
  }
  selected.sheet.updateFloatDom(selected.item.id, {
    position: { width: nextWidth, height: nextHeight, angle: nextAngle },
  })
  univerAPI.showMessage({ content: 'Shape properties updated.', type: 'success' })
  return true
}

function deleteGridNovaShape(univerAPI) {
  const selected = getSelectedGridNovaFloatDom(univerAPI)
  if (!selected) return false
  if (!window.confirm('Delete the selected GridNova shape or text overlay?')) return false
  selected.sheet.removeFloatDom(selected.item.id)
  univerAPI.showMessage({ content: 'GridNova object deleted.', type: 'success' })
  return true
}

function insertInkCanvas(univerAPI, mode = 'pen') {
  const sheet = univerAPI.getActiveWorkbook()?.getActiveSheet()
  if (!sheet || typeof sheet.addFloatDomToRange !== 'function') {
    univerAPI.showMessage({ content: 'The worksheet Drawing Facade is unavailable.', type: 'warning' })
    return false
  }

  const range = sheet.getActiveRange() || sheet.getRange('A1')
  const id = 'gridnova-ink-' + (globalThis.crypto?.randomUUID?.() || Date.now())
  const data = { id, mode, createdAt: Date.now(), strokes: [] }
  const drawing = sheet.addFloatDomToRange(
    range,
    { componentKey: 'gridnova-ink-canvas', data, allowTransform: true },
    { width: 420, height: 240 },
    id,
  )
  if (!drawing) {
    univerAPI.showMessage({ content: 'The drawing canvas could not be inserted.', type: 'error' })
    return false
  }
  univerAPI.showMessage({
    content: (mode === 'highlighter' ? 'Highlighter' : 'Pen') + ' canvas added. Drag on it to create saved strokes.',
    type: 'success',
  })
  return true
}

function getLatestInkCanvas(univerAPI) {
  const sheet = univerAPI.getActiveWorkbook()?.getActiveSheet()
  if (!sheet || typeof sheet.getAllFloatDoms !== 'function') return null
  const canvases = sheet.getAllFloatDoms()
    .filter((item) => item.componentKey === 'gridnova-ink-canvas')
    .sort((a, b) => Number(b.data?.createdAt || 0) - Number(a.data?.createdAt || 0))
  return canvases.length ? { sheet, item: canvases[0] } : null
}

function eraseLastInkStroke(univerAPI) {
  const canvas = getLatestInkCanvas(univerAPI)
  if (!canvas) {
    univerAPI.showMessage({ content: 'Add a Pen or Highlighter canvas before using Eraser.', type: 'warning' })
    return false
  }
  const strokes = Array.isArray(canvas.item.data?.strokes) ? canvas.item.data.strokes : []
  if (!strokes.length) {
    univerAPI.showMessage({ content: 'There are no strokes to erase on the latest drawing canvas.', type: 'info' })
    return false
  }
  canvas.sheet.updateFloatDom(canvas.item.id, {
    data: { ...canvas.item.data, strokes: strokes.slice(0, -1) },
  })
  univerAPI.showMessage({ content: 'Last stroke erased from the latest drawing canvas.', type: 'success' })
  return true
}

function clearLatestInkCanvas(univerAPI) {
  const canvas = getLatestInkCanvas(univerAPI)
  if (!canvas) {
    univerAPI.showMessage({ content: 'There is no GridNova drawing canvas on this worksheet.', type: 'warning' })
    return false
  }
  const strokes = Array.isArray(canvas.item.data?.strokes) ? canvas.item.data.strokes : []
  if (!strokes.length) return false
  if (!window.confirm('Clear all strokes from the latest drawing canvas?')) return false
  canvas.sheet.updateFloatDom(canvas.item.id, {
    data: { ...canvas.item.data, strokes: [] },
  })
  univerAPI.showMessage({ content: 'Drawing canvas cleared.', type: 'success' })
  return true
}

const PRINT_PAPER_SIZES = ['A3', 'A4', 'A5', 'Letter', 'Legal', 'Tabloid']

function getWorkbookPrintSettings(univerAPI) {
  const workbookId = univerAPI.getActiveWorkbook()?.getId?.()
  if (!workbookId) return null
  const defaults = { orientation: 'portrait', paperSize: 'A4', margin: 12, scale: 100, printBackground: true }
  try {
    const saved = JSON.parse(window.localStorage.getItem('gridnova:print-settings:' + workbookId) || '{}')
    return {
      ...defaults,
      orientation: saved.orientation === 'landscape' ? 'landscape' : 'portrait',
      paperSize: PRINT_PAPER_SIZES.includes(saved.paperSize) ? saved.paperSize : 'A4',
      margin: Number.isFinite(Number(saved.margin)) ? Math.min(40, Math.max(5, Number(saved.margin))) : 12,
      scale: Number.isFinite(Number(saved.scale)) ? Math.min(150, Math.max(40, Number(saved.scale))) : 100,
      printBackground: saved.printBackground !== false,
    }
  } catch {
    return defaults
  }
}

function applyWorkbookPrintSettings(univerAPI, patch = {}, showNotice = true) {
  const workbookId = univerAPI.getActiveWorkbook()?.getId?.()
  if (!workbookId) return false
  const current = getWorkbookPrintSettings(univerAPI)
  if (!current) return false
  const next = { ...current, ...patch }
  next.orientation = next.orientation === 'landscape' ? 'landscape' : 'portrait'
  next.paperSize = PRINT_PAPER_SIZES.includes(next.paperSize) ? next.paperSize : 'A4'
  next.margin = Math.min(40, Math.max(5, Number(next.margin) || 12))
  next.scale = Math.min(150, Math.max(40, Number(next.scale) || 100))
  next.printBackground = next.printBackground !== false

  try {
    window.localStorage.setItem('gridnova:print-settings:' + workbookId, JSON.stringify(next))
  } catch {
    univerAPI.showMessage({ content: 'Print preferences could not be saved in this browser.', type: 'warning' })
  }

  const styleId = 'gridnova-print-settings-' + workbookId.replace(/[^a-zA-Z0-9_-]/g, '')
  let style = document.getElementById(styleId)
  if (!style) {
    style = document.createElement('style')
    style.id = styleId
    document.head.appendChild(style)
  }
  style.textContent = `
    @media print {
      @page { size: ${next.paperSize} ${next.orientation}; margin: ${next.margin}mm; }
      .univer-mount {
        zoom: ${next.scale / 100};
        -webkit-print-color-adjust: ${next.printBackground ? 'exact' : 'economy'} !important;
        print-color-adjust: ${next.printBackground ? 'exact' : 'economy'} !important;
      }
    }
  `
  if (showNotice) {
    univerAPI.showMessage({
      content: 'Print layout saved: ' + next.paperSize + ' ' + next.orientation + ', ' + next.margin + ' mm margins, ' + next.scale + '% scale.',
      type: 'success',
    })
  }
  return true
}

function setPrintOrientation(univerAPI) {
  const current = getWorkbookPrintSettings(univerAPI)
  if (!current) return false
  const value = window.prompt('Page orientation: portrait or landscape.', current.orientation)
  if (value == null) return false
  const orientation = value.trim().toLowerCase()
  if (!['portrait', 'landscape'].includes(orientation)) {
    univerAPI.showMessage({ content: 'Enter portrait or landscape.', type: 'warning' })
    return false
  }
  return applyWorkbookPrintSettings(univerAPI, { orientation })
}

function setPrintPaperSize(univerAPI) {
  const current = getWorkbookPrintSettings(univerAPI)
  if (!current) return false
  const value = window.prompt('Paper size: ' + PRINT_PAPER_SIZES.join(', '), current.paperSize)
  if (value == null) return false
  const paperSize = PRINT_PAPER_SIZES.find((item) => item.toLowerCase() === value.trim().toLowerCase())
  if (!paperSize) {
    univerAPI.showMessage({ content: 'Choose A3, A4, A5, Letter, Legal or Tabloid.', type: 'warning' })
    return false
  }
  return applyWorkbookPrintSettings(univerAPI, { paperSize })
}

function setPrintMargins(univerAPI) {
  const current = getWorkbookPrintSettings(univerAPI)
  if (!current) return false
  const value = window.prompt('Print margins in millimeters (5–40).', String(current.margin))
  if (value == null) return false
  const margin = Number(value)
  if (!Number.isFinite(margin) || margin < 5 || margin > 40) {
    univerAPI.showMessage({ content: 'Margins must be from 5 to 40 millimeters.', type: 'warning' })
    return false
  }
  return applyWorkbookPrintSettings(univerAPI, { margin })
}

function setPrintScale(univerAPI) {
  const current = getWorkbookPrintSettings(univerAPI)
  if (!current) return false
  const value = window.prompt('Print scale percentage (40–150).', String(current.scale))
  if (value == null) return false
  const scale = Number(value)
  if (!Number.isFinite(scale) || scale < 40 || scale > 150) {
    univerAPI.showMessage({ content: 'Print scale must be from 40% to 150%.', type: 'warning' })
    return false
  }
  return applyWorkbookPrintSettings(univerAPI, { scale })
}

function togglePrintBackground(univerAPI) {
  const current = getWorkbookPrintSettings(univerAPI)
  if (!current) return false
  return applyWorkbookPrintSettings(univerAPI, { printBackground: !current.printBackground })
}

function applySavedWorkbookPrintSettings(univerAPI) {
  const current = getWorkbookPrintSettings(univerAPI)
  if (!current) return false
  return applyWorkbookPrintSettings(univerAPI, current, false)
}

function insertTextBox(univerAPI) {
  const sheet = univerAPI.getActiveWorkbook()?.getActiveSheet()
  if (!sheet || typeof sheet.addFloatDomToRange !== 'function') {
    univerAPI.showMessage({ content: 'Text overlays are unavailable in this editor build.', type: 'warning' })
    return false
  }

  const text = window.prompt('Enter text for the worksheet text box.', 'Add your text here')
  if (text == null) return false
  if (!text.trim()) {
    univerAPI.showMessage({ content: 'Enter some text before inserting the text box.', type: 'warning' })
    return false
  }

  const range = sheet.getActiveRange() || sheet.getRange('A1')
  const id = 'gridnova-text-' + (globalThis.crypto?.randomUUID?.() || Date.now())
  const drawing = sheet.addFloatDomToRange(
    range,
    { componentKey: 'gridnova-text-box', data: { text: text.trim() }, allowTransform: true },
    { width: 300, height: 100 },
    id,
  )
  if (!drawing) {
    univerAPI.showMessage({ content: 'The text box could not be inserted.', type: 'error' })
    return false
  }
  univerAPI.showMessage({ content: 'Text box inserted over the selected range. Drag or resize it on the sheet.', type: 'success' })
  return true
}

function insertSymbol(univerAPI) {
  const sheet = univerAPI.getActiveWorkbook()?.getActiveSheet()
  if (!sheet || typeof sheet.addFloatDomToRange !== 'function') {
    univerAPI.showMessage({ content: 'Symbol overlays are unavailable in this editor build.', type: 'warning' })
    return false
  }

  const symbol = window.prompt('Enter a symbol or Unicode character to insert.', '★')
  if (symbol == null) return false
  if (!symbol.trim()) {
    univerAPI.showMessage({ content: 'Enter a symbol before inserting it.', type: 'warning' })
    return false
  }

  const range = sheet.getActiveRange() || sheet.getRange('A1')
  const id = 'gridnova-symbol-' + (globalThis.crypto?.randomUUID?.() || Date.now())
  const drawing = sheet.addFloatDomToRange(
    range,
    { componentKey: 'gridnova-symbol', data: { text: symbol.trim().slice(0, 8) }, allowTransform: true },
    { width: 100, height: 100 },
    id,
  )
  if (!drawing) {
    univerAPI.showMessage({ content: 'The symbol could not be inserted.', type: 'error' })
    return false
  }
  univerAPI.showMessage({ content: 'Symbol inserted over the selected range. Drag or resize it on the sheet.', type: 'success' })
  return true
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

async function insertAutoSum(univerAPI) {
  try {
    const result = await univerAPI.executeCommand('formula-ui.operation.insert-function', { value: 'SUM' })
    if (result === false) {
      univerAPI.showMessage({ content: 'Select a cell or numeric range before using AutoSum.', type: 'warning' })
    }
  } catch (error) {
    univerAPI.showMessage({ content: error instanceof Error ? error.message : 'AutoSum could not be inserted.', type: 'error' })
  }
}

function goToCell(univerAPI) {
  const sheet = univerAPI.getActiveWorkbook()?.getActiveSheet()
  if (!sheet) return
  const address = window.prompt('Enter a cell or range address (for example, C12 or A1:D8).', 'A1')
  if (address == null || !address.trim()) return
  try {
    sheet.getRange(address.trim()).activate()
  } catch (error) {
    univerAPI.showMessage({ content: error instanceof Error ? error.message : 'The cell or range address is invalid.', type: 'warning' })
  }
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

function hideSelectedColumns(univerAPI) {
  const sheet = univerAPI.getActiveWorkbook()?.getActiveSheet()
  const range = sheet?.getActiveRange()
  if (!sheet || !range) return
  sheet.hideColumns(range.getColumn(), range.getLastColumn() - range.getColumn() + 1)
}

function showSelectedColumns(univerAPI) {
  const sheet = univerAPI.getActiveWorkbook()?.getActiveSheet()
  const range = sheet?.getActiveRange()
  if (!sheet || !range) return
  sheet.showColumns(range.getColumn(), range.getLastColumn() - range.getColumn() + 1)
}

function setSelectedRowHeight(univerAPI) {
  const sheet = univerAPI.getActiveWorkbook()?.getActiveSheet()
  const range = sheet?.getActiveRange()
  if (!sheet || !range) return
  const start = range.getRow()
  const count = range.getLastRow() - start + 1
  const current = sheet.getRowHeight?.(start) || 23
  const value = window.prompt('Set selected row height in pixels.', String(Math.round(current)))
  if (value == null) return
  const height = Number(value)
  if (!Number.isFinite(height) || height < 10 || height > 409) {
    univerAPI.showMessage({ content: 'Row height must be between 10 and 409 pixels.', type: 'warning' })
    return
  }
  sheet.setRowHeights(start, count, height)
}

function setSelectedColumnWidth(univerAPI) {
  const sheet = univerAPI.getActiveWorkbook()?.getActiveSheet()
  const range = sheet?.getActiveRange()
  if (!sheet || !range) return
  const start = range.getColumn()
  const count = range.getLastColumn() - start + 1
  const current = sheet.getColumnWidth?.(start) || 100
  const value = window.prompt('Set selected column width in pixels.', String(Math.round(current)))
  if (value == null) return
  const width = Number(value)
  if (!Number.isFinite(width) || width < 10 || width > 2048) {
    univerAPI.showMessage({ content: 'Column width must be between 10 and 2048 pixels.', type: 'warning' })
    return
  }
  sheet.setColumnWidths(start, count, width)
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

  const lastUsedColumn = sheet.getLastColumn?.() ?? -1
  if (range.getColumn() !== 0 || range.getLastColumn() !== lastUsedColumn) {
    univerAPI.showMessage({
      content: 'To remove worksheet rows safely, select the full data width from column A through the last used column.',
      type: 'warning',
    })
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

function columnIndexToLetters(index) {
  let value = index + 1
  let letters = ''
  while (value > 0) {
    const remainder = (value - 1) % 26
    letters = String.fromCharCode(65 + remainder) + letters
    value = Math.floor((value - 1) / 26)
  }
  return letters
}

function insertSubtotalRow(univerAPI) {
  const sheet = univerAPI.getActiveWorkbook()?.getActiveSheet()
  const range = sheet?.getActiveRange()
  if (!sheet || !range) {
    univerAPI.showMessage({ content: 'Select a contiguous data range before inserting a subtotal.', type: 'warning' })
    return false
  }

  const values = range.getValues()
  if (!Array.isArray(values) || values.length < 2) {
    univerAPI.showMessage({ content: 'Select a header row and at least one data row.', type: 'warning' })
    return false
  }

  const hasHeader = window.confirm('Does the first selected row contain column headers?')
  const dataOffset = hasHeader ? 1 : 0
  const firstDataRow = range.getRow() + dataOffset
  const lastDataRow = range.getLastRow()
  if (firstDataRow > lastDataRow) {
    univerAPI.showMessage({ content: 'The selection does not contain any data rows.', type: 'warning' })
    return false
  }

  const aggregate = window.prompt('Subtotal function: SUM, AVERAGE, COUNT, MAX or MIN.', 'SUM')
  if (aggregate == null) return false
  const functionName = aggregate.trim().toUpperCase()
  const subtotalCodes = { AVERAGE: 1, COUNT: 2, MAX: 4, MIN: 5, SUM: 9 }
  const code = subtotalCodes[functionName]
  if (!code) {
    univerAPI.showMessage({ content: 'Choose SUM, AVERAGE, COUNT, MAX or MIN.', type: 'warning' })
    return false
  }

  const afterRow = range.getLastRow()
  const maxRows = sheet.getMaxRows?.() ?? 0
  if (afterRow >= maxRows - 1) {
    univerAPI.showMessage({ content: 'There is no available row below the selected range.', type: 'warning' })
    return false
  }

  const firstColumn = range.getColumn()
  const lastColumn = range.getLastColumn()
  const columnCount = lastColumn - firstColumn + 1
  const dataRows = values.slice(dataOffset)
  const numericColumns = []

  for (let columnOffset = 0; columnOffset < columnCount; columnOffset++) {
    const columnValues = dataRows.map((row) => row?.[columnOffset])
    const populated = columnValues.filter((value) => value !== null && value !== undefined && value !== '')
    if (populated.length && populated.some((value) => Number.isFinite(typeof value === 'number' ? value : Number(String(value).replace(/,/g, ''))))) {
      numericColumns.push(columnOffset)
    }
  }

  if (!numericColumns.length) {
    univerAPI.showMessage({ content: 'No numeric columns were found in the selected range.', type: 'warning' })
    return false
  }

  try {
    sheet.insertRowAfter(afterRow, 1)
    const subtotalRow = afterRow + 1
    const firstColumnValues = dataRows.map((row) => row?.[0])
    const firstColumnIsNumeric = firstColumnValues.some((value) =>
      value !== null && value !== undefined && value !== '' && Number.isFinite(typeof value === 'number' ? value : Number(String(value).replace(/,/g, '')))
    )

    if (!firstColumnIsNumeric) {
      sheet.getRange(subtotalRow, firstColumn).setValue('Subtotal')
    }

    numericColumns.forEach((columnOffset) => {
      const columnIndex = firstColumn + columnOffset
      const columnLetters = columnIndexToLetters(columnIndex)
      const formula = '=SUBTOTAL(' + code + ',' + columnLetters + (firstDataRow + 1) + ':' + columnLetters + (lastDataRow + 1) + ')'
      sheet.getRange(subtotalRow, columnIndex).setFormula(formula)
    })

    sheet.getRange(subtotalRow, firstColumn, 1, columnCount).setFontWeight('bold')
    univerAPI.showMessage({
      content: functionName + ' subtotal row inserted below the selected data.',
      type: 'success',
    })
    return true
  } catch (error) {
    univerAPI.showMessage({
      content: error instanceof Error ? error.message : 'Could not insert the subtotal row.',
      type: 'error',
    })
    return false
  }
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
          { id: 'import-workbook', title: 'Open Workbook (XLSX / CSV)', icon: 'import', action: 'import' },
          { id: 'save', title: 'Save', icon: 'save', action: 'save' },
          { id: 'save-as', title: 'Save As', icon: 'copy', action: 'save-as' },
          { id: 'export-xlsx', title: 'Export Excel', icon: 'export', action: 'export-xlsx' },
          { id: 'export-csv', title: 'Export Active Sheet CSV', icon: 'export', action: 'export-csv' },
        ] },
        { id: 'information', title: 'Workbook Information', items: [
          { id: 'copy-id', title: 'Copy Workbook ID', action: 'copy-workbook-id' },
        ] },
        { id: 'print', title: 'Print', items: [
          { id: 'print-workbook', title: 'Print via Browser Dialog', icon: 'print', action: 'print' },
        ] },
      ],
    },
    {
      id: 'share', title: 'Share', order: 4, groups: [
        { id: 'share-copy', title: 'Share & Access', items: [
          { id: 'share-file', title: 'Share Workbook File', icon: 'share', action: 'share-workbook' },
          { id: 'copy-link', title: 'Copy Workbook Link', icon: 'link', action: 'copy-workbook-link' },
          { id: 'share-comments', title: 'Comments', icon: 'review', action: 'open-comments' },
          { id: 'manage-access', title: 'Manage Access (Unavailable)', action: 'manage-access-unavailable' },
        ] },
      ],
    },
    {
      id: 'page-layout', title: 'Page Layout', order: 5, groups: [
        { id: 'page-setup', title: 'Page Setup', items: [
          { id: 'print-layout', title: 'Page Setup / Print Preview (Unavailable)', icon: 'layout', action: 'page-layout-unavailable' },
          { id: 'print', title: 'Print via Browser Dialog', icon: 'print', action: 'print' },
          { id: 'orientation', title: 'Page Orientation', icon: 'layout', run: setPrintOrientation },
          { id: 'paper-size', title: 'Paper Size', icon: 'layout', run: setPrintPaperSize },
          { id: 'margins', title: 'Margins', icon: 'layout', run: setPrintMargins },
          { id: 'print-area', title: 'Print Area (Unavailable)', action: 'page-layout-unavailable' },
          { id: 'page-breaks', title: 'Manual Page Breaks (Unavailable)', action: 'page-layout-unavailable' },
        ] },
        { id: 'scale-to-fit', title: 'Scale to Fit', items: [
          { id: 'print-scale', title: 'Print Scale', icon: 'layout', run: setPrintScale },
          { id: 'fit-width', title: 'Fit to Width (Unavailable)', action: 'page-layout-unavailable' },
          { id: 'fit-height', title: 'Fit to Height (Unavailable)', action: 'page-layout-unavailable' },
        ] },
        { id: 'sheet-background', title: 'Sheet Background', items: [
          { id: 'print-background', title: 'Print Background Colors', icon: 'paint', run: togglePrintBackground },
          { id: 'background-image', title: 'Sheet Background Image (Unavailable)', action: 'page-layout-unavailable' },
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
          { id: 'sheet-permissions', title: 'Sheet Permission Settings', command: 'sheet-permission.operation.openDialog' },
          { id: 'protect-workbook-unavailable', title: 'Protect Workbook (Unavailable)', action: 'sheet-protection-unavailable' },
          { id: 'show-changes', title: 'Show Changes (Unavailable)', action: 'review-tools-unavailable' },
        ] },
        { id: 'history', title: 'History', items: [
          { id: 'edit-history', title: 'Local Activity Feed', action: 'open-catch-up' },
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
          { id: 'pen', title: 'Pen', icon: 'draw', run: (api) => insertInkCanvas(api, 'pen') },
          { id: 'highlighter', title: 'Highlighter', icon: 'paint', run: (api) => insertInkCanvas(api, 'highlighter') },
          { id: 'eraser', title: 'Eraser (Last Stroke)', icon: 'cut', run: eraseLastInkStroke },
          { id: 'clear-canvas', title: 'Clear Latest Canvas', icon: 'cut', run: clearLatestInkCanvas },
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
          { id: 'shape-fill', title: 'Fill Color', icon: 'fillColor', run: setGridNovaShapeFill },
        { id: 'shape-outline', title: 'Outline Color', icon: 'border', run: setGridNovaShapeOutline },
        { id: 'shape-properties', title: 'Size / Rotation Properties', icon: 'layout', run: setGridNovaShapeProperties },
        { id: 'shape-delete', title: 'Delete Selected Shape', icon: 'cut', run: deleteGridNovaShape },
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
    { id: 'paste-values', title: 'Paste Values', icon: 'paste', command: 'sheet.command.paste-value', gridLayout: { row: 2, column: 3 } },
    { id: 'paste-formats', title: 'Paste Formats', icon: 'paint', command: 'sheet.command.paste-format', gridLayout: { row: 2, column: 4 } },
    { id: 'paste-column-width', title: 'Paste Column Width', icon: 'layout', command: 'sheet.command.paste-col-width', gridLayout: { row: 2, column: 5 } },
  ].map((item) => ({
    ...item,
    gridLayout: { ...item.gridLayout, showLabel: true },
  })).forEach((item) => {
    clipboardGroup.addSubmenu(univerAPI.createMenu({
      id: 'gridnova.ribbon.home.clipboard.' + item.id,
      title: item.title,
      tooltip: item.title,
      icon: item.icon ? 'gridnova-icon-' + item.icon : undefined,
      gridLayout: item.gridLayout,
      action: () => executeRibbonItem(univerAPI, item, onRibbonAction),
    }))
  })
  clipboardGroup.appendTo('ribbon.start')

  const homeGroups = [
    {
      id: 'styles', title: 'Styles', order: 5, items: [
        { id: 'conditional-formatting', title: 'Conditional Formatting', icon: 'layout', gridLayout: { row: 1, column: 1 }, command: 'sheet.operation.open.conditional.formatting.panel', params: { value: 3 } },
        { id: 'format-as-table', title: 'Format as Table', icon: 'table', gridLayout: { row: 1, column: 2 }, run: formatSelectionAsTable },
        { id: 'cell-style-title', title: 'Title Style', icon: 'paint', run: (api) => applyCellStylePreset(api, 'title') },
        { id: 'cell-style-heading', title: 'Heading Style', icon: 'paint', run: (api) => applyCellStylePreset(api, 'heading') },
        { id: 'cell-style-input', title: 'Input Style', icon: 'paint', run: (api) => applyCellStylePreset(api, 'input') },
        { id: 'cell-style-good', title: 'Good Style', icon: 'paint', run: (api) => applyCellStylePreset(api, 'good') },
        { id: 'cell-style-bad', title: 'Bad Style', icon: 'paint', run: (api) => applyCellStylePreset(api, 'bad') },
        { id: 'cell-style-neutral', title: 'Neutral Style', icon: 'paint', run: (api) => applyCellStylePreset(api, 'neutral') },
        { id: 'number-format-settings', title: 'Number Format Settings', icon: 'calculator', command: 'sheet.operation.open.numfmt.panel' },
        { id: 'format-cells', title: 'Format Cells / Number Format', icon: 'layout', run: openFormatCellsDialog },
      ],
    },
    {
      id: 'cells', title: 'Cells', order: 6, items: [
        { id: 'insert-row', title: 'Insert Row', icon: 'table', command: 'sheet.command.insert-row-before' },
        { id: 'insert-column', title: 'Insert Column', icon: 'table', command: 'sheet.command.insert-col-before' },
        { id: 'delete-row', title: 'Delete Row', icon: 'cut', command: 'sheet.command.remove-row-confirm' },
        { id: 'delete-column', title: 'Delete Column', icon: 'cut', command: 'sheet.command.remove-col-confirm' },
        { id: 'row-height', title: 'Row Height', icon: 'layout', run: setSelectedRowHeight },
        { id: 'column-width', title: 'Column Width', icon: 'layout', run: setSelectedColumnWidth },
        { id: 'hide-rows', title: 'Hide Rows', icon: 'layout', run: hideSelectedRows },
        { id: 'unhide-rows', title: 'Unhide Rows', icon: 'layout', run: showSelectedRows },
        { id: 'hide-columns', title: 'Hide Columns', icon: 'layout', run: hideSelectedColumns },
        { id: 'unhide-columns', title: 'Unhide Columns', icon: 'layout', run: showSelectedColumns },
      ],
    },
    {
      id: 'editing', title: 'Editing', order: 7, items: [
        { id: 'auto-sum', title: 'AutoSum', icon: 'calculator', run: insertAutoSum },
        { id: 'sort-ascending', title: 'Sort A to Z', icon: 'sort', command: 'sheet.command.sort-range-asc' },
        { id: 'sort-descending', title: 'Sort Z to A', icon: 'sort', command: 'sheet.command.sort-range-desc' },
        { id: 'filter', title: 'Filter', icon: 'filter', command: 'sheet.command.smart-toggle-filter' },
        { id: 'clear-contents', title: 'Clear Contents', icon: 'table', command: 'sheet.command.clear-selection-content' },
        { id: 'clear-formats', title: 'Clear Formats', icon: 'paint', command: 'sheet.command.clear-selection-format' },
        { id: 'clear-all', title: 'Clear All', icon: 'table', command: 'sheet.command.clear-selection-all' },
        { id: 'find', title: 'Find', icon: 'search', command: 'ui.operation.open-find-dialog' },
        { id: 'replace', title: 'Replace', icon: 'search', command: 'ui.operation.open-find-dialog' },
        { id: 'go-to', title: 'Go To', icon: 'search', run: goToCell },
        { id: 'fill', title: 'Fill (Unavailable)', icon: 'table', action: 'fill-unavailable' },
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
        { id: 'insert-rectangle', title: 'Rectangle Shape', icon: 'draw', run: (api) => insertDrawingShape(api, 'rectangle') },
        { id: 'insert-ellipse', title: 'Ellipse Shape', icon: 'draw', run: (api) => insertDrawingShape(api, 'ellipse') },
        { id: 'insert-line', title: 'Line Shape', icon: 'draw', run: (api) => insertDrawingShape(api, 'line') },
      ],
    },
    {
      path: 'ribbon.insert', id: 'charts', title: 'Charts', order: 2, items: [
        { id: 'insert-bar-chart', title: 'Bar Chart Overlay', icon: 'table', run: (api) => insertChartOverlay(api, 'bar') },
        { id: 'insert-column-chart', title: 'Column Chart Overlay', icon: 'table', run: (api) => insertChartOverlay(api, 'column') },
        { id: 'insert-line-chart', title: 'Line Chart Overlay', icon: 'table', run: (api) => insertChartOverlay(api, 'line') },
        { id: 'insert-pie-chart', title: 'Pie Chart Overlay', icon: 'table', run: (api) => insertChartOverlay(api, 'pie') },
        { id: 'insert-area-chart', title: 'Area Chart Overlay', icon: 'table', run: (api) => insertChartOverlay(api, 'area') },
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
        { id: 'insert-text-box', title: 'Text Box', icon: 'text', run: insertTextBox },
        { id: 'insert-symbol', title: 'Symbol', icon: 'text', run: insertSymbol },
      ],
    },
    {
      path: 'ribbon.formulas', id: 'function-library', title: 'Function Library', order: 0, items: [
        { id: 'insert-function', title: 'Insert Function', icon: 'calculator', run: (api) => insertFormulaFunction(api, 'SUM') },
        { id: 'auto-sum', title: 'AutoSum', icon: 'calculator', run: insertAutoSum },
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
      path: 'ribbon.formulas', id: 'defined-names', title: 'Defined Names', order: 8, items: [
        { id: 'manage-names', title: 'Name Manager', command: 'sidebar.operation.defined-name', params: { value: 'open' } },
      ],
    },
    {
      path: 'ribbon.formulas', id: 'formula-auditing', title: 'Formula Auditing', order: 9, items: [
        { id: 'trace-formulas-unavailable', title: 'Trace Precedents / Dependents (Unavailable)', action: 'formula-feature-unavailable' },
        { id: 'show-formulas-unavailable', title: 'Show Formulas (Unavailable)', action: 'formula-feature-unavailable' },
      ],
    },
    {
      path: 'ribbon.formulas', id: 'calculation', title: 'Calculation', order: 10, items: [
        { id: 'recalculate', title: 'Recalculate Now', icon: 'calculator', run: recalculateFormulas },
        { id: 'calculation-options', title: 'Calculation Options (Unavailable)', action: 'formula-feature-unavailable' },
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
        { id: 'data-validation', title: 'Data Validation', command: 'data-validation.operation.open-validation-panel', params: { isAdd: true } },
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
        { id: 'group-unavailable', title: 'Group Rows (Unavailable)', action: 'data-outline-unavailable' },
        { id: 'ungroup-unavailable', title: 'Ungroup Rows (Unavailable)', action: 'data-outline-unavailable' },
        { id: 'insert-subtotal', title: 'Insert Subtotal Row', icon: 'calculator', run: insertSubtotalRow },
      ],
    },
    {
      path: 'ribbon.view', id: 'show', title: 'Show', order: 0, items: [
        { id: 'toggle-gridlines', title: 'Gridlines', icon: 'layout', run: toggleGridlines },
        { id: 'formula-bar-toggle', title: 'Show / Hide Formula Bar', icon: 'text', action: 'toggle-formula-bar' },
        { id: 'headings-unavailable', title: 'Show / Hide Headings (Unavailable)', action: 'view-tools-unavailable' },
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
    withRibbonGridLayouts(group.items).forEach((item) => {
      groupMenu.addSubmenu(univerAPI.createMenu({
        id: 'gridnova.ribbon.' + group.id + '.' + item.id,
        title: item.title,
        tooltip: item.title,
        icon: item.icon ? 'gridnova-icon-' + item.icon : undefined,
        gridLayout: item.gridLayout,
        action: () => executeRibbonItem(univerAPI, item, onRibbonAction),
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
    withRibbonGridLayouts(group.items).forEach((item) => {
      groupMenu.addSubmenu(univerAPI.createMenu({
        id: 'gridnova.ribbon.home.' + item.id,
        title: item.title,
        tooltip: item.title,
        icon: item.icon ? 'gridnova-icon-' + item.icon : undefined,
        gridLayout: item.gridLayout,
        action: () => executeRibbonItem(univerAPI, item, onRibbonAction),
      }))
    })
    groupMenu.appendTo('ribbon.start')
    return groupMenu
  })

  univerAPI.updateMenuConfig({
    'ribbon.start': { title: 'Home', order: 1 },
    'ribbon.insert': { title: 'Insert', order: 2, hidden: false },
    'ribbon.start.clipboard': { hidden: true },
    'ribbon.start.history': { hidden: true },
    'ribbon.start.format': { title: 'Font', order: 2 },
    'ribbon.start.layout': { title: 'Alignment', order: 3 },
    'ribbon.start.others': { title: 'More Editing', order: 8 },
    'univer.command.undo': { icon: 'gridnova-icon-undo' },
    'univer.command.redo': { icon: 'gridnova-icon-redo' },
    'ui.operation.activate-format-painter': { icon: 'gridnova-icon-paint' },
    'sheet.command.set-once-format-painter': { icon: 'gridnova-icon-paint' },
    'sheet.command.set-range-bold': { icon: 'gridnova-icon-bold', gridLayout: { row: 1, column: 2, showLabel: true } },
    'sheet.command.set-range-font-increase': { icon: 'gridnova-icon-calculator', gridLayout: { row: 1, column: 6, showLabel: true } },
    'sheet.command.set-range-font-decrease': { icon: 'gridnova-icon-calculator', gridLayout: { row: 2, column: 6, showLabel: true } },
    'sheet.command.set-range-italic': { icon: 'gridnova-icon-italic', gridLayout: { row: 2, column: 2, showLabel: true } },
    // Place expanded Home controls into a labeled two-row grid. Single Line mode
    // ignores these grid coordinates and keeps the native one-row toolbar.
    'sheet.command.set-range-font-family': { order: 0, gridLayout: { row: 1, column: 1, width: 150, showLabel: true } },
    'sheet.command.set-range-fontsize': { order: 1, gridLayout: { row: 2, column: 1, width: 58, showLabel: true } },
    'sheet.command.add-worksheet-merge': { icon: 'gridnova-icon-layout', gridLayout: { row: 1, column: 4, showLabel: true } },
    'sheet.command.add-worksheet-merge-all': { icon: 'gridnova-icon-layout', gridLayout: { row: 2, column: 4, showLabel: true } },
    'sheet.command.add-worksheet-merge-vertical': { icon: 'gridnova-icon-layout', gridLayout: { row: 1, column: 5, showLabel: true } },
    'sheet.command.add-worksheet-merge-horizontal': { icon: 'gridnova-icon-layout', gridLayout: { row: 2, column: 5, showLabel: true } },
    'sheet.command.remove-worksheet-merge': { icon: 'gridnova-icon-layout', gridLayout: { row: 1, column: 6, showLabel: true } },
    'sheet.command.set-text-rotation': { icon: 'gridnova-icon-layout', gridLayout: { row: 2, column: 6, showLabel: true } },
    'sheet.command.set-range-underline': { icon: 'gridnova-icon-underline', gridLayout: { row: 1, column: 3, showLabel: true } },
    'sheet.command.set-range-stroke': { icon: 'gridnova-icon-strike', gridLayout: { row: 2, column: 3, showLabel: true } },
    'sheet.command.set-range-text-color': { icon: 'gridnova-icon-textColor', gridLayout: { row: 1, column: 4, showLabel: true } },
    'sheet.command.set-background-color': { icon: 'gridnova-icon-fillColor', gridLayout: { row: 2, column: 4, showLabel: true } },
    'sheet.command.set-border-basic': { icon: 'gridnova-icon-border', gridLayout: { row: 1, column: 5, showLabel: true } },
    'sheet.command.set-horizontal-text-align': { icon: 'gridnova-icon-alignLeft', gridLayout: { row: 1, column: 1, showLabel: true } },
    'sheet.command.set-vertical-text-align': { icon: 'gridnova-icon-alignCenter', gridLayout: { row: 2, column: 1, showLabel: true } },
    'sheet.command.set-text-wrap': { icon: 'gridnova-icon-alignRight', gridLayout: { row: 1, column: 2, showLabel: true } },
    'sheet.command.copy': { icon: 'gridnova-icon-copy' },
    'sheet.command.cut': { icon: 'gridnova-icon-cut' },
    'sheet.command.paste': { icon: 'gridnova-icon-paste' },
    'sheet.command.smart-toggle-filter': { icon: 'gridnova-icon-filter' },
    'sheet.command.sort-range-asc': { icon: 'gridnova-icon-sort' },
    'ui.operation.open-find-dialog': { icon: 'gridnova-icon-search' },
    'sheet.command.numfmt.set.currency': { icon: 'gridnova-icon-currency', gridLayout: { row: 2, column: 1, showLabel: true } },
    'sheet.command.numfmt.set.percent': { icon: 'gridnova-icon-percent', gridLayout: { row: 1, column: 2, showLabel: true } },
    'sheet.command.numfmt.add.decimal.command': { icon: 'gridnova-icon-calculator', gridLayout: { row: 2, column: 2, showLabel: true } },
    'sheet.command.numfmt.subtract.decimal.command': { icon: 'gridnova-icon-calculator', gridLayout: { row: 1, column: 3, showLabel: true } },
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
            formulaBar: true,
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
      applySavedWorkbookPrintSettings(univerAPI)
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
        setRibbonType(type) {
          univerAPI.setRibbonType(type)
        },
        setFormulaBarVisible(visible) {
          const formulaBar = containerRef.current?.querySelector('[data-u-comp="formula-bar"]')
          if (!formulaBar) return false
          formulaBar.style.display = visible ? '' : 'none'
          return true
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
