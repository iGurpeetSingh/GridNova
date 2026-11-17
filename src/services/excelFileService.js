import ExcelJS from 'exceljs'

const DEFAULT_ROWS = 1000
const DEFAULT_COLUMNS = 26

function colorToHex(color) {
  const raw = color?.argb || color?.rgb
  if (!raw) return undefined
  const value = String(raw).replace(/^#/, '')
  return value.length === 8 ? '#' + value.slice(2) : '#' + value
}

function excelDateToSerial(value) {
  return value.getTime() / 86400000 + 25569
}

function normalizeImportedValue(value) {
  if (value instanceof Date) return excelDateToSerial(value)
  if (value && typeof value === 'object') {
    if (Array.isArray(value.richText)) return value.richText.map((part) => part.text || '').join('')
    if (value.text && value.hyperlink) return value.text
    if (value.formula || value.sharedFormula) {
      return {
        v: value.result instanceof Date ? excelDateToSerial(value.result) : value.result ?? null,
        f: value.formula || value.sharedFormula,
      }
    }
    if (value.error) return value.error
  }
  return value ?? null
}

function mapExcelStyle(cell) {
  const style = {}
  const font = cell.font || {}
  const fill = cell.fill || {}
  const alignment = cell.alignment || {}
  if (font.name) style.ff = font.name
  if (font.size) style.fs = font.size
  if (font.bold) style.bl = 1
  if (font.italic) style.it = 1
  if (font.underline) style.ul = { s: 1, c: 1, t: 0 }
  if (font.strike) style.st = { s: 1, c: 1, t: 0 }
  const fontColor = colorToHex(font.color)
  if (fontColor) style.cl = { rgb: fontColor }
  const fillColor = colorToHex(fill.fgColor)
  if (fillColor && fill.type === 'pattern' && fill.pattern === 'solid') style.bg = { rgb: fillColor }
  if (alignment.horizontal) style.ht = alignment.horizontal
  if (alignment.vertical) style.vt = alignment.vertical
  if (alignment.wrapText) style.tb = 2
  if (cell.numFmt) style.n = { pattern: cell.numFmt }
  return Object.keys(style).length ? style : undefined
}

function toSheetId(name, index) {
  const safe = String(name || 'Sheet' + (index + 1)).replace(/[^a-zA-Z0-9_-]/g, '').slice(0, 20)
  return 'sheet_' + (index + 1) + '_' + (safe || 'data')
}

function createSheetData(id, name, cellData, rowCount = DEFAULT_ROWS, columnCount = DEFAULT_COLUMNS) {
  return {
    id,
    name: String(name || 'Sheet1').slice(0, 31),
    tabColor: '',
    hidden: 0,
    rowCount: Math.max(DEFAULT_ROWS, rowCount || 0),
    columnCount: Math.max(DEFAULT_COLUMNS, columnCount || 0),
    zoomRatio: 1,
    freeze: { startRow: -1, startColumn: -1, ySplit: 0, xSplit: 0 },
    scrollTop: 0,
    scrollLeft: 0,
    defaultColumnWidth: 100,
    defaultRowHeight: 23,
    mergeData: [],
    cellData,
    rowData: {},
    columnData: {},
    rowHeader: { width: 46, hidden: 0 },
    columnHeader: { height: 20, hidden: 0 },
    showGridlines: 1,
    rightToLeft: 0,
  }
}

function parseRange(address) {
  const match = /^([A-Z]+)(\\d+):([A-Z]+)(\\d+)$/i.exec(address)
  if (!match) return null
  const toColumn = (letters) => letters.toUpperCase().split('').reduce((total, char) => total * 26 + char.charCodeAt(0) - 64, 0)
  return {
    startRow: Number(match[2]) - 1,
    startColumn: toColumn(match[1]) - 1,
    endRow: Number(match[4]) - 1,
    endColumn: toColumn(match[3]) - 1,
  }
}

function parseDelimitedText(text, delimiter) {
  const rows = []
  let row = []
  let value = ''
  let quoted = false
  for (let index = 0; index < text.length; index += 1) {
    const char = text[index]
    if (char === '"') {
      if (quoted && text[index + 1] === '"') {
        value += '"'
        index += 1
      } else quoted = !quoted
    } else if (char === delimiter && !quoted) {
      row.push(value)
      value = ''
    } else if ((char === '\\n' || char === '\\r') && !quoted) {
      if (char === '\\r' && text[index + 1] === '\\n') index += 1
      row.push(value)
      if (row.some((item) => item !== '') || rows.length === 0) rows.push(row)
      row = []
      value = ''
    } else value += char
  }
  row.push(value)
  if (row.some((item) => item !== '') || row.length > 1) rows.push(row)
  return rows
}

export async function importExcelFile(file) {
  if (!file) throw new Error('Choose an Excel or CSV file first.')
  const fileName = file.name || 'Imported workbook.xlsx'
  const lowerName = fileName.toLowerCase()
  const workbook = new ExcelJS.Workbook()
  const buffer = await file.arrayBuffer()

  if (lowerName.endsWith('.csv') || lowerName.endsWith('.tsv')) {
    const delimiter = lowerName.endsWith('.tsv') ? '\\t' : ','
    const rows = parseDelimitedText(new TextDecoder().decode(buffer).replace(/^\uFEFF/, ''), delimiter)
    const sheetId = 'sheet_1_imported'
    const cellData = {}
    rows.forEach((row, rowIndex) => {
      const cells = {}
      row.forEach((value, columnIndex) => {
        if (value !== '') cells[columnIndex] = { v: value }
      })
      if (Object.keys(cells).length) cellData[rowIndex] = cells
    })
    return {
      id: 'gridnova-' + crypto.randomUUID(),
      name: fileName.replace(/\\.(csv|tsv)$/i, ''),
      appVersion: '1.0.0',
      locale: 'enUS',
      styles: {},
      sheetOrder: [sheetId],
      sheets: {
        [sheetId]: createSheetData(sheetId, 'Sheet1', cellData, rows.length, Math.max(1, ...rows.map((row) => row.length))),
      },
    }
  }

  await workbook.xlsx.load(buffer)
  const sheetOrder = []
  const sheets = {}

  workbook.worksheets.forEach((worksheet, index) => {
    const sheetId = toSheetId(worksheet.name, index)
    sheetOrder.push(sheetId)
    const cellData = {}
    worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      const rowData = {}
      row.eachCell({ includeEmpty: false }, (cell, columnNumber) => {
        const imported = normalizeImportedValue(cell.value)
        const cellDataItem = imported && typeof imported === 'object' && Object.prototype.hasOwnProperty.call(imported, 'f')
          ? { v: imported.v, f: imported.f }
          : { v: imported }
        const style = mapExcelStyle(cell)
        if (style) cellDataItem.s = style
        if (cellDataItem.v !== null && cellDataItem.v !== undefined || cellDataItem.f) rowData[columnNumber - 1] = cellDataItem
      })
      if (Object.keys(rowData).length) cellData[rowNumber - 1] = rowData
    })

    const mergeData = []
    ;(worksheet.model?.merges || []).forEach((address) => {
      const parsed = parseRange(address)
      if (parsed) mergeData.push(parsed)
    })
    const rowHeights = {}
    worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      if (row.height) rowHeights[rowNumber - 1] = { h: row.height }
    })
    const columnWidths = {}
    worksheet.columns.forEach((column, columnIndex) => {
      if (column.width) columnWidths[columnIndex] = { w: Math.round(column.width * 7), hd: 0 }
    })
    sheets[sheetId] = {
      ...createSheetData(sheetId, worksheet.name, cellData, worksheet.rowCount, worksheet.columnCount),
      mergeData,
      rowData: rowHeights,
      columnData: columnWidths,
    }
  })

  if (!sheetOrder.length) {
    const sheetId = 'sheet_1_data'
    sheetOrder.push(sheetId)
    sheets[sheetId] = createSheetData(sheetId, 'Sheet1', {}, DEFAULT_ROWS, DEFAULT_COLUMNS)
  }
  return {
    id: 'gridnova-' + crypto.randomUUID(),
    name: fileName.replace(/\\.(xlsx|xls)$/i, ''),
    appVersion: '1.0.0',
    locale: 'enUS',
    styles: {},
    sheetOrder,
    sheets,
  }
}

function getStyle(cell, styles = {}) {
  if (!cell?.s) return {}
  return typeof cell.s === 'string' ? styles[cell.s] || {} : cell.s
}

function toArgb(color) {
  const value = color?.rgb || color
  if (!value) return undefined
  const hex = String(value).replace('#', '')
  return hex.length === 6 ? 'FF' + hex : hex
}

function applyExportStyle(cell, source, styles) {
  const style = getStyle(source, styles)
  if (style.ff || style.fs || style.bl || style.it || style.ul || style.st || style.cl) {
    cell.font = {
      ...(style.ff ? { name: style.ff } : {}),
      ...(style.fs ? { size: style.fs } : {}),
      ...(style.bl ? { bold: true } : {}),
      ...(style.it ? { italic: true } : {}),
      ...(style.ul ? { underline: true } : {}),
      ...(style.st ? { strike: true } : {}),
      ...(style.cl?.rgb ? { color: { argb: toArgb(style.cl) } } : {}),
    }
  }
  if (style.bg?.rgb) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: toArgb(style.bg) } }
  if (style.ht || style.vt || style.tb) {
    cell.alignment = {
      ...(style.ht ? { horizontal: style.ht } : {}),
      ...(style.vt ? { vertical: style.vt } : {}),
      ...(style.tb ? { wrapText: true } : {}),
    }
  }
  if (style.n?.pattern) cell.numFmt = style.n.pattern
}

function toColumnLetters(number) {
  let value = number + 1
  let output = ''
  while (value > 0) {
    const remainder = (value - 1) % 26
    output = String.fromCharCode(65 + remainder) + output
    value = Math.floor((value - 1) / 26)
  }
  return output
}

export async function exportExcelFile(snapshot, fileName = 'GridNova Workbook.xlsx') {
  if (!snapshot?.sheetOrder?.length) throw new Error('There is no workbook data to export.')
  const workbook = new ExcelJS.Workbook()
  workbook.creator = 'GridNova'
  workbook.created = new Date()

  snapshot.sheetOrder.forEach((sheetId) => {
    const source = snapshot.sheets?.[sheetId]
    if (!source) return
    const worksheet = workbook.addWorksheet(String(source.name || 'Sheet').slice(0, 31))
    Object.entries(source.cellData || {}).forEach(([rowKey, columns]) => {
      Object.entries(columns || {}).forEach(([columnKey, cellData]) => {
        const cell = worksheet.getCell(Number(rowKey) + 1, Number(columnKey) + 1)
        const formula = cellData?.f
        const value = cellData?.v
        cell.value = formula
          ? { formula: String(formula).replace(/^=/, ''), result: value ?? null }
          : value ?? null
        applyExportStyle(cell, cellData, snapshot.styles)
      })
    })
    Object.entries(source.rowData || {}).forEach(([rowIndex, rowData]) => {
      if (rowData?.h) worksheet.getRow(Number(rowIndex) + 1).height = rowData.h
    })
    Object.entries(source.columnData || {}).forEach(([columnIndex, columnData]) => {
      if (columnData?.w) worksheet.getColumn(Number(columnIndex) + 1).width = columnData.w / 7
    })
    ;(source.mergeData || []).forEach((range) => {
      worksheet.mergeCells(
        toColumnLetters(range.startColumn) + (range.startRow + 1) + ':' +
        toColumnLetters(range.endColumn) + (range.endRow + 1),
      )
    })
  })

  const buffer = await workbook.xlsx.writeBuffer()
  downloadBlob(buffer, fileName, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
}

export function exportCsvFile(snapshot, fileName = 'GridNova Sheet.csv', sheetId) {
  const id = sheetId || snapshot?.sheetOrder?.[0]
  const sheet = snapshot?.sheets?.[id]
  if (!sheet) throw new Error('No active worksheet is available for CSV export.')
  const cellData = sheet.cellData || {}
  const rowIndexes = Object.keys(cellData).map(Number)
  const maxRow = Math.max(0, ...rowIndexes)
  const maxColumn = Math.max(0, ...rowIndexes.flatMap((row) => Object.keys(cellData[row] || {}).map(Number)))
  const rows = []
  for (let row = 0; row <= maxRow; row += 1) {
    const values = []
    for (let column = 0; column <= maxColumn; column += 1) {
      const cell = cellData[row]?.[column]
      values.push(cell?.v ?? (cell?.f ? '=' + cell.f : ''))
    }
    rows.push(values.map(escapeCsvValue).join(','))
  }
  downloadBlob(new Blob(['\\uFEFF' + rows.join('\\r\\n')], { type: 'text/csv;charset=utf-8' }), fileName, 'text/csv;charset=utf-8')
}

function escapeCsvValue(value) {
  const text = String(value ?? '')
  return /[",\\r\\n]/.test(text) ? '"' + text.replace(/"/g, '""') + '"' : text
}

function downloadBlob(content, fileName, type) {
  const blob = content instanceof Blob ? content : new Blob([content], { type })
  const url = URL.createObjectURL(blob)
  const anchor = document.createElement('a')
  anchor.href = url
  anchor.download = fileName
  anchor.style.display = 'none'
  document.body.appendChild(anchor)
  anchor.click()
  anchor.remove()
  window.setTimeout(() => URL.revokeObjectURL(url), 1000)
}
