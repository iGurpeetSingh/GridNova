import Papa from 'papaparse'
import { getDocument, GlobalWorkerOptions } from 'pdfjs-dist'
import pdfWorkerUrl from 'pdfjs-dist/build/pdf.worker.min.mjs?url'
import { jsPDF } from 'jspdf'
import { autoTable } from 'jspdf-autotable'

GlobalWorkerOptions.workerSrc = pdfWorkerUrl

const CSV_DELIMITERS = [',', ';', '\t']

function normalizeRows(rows) {
  const output = (Array.isArray(rows) ? rows : []).map((row) =>
    (Array.isArray(row) ? row : [row]).map((value) => value == null ? '' : String(value)),
  )
  while (output.length && output.at(-1).every((value) => value === '')) output.pop()
  return output
}

function validateCsvResult(result) {
  const fatal = (result.errors || []).filter((error) =>
    error.type === 'Quotes' || error.code === 'UndetectableDelimiter',
  )
  if (fatal.length) {
    throw new Error('The CSV file contains invalid quoting or an unreadable delimiter: ' + fatal[0].message)
  }
  const rows = normalizeRows(result.data)
  if (!rows.length || !rows.some((row) => row.some((value) => value.trim() !== ''))) {
    throw new Error('The CSV file is empty or does not contain readable rows.')
  }
  return {
    rows,
    delimiter: result.meta?.delimiter || ',',
    warnings: (result.errors || []).map((error) => error.message).filter(Boolean),
  }
}

export function parseCsvText(text) {
  const result = Papa.parse(String(text ?? '').replace(/^\uFEFF/, ''), {
    delimiter: '',
    delimitersToGuess: CSV_DELIMITERS,
    quoteChar: '"',
    escapeChar: '"',
    dynamicTyping: false,
    skipEmptyLines: false,
    newline: '',
  })
  return validateCsvResult(result)
}

export function parseCsvFile(file, options = {}) {
  if (!file) return Promise.reject(new Error('Choose a CSV file first.'))
  const name = String(file.name || '').toLowerCase()
  if (!name.endsWith('.csv') && !name.endsWith('.tsv')) {
    return Promise.reject(new Error('Unsupported file type. Choose a .csv or .tsv file.'))
  }

  return new Promise((resolve, reject) => {
    Papa.parse(file, {
      delimiter: '',
      delimitersToGuess: CSV_DELIMITERS,
      quoteChar: '"',
      escapeChar: '"',
      dynamicTyping: false,
      skipEmptyLines: false,
      worker: file.size > 1024 * 1024 && Papa.WORKERS_SUPPORTED,
      complete: (result) => {
        try {
          const parsed = validateCsvResult(result)
          options.onProgress?.(100)
          resolve(parsed)
        } catch (error) {
          reject(error)
        }
      },
      error: (error) => reject(new Error(error?.message || 'The CSV file could not be read.')),
    })
  })
}

function textItemMetrics(item) {
  const transform = Array.isArray(item.transform) ? item.transform : []
  const x = Number(transform[4]) || 0
  const y = Number(transform[5]) || 0
  const fontHeight = Math.max(5, Math.abs(Number(transform[3]) || Number(item.height) || 10))
  return { x, y, fontHeight, text: String(item.str || '').replace(/\u0000/g, '') }
}

export function groupPdfTextItems(items, yTolerance = 3) {
  const normalized = (items || [])
    .filter((item) => typeof item?.str === 'string' && item.str.trim() !== '')
    .map(textItemMetrics)
    .sort((a, b) => b.y - a.y || a.x - b.x)

  const lines = []
  normalized.forEach((item) => {
    let target = null
    for (let index = lines.length - 1; index >= Math.max(0, lines.length - 3); index -= 1) {
      const line = lines[index]
      const tolerance = Math.max(yTolerance, item.fontHeight * 0.38, line.fontHeight * 0.38)
      if (Math.abs(line.y - item.y) <= tolerance) {
        target = line
        break
      }
    }
    if (!target) {
      target = { y: item.y, fontHeight: item.fontHeight, items: [] }
      lines.push(target)
    }
    target.items.push(item)
    target.fontHeight = Math.max(target.fontHeight, item.fontHeight)
    target.y = target.items.reduce((sum, entry) => sum + entry.y, 0) / target.items.length
  })

  return lines
    .sort((a, b) => b.y - a.y)
    .map((line) => {
      const sorted = line.items.sort((a, b) => a.x - b.x)
      const gapThreshold = Math.max(9, line.fontHeight * 1.45)
      const cells = []
      let current = ''
      let previousRight = null

      sorted.forEach((item) => {
        const gap = previousRight == null ? 0 : item.x - previousRight
        if (current && gap > gapThreshold) {
          cells.push(current.trim())
          current = item.text
        } else {
          current += (current ? ' ' : '') + item.text
        }
        previousRight = item.x + Math.max(0, Number(item.width) || item.text.length * line.fontHeight * 0.45)
      })
      if (current || !cells.length) cells.push(current.trim())
      return cells
    })
}

export async function parsePdfFile(file, options = {}) {
  if (!file) throw new Error('Choose a PDF file first.')
  if (!String(file.name || '').toLowerCase().endsWith('.pdf') && file.type !== 'application/pdf') {
    throw new Error('Unsupported file type. Choose a .pdf file.')
  }

  let loadingTask
  let pdf
  try {
    const data = new Uint8Array(await file.arrayBuffer())
    loadingTask = getDocument({ data, isEvalSupported: false })
    pdf = await loadingTask.promise
    const allRows = []
    let extractedItems = 0

    for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
      options.onProgress?.(Math.round((pageNumber - 1) / pdf.numPages * 100))
      const page = await pdf.getPage(pageNumber)
      const content = await page.getTextContent()
      const items = (content.items || []).filter((item) => typeof item?.str === 'string' && item.str.trim() !== '')
      extractedItems += items.length
      const rows = groupPdfTextItems(items)
      if (rows.length) {
        if (allRows.length) allRows.push([])
        allRows.push(...rows)
      }
      page.cleanup?.()
    }

    if (!extractedItems || !allRows.some((row) => row.some((cell) => String(cell).trim()))) {
      throw new Error('This PDF has no selectable text. It may be a scanned image; run OCR on the PDF before importing it.')
    }

    options.onProgress?.(100)
    const maxColumns = Math.max(1, ...allRows.map((row) => row.length))
    const rows = allRows.map((row) => Array.from({ length: maxColumns }, (_, index) => row[index] ?? ''))
    return { rows, pages: pdf.numPages, extractedItems }
  } catch (error) {
    if (error instanceof Error && /no selectable text|scanned image/i.test(error.message)) throw error
    throw new Error(error instanceof Error ? 'Could not read this PDF: ' + error.message : 'The PDF appears to be corrupted or unsupported.')
  } finally {
    try { await loadingTask?.destroy?.() } catch { /* the PDF worker may already be closed */ }
  }
}

export function buildCsvContent(values) {
  const rows = normalizeRows(values)
  return Papa.unparse(rows, {
    delimiter: ',',
    newline: '\r\n',
    quotes: false,
    escapeFormulae: true,
  })
}

export function downloadCsvFile(values, fileName = 'GridNova Sheet.csv') {
  const csv = buildCsvContent(values)
  const blob = new Blob(['\uFEFF', csv], { type: 'text/csv;charset=utf-8' })
  downloadBlob(blob, fileName)
}

function resolveStyle(cell, styles = {}) {
  if (!cell?.s) return {}
  return typeof cell.s === 'string' ? styles[cell.s] || {} : cell.s
}

function normalizeColor(color) {
  const value = color?.rgb || color
  if (!value) return undefined
  const hex = String(value).replace(/^#/, '')
  if (/^[0-9a-f]{8}$/i.test(hex)) return '#' + hex.slice(2)
  if (/^[0-9a-f]{6}$/i.test(hex)) return '#' + hex
  return undefined
}

function normalizeHorizontalAlignment(value) {
  if (['left', 'center', 'right', 'justify'].includes(value)) return value === 'justify' ? 'left' : value
  if (value === 2 || value === '2') return 'center'
  if (value === 3 || value === '3') return 'right'
  return 'left'
}

export function createPdfAutoTableOptions({ values, cellData = {}, styles = {}, sheetName = 'Sheet1' }) {
  const rows = normalizeRows(values)
  return {
    body: rows.map((row) => row.map((value) => String(value ?? ''))),
    theme: 'grid',
    margin: { top: 48, right: 28, bottom: 28, left: 28 },
    styles: {
      font: 'helvetica',
      fontSize: 8,
      cellPadding: 3,
      overflow: 'linebreak',
      valign: 'middle',
      lineColor: [220, 226, 232],
      lineWidth: 0.35,
      textColor: [38, 50, 56],
    },
    didDrawPage: (data) => {
      if (data.pageNumber === 1) {
        data.doc.setFont('helvetica', 'bold')
        data.doc.setFontSize(12)
        data.doc.setTextColor(34, 51, 43)
        data.doc.text(String(sheetName || 'GridNova Sheet').slice(0, 120), data.settings.margin.left, 28)
      }
      data.doc.setFont('helvetica', 'normal')
      data.doc.setFontSize(7)
      data.doc.setTextColor(130)
      data.doc.text('GridNova · ' + data.pageNumber, data.settings.margin.left, data.doc.internal.pageSize.getHeight() - 12)
    },
    didParseCell: (hook) => {
      if (hook.section !== 'body') return
      const source = cellData?.[hook.row.index]?.[hook.column.index]
      const style = resolveStyle(source, styles)
      const fontStyle = style.bl && style.it ? 'bolditalic' : style.bl ? 'bold' : style.it ? 'italic' : 'normal'
      hook.cell.styles.fontStyle = fontStyle
      const background = normalizeColor(style.bg)
      const textColor = normalizeColor(style.cl)
      if (background) hook.cell.styles.fillColor = background
      if (textColor) hook.cell.styles.textColor = textColor
      hook.cell.styles.halign = normalizeHorizontalAlignment(style.ht)
      if (style.tb) hook.cell.styles.overflow = 'linebreak'
    },
  }
}

export function exportPdfFile({
  values,
  cellData = {},
  styles = {},
  sheetName = 'Sheet1',
  fileName = 'GridNova Sheet.pdf',
  orientation = 'portrait',
}) {
  if (!Array.isArray(values) || !values.length || !values.some((row) => row?.some((value) => value !== ''))) {
    throw new Error('The active worksheet has no data to export as PDF.')
  }
  const doc = new jsPDF({ orientation: orientation === 'landscape' ? 'landscape' : 'portrait', unit: 'pt', format: 'a4' })
  autoTable(doc, createPdfAutoTableOptions({ values, cellData, styles, sheetName }))
  doc.save(fileName)
}

function downloadBlob(blob, fileName) {
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
