import test from 'node:test'
import assert from 'node:assert/strict'
import {
  buildCsvContent,
  createPdfAutoTableOptions,
  groupPdfTextItems,
  parseCsvText,
} from '../src/services/dataFileService.js'

test('CSV parser auto-detects semicolon delimiters and preserves quoted delimiters', () => {
  const result = parseCsvText('Name;Age\r\n"Ada; Lovelace";36')
  assert.equal(result.delimiter, ';')
  assert.deepEqual(result.rows, [
    ['Name', 'Age'],
    ['Ada; Lovelace', '36'],
  ])
})

test('CSV parser handles quoted commas, embedded newlines and escaped quotes', () => {
  const result = parseCsvText('Name,Note\r\nAda,"said ""hello, world""\nand left"')
  assert.deepEqual(result.rows, [
    ['Name', 'Note'],
    ['Ada', 'said "hello, world"\nand left'],
  ])
})

test('CSV parser recognizes tab-delimited content', () => {
  const result = parseCsvText('A\tB\n1\t2')
  assert.equal(result.delimiter, '\t')
  assert.deepEqual(result.rows, [['A', 'B'], ['1', '2']])
})

test('CSV parser accepts a valid single-column file', () => {
  const result = parseCsvText('one value')
  assert.deepEqual(result.rows, [['one value']])
})

test('CSV parser rejects empty input', () => {
  assert.throws(() => parseCsvText(''), /empty|readable rows/i)
})

test('CSV export escapes commas, quotes, newlines and formula injection prefixes', () => {
  const csv = buildCsvContent([
    ['Name', 'Details', 'Formula-like'],
    ['A, B', 'Said "hi"\nand left', '=1+1'],
  ])
  assert.match(csv, /"A, B"/)
  assert.match(csv, /"Said ""hi""\nand left"/)
  assert.match(csv, /'=1\+1/)
})

test('PDF text items are grouped into Y rows and X-separated table cells', () => {
  const items = [
    { str: 'Amount', width: 35, transform: [1, 0, 0, 10, 120, 100] },
    { str: 'Item', width: 18, transform: [1, 0, 0, 10, 10, 100] },
    { str: '120', width: 16, transform: [1, 0, 0, 10, 120, 80] },
    { str: 'Pencils', width: 32, transform: [1, 0, 0, 10, 10, 80] },
  ]
  assert.deepEqual(groupPdfTextItems(items), [
    ['Item', 'Amount'],
    ['Pencils', '120'],
  ])
})

test('PDF AutoTable options map Univer bold, fill, text color and alignment styles', () => {
  const options = createPdfAutoTableOptions({
    values: [['Styled']],
    cellData: { 0: { 0: { s: 'heading' } } },
    styles: {
      heading: { bl: 1, bg: { rgb: '#D9EAD3' }, cl: { rgb: '#274E13' }, ht: 'center' },
    },
    sheetName: 'Report',
  })
  assert.deepEqual(options.body, [['Styled']])
  const hook = { section: 'body', row: { index: 0 }, column: { index: 0 }, cell: { styles: {} } }
  options.didParseCell(hook)
  assert.equal(hook.cell.styles.fontStyle, 'bold')
  assert.equal(hook.cell.styles.fillColor, '#D9EAD3')
  assert.equal(hook.cell.styles.textColor, '#274E13')
  assert.equal(hook.cell.styles.halign, 'center')
})


test('CSV export keeps numeric negative values numeric while escaping untrusted strings', () => {
  assert.equal(buildCsvContent([['Text', 'Number'], ['-not a number', -125]]), "'-not a number,-125")
})
