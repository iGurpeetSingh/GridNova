import test from 'node:test'
import assert from 'node:assert/strict'
import {
  isRibbonType,
  normalizeRibbonType,
  RIBBON_MODE_STORAGE_PREFIX,
  RIBBON_TYPE,
  getRibbonPresentationMode,
} from '../src/utils/ribbonMode.js'
import { withRibbonGridLayouts } from '../src/utils/ribbonLayout.js'

test('Single Line keeps the full tabs visible with a one-row Univer 1.5.2 toolbar', () => {
  assert.equal(RIBBON_TYPE.SINGLE_LINE, 'classic')
  assert.equal(normalizeRibbonType(null), RIBBON_TYPE.SINGLE_LINE)
  assert.equal(normalizeRibbonType('classic'), RIBBON_TYPE.SINGLE_LINE)
  assert.equal(normalizeRibbonType('collapsed'), RIBBON_TYPE.SINGLE_LINE)
  assert.equal(normalizeRibbonType('unknown'), RIBBON_TYPE.SINGLE_LINE)
})

test('Expanded Classic uses Univer grid mode with two command rows', () => {
  assert.equal(RIBBON_TYPE.CLASSIC, 'grid')
  assert.equal(normalizeRibbonType('grid'), RIBBON_TYPE.CLASSIC)
})

test('Only the two supported GridNova ribbon modes can be selected', () => {
  assert.equal(isRibbonType(RIBBON_TYPE.SINGLE_LINE), true)
  assert.equal(isRibbonType(RIBBON_TYPE.CLASSIC), true)
  assert.equal(isRibbonType('collapsed'), false)
  assert.equal(isRibbonType('simple'), false)
})

test('Ribbon preference key is workbook-scoped and versioned for collapsed-mode migration', () => {
  assert.match(RIBBON_MODE_STORAGE_PREFIX, /^gridnova:ribbon-mode-v9:$/)
})

const formulaCommands = [
  { id: 'insert-function', title: 'Insert Function', command: 'formula-ui.operation.insert-function' },
  { id: 'auto-sum', title: 'AutoSum', run: () => true },
  { id: 'function-if', title: 'IF', run: () => true },
  { id: 'function-vlookup', title: 'VLOOKUP', run: () => true },
]

test('Single Line mode keeps every Formula command without two-row grid coordinates', () => {
  const items = withRibbonGridLayouts(formulaCommands, RIBBON_TYPE.SINGLE_LINE)
  assert.equal(items.length, formulaCommands.length)
  items.forEach((item, index) => {
    assert.equal(item.id, formulaCommands[index].id)
    assert.equal(item.title, formulaCommands[index].title)
    assert.equal(item.command, formulaCommands[index].command)
    assert.equal(item.run, formulaCommands[index].run)
    assert.equal(Object.hasOwn(item, 'gridLayout'), false)
  })
})

test('Expanded Classic two-row mode keeps captions and places items in two rows', () => {
  const items = withRibbonGridLayouts(formulaCommands, RIBBON_TYPE.CLASSIC)
  assert.deepEqual(items.map((item) => [item.gridLayout.row, item.gridLayout.column]), [
    [1, 1], [2, 1], [1, 2], [2, 2],
  ])
  items.forEach((item) => assert.equal(item.gridLayout.showLabel, true))
})

test('Single Line and expanded Classic apply their own CSS presentation state', () => {
  assert.equal(getRibbonPresentationMode(RIBBON_TYPE.SINGLE_LINE), 'single-line')
  assert.equal(getRibbonPresentationMode(RIBBON_TYPE.CLASSIC), 'classic')
})
