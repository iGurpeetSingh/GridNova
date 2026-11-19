import test from 'node:test'
import assert from 'node:assert/strict'
import {
  isRibbonType,
  normalizeRibbonType,
  RIBBON_MODE_STORAGE_PREFIX,
  RIBBON_TYPE,
  getRibbonPresentationMode,
} from '../src/utils/ribbonMode.js'

test('Single Line keeps the tab strip visible with Univer classic toolbar', () => {
  assert.equal(RIBBON_TYPE.SINGLE_LINE, 'classic')
  assert.equal(normalizeRibbonType(null), RIBBON_TYPE.SINGLE_LINE)
  assert.equal(normalizeRibbonType('classic'), RIBBON_TYPE.SINGLE_LINE)
  assert.equal(normalizeRibbonType('collapsed'), RIBBON_TYPE.SINGLE_LINE)
  assert.equal(normalizeRibbonType('unknown'), RIBBON_TYPE.SINGLE_LINE)
})

test('Classic uses Univer grid two-row expanded ribbon', () => {
  assert.equal(RIBBON_TYPE.CLASSIC, 'grid')
  assert.equal(normalizeRibbonType('grid'), RIBBON_TYPE.CLASSIC)
})

test('Only the two supported GridNova ribbon modes can be selected', () => {
  assert.equal(isRibbonType(RIBBON_TYPE.SINGLE_LINE), true)
  assert.equal(isRibbonType(RIBBON_TYPE.CLASSIC), true)
  assert.equal(isRibbonType('collapsed'), false)
  assert.equal(isRibbonType('simple'), false)
})

test('Ribbon preference key is workbook-scoped and versioned', () => {
  assert.match(RIBBON_MODE_STORAGE_PREFIX, /^gridnova:ribbon-mode-v7:$/)
})


import { withRibbonGridLayouts } from '../src/utils/ribbonLayout.js'

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

test('Classic two-row mode keeps visible captions and places items in two rows', () => {
  const items = withRibbonGridLayouts(formulaCommands, RIBBON_TYPE.CLASSIC)
  assert.equal(items.length, formulaCommands.length)
  assert.deepEqual(items.map((item) => [item.gridLayout.row, item.gridLayout.column]), [
    [1, 1], [2, 1], [1, 2], [2, 2],
  ])
  items.forEach((item) => assert.equal(item.gridLayout.showLabel, true))
})

test('Single Line mode preserves commands and callbacks used across ribbon tabs', () => {
  const commands = [
    { id: 'paste-values', command: 'sheet.command.paste-value' },
    { id: 'add-text-box', run: () => true },
    { id: 'insert-chart', run: () => true },
    { id: 'sort-ascending', command: 'sheet.command.sort-range-asc' },
    { id: 'toggle-formula-bar', action: 'toggle-formula-bar' },
  ]
  const items = withRibbonGridLayouts(commands, RIBBON_TYPE.SINGLE_LINE)
  assert.deepEqual(items.map(({ id, command, action }) => ({ id, command, action })), [
    { id: 'paste-values', command: 'sheet.command.paste-value', action: undefined },
    { id: 'add-text-box', command: undefined, action: undefined },
    { id: 'insert-chart', command: undefined, action: undefined },
    { id: 'sort-ascending', command: 'sheet.command.sort-range-asc', action: undefined },
    { id: 'toggle-formula-bar', command: undefined, action: 'toggle-formula-bar' },
  ])
})


test('Single Line and expanded Classic apply their own CSS presentation state', () => {
  assert.equal(getRibbonPresentationMode(RIBBON_TYPE.SINGLE_LINE), 'single-line')
  assert.equal(getRibbonPresentationMode(RIBBON_TYPE.CLASSIC), 'classic')
})

test('The Single Line layout uses Univer classic to keep tabs above a separate toolbar row', () => {
  assert.equal(RIBBON_TYPE.SINGLE_LINE, 'classic')
  assert.equal(RIBBON_TYPE.CLASSIC, 'grid')
})
