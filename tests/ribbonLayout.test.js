import test from 'node:test'
import assert from 'node:assert/strict'
import { withRibbonGridLayouts } from '../src/utils/ribbonLayout.js'

test('Ribbon items use unique positions across the two rows', () => {
  const result = withRibbonGridLayouts([
    { id: 'one' },
    { id: 'two' },
    { id: 'three' },
    { id: 'four' },
    { id: 'five' },
  ])

  assert.deepEqual(
    result.map(({ gridLayout }) => [gridLayout.row, gridLayout.column]),
    [[1, 1], [2, 1], [1, 2], [2, 2], [1, 3]],
  )
})

test('Custom Ribbon command labels are shown by default', () => {
  const result = withRibbonGridLayouts([{ id: 'save' }, { id: 'copy' }])
  assert.ok(result.every((item) => item.gridLayout.showLabel === true))
})

test('Existing grid width and icon sizing are retained', () => {
  const [result] = withRibbonGridLayouts([
    { id: 'font-family', gridLayout: { width: 150, iconSize: 18, row: 2, column: 9 } },
  ])

  assert.equal(result.gridLayout.width, 150)
  assert.equal(result.gridLayout.iconSize, 18)
  assert.deepEqual([result.gridLayout.row, result.gridLayout.column], [1, 1])
})


test('Single Line mode preserves command identity while removing captions through menu config', () => {
  const commands = [
    'sheet.command.set-range-bold',
    'sheet.command.set-range-italic',
    'sheet.command.set-text-wrap',
    'sheet.command.add-worksheet-merge',
    'sheet.command.smart-toggle-filter',
    'sheet.command.insert-row-before',
  ]
  const items = withRibbonGridLayouts(commands.map((id) => ({ id, command: id })), 'collapsed')
  assert.deepEqual(items.map((item) => item.command), commands)
  assert.ok(items.every((item) => !item.gridLayout))
})

