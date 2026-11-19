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
