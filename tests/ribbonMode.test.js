import test from 'node:test'
import assert from 'node:assert/strict'
import {
  isRibbonType,
  normalizeRibbonType,
  RIBBON_MODE_STORAGE_PREFIX,
  RIBBON_TYPE,
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
