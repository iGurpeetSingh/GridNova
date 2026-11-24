import test from 'node:test'
import assert from 'node:assert/strict'
import {
  HOME_COMPACT_SELECTOR_IDS,
  HOME_FORMAT_PAINTER_MENU_ITEM_ID,
  HOME_RIBBON_GROUP_CONFIG,
  HOME_RIBBON_ICON_OVERRIDES,
  HOME_VALUE_SELECTOR_IDS,
  buildHomeRibbonMenuConfig,
} from '../src/utils/homeRibbonConfig.js'

test('Home tab keeps selected-value controls visible instead of converting them to icon-only buttons', () => {
  assert.deepEqual(HOME_VALUE_SELECTOR_IDS, [
    'sheet.command.set-range-font-family',
    'sheet.command.set-range-fontsize',
    'sheet.operation.open.numfmt.panel',
  ])
  const config = buildHomeRibbonMenuConfig()
  HOME_VALUE_SELECTOR_IDS.forEach((commandId) => {
    assert.equal(Object.hasOwn(config, commandId), false)
  })
})

test('Home tab includes compact selector controls for alignment, wrap, merge, color and borders', () => {
  for (const commandId of [
    'sheet.command.set-horizontal-text-align',
    'sheet.command.set-vertical-text-align',
    'sheet.command.set-text-wrap',
    'sheet.command.add-worksheet-merge',
    'sheet.command.set-range-text-color',
    'sheet.command.set-background-color',
    'sheet.command.set-border-basic',
  ]) {
    assert.ok(HOME_COMPACT_SELECTOR_IDS.includes(commandId), commandId)
  }
})

test('Home tab only overrides icons where commands do not have state-aware native icons', () => {
  const config = buildHomeRibbonMenuConfig()
  assert.deepEqual(config['sheet.command.smart-toggle-filter'], {
    icon: 'gridnova-icon-filter',
    tooltip: 'Toggle Filter',
  })
  assert.equal(config['sheet.command.sort-range-asc'].tooltip, 'Sort A to Z')
  assert.equal(config['sheet.command.sort-range-desc'].tooltip, 'Sort Z to A')
  assert.equal(config['ui.operation.open-find-dialog'].tooltip, 'Find')

  // Univer supplies the selected-state icons for its native font/alignment/color controls.
  for (const commandId of [
    'sheet.command.set-range-bold',
    'sheet.command.set-range-italic',
    'sheet.command.set-range-text-color',
    'sheet.command.set-background-color',
    'sheet.command.set-horizontal-text-align',
    'sheet.command.set-text-wrap',
  ]) {
    assert.equal(Object.hasOwn(HOME_RIBBON_ICON_OVERRIDES, commandId), false)
  }
})


test('Home Ribbon restores the native History and Number groups in Excel order', () => {
  assert.deepEqual(HOME_RIBBON_GROUP_CONFIG['ribbon.start.history'], { hidden: false, order: 0 })
  assert.deepEqual(HOME_RIBBON_GROUP_CONFIG['ribbon.start.clipboard'], { hidden: true })
  assert.deepEqual(HOME_RIBBON_GROUP_CONFIG['ribbon.start.format'], { title: 'Font', order: 2 })
  assert.deepEqual(HOME_RIBBON_GROUP_CONFIG['ribbon.start.layout'], { title: 'Alignment', order: 3 })
  assert.deepEqual(HOME_RIBBON_GROUP_CONFIG['ribbon.start.others'], {
    title: 'Number & Tools', hidden: false, order: 4,
  })
})

test('Format Painter uses the installed Univer toolbar menu item ID', () => {
  assert.equal(HOME_FORMAT_PAINTER_MENU_ITEM_ID, 'ui.operation.activate-format-painter')
})
