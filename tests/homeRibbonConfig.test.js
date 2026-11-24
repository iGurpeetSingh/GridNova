import test from 'node:test'
import assert from 'node:assert/strict'
import {
  HOME_COMPACT_SELECTOR_IDS,
  HOME_EDITING_COMMAND_IDS,
  HOME_FORMAT_PAINTER_MENU_ITEM_ID,
  HOME_RIBBON_GROUP_CONFIG,
  HOME_RIBBON_ICON_OVERRIDES,
  HOME_VALUE_SELECTOR_IDS,
  buildHomeRibbonMenuConfig,
} from '../src/utils/homeRibbonConfig.js'

test('Home Ribbon uses the installed Univer History/Clipboard, Font, Alignment, Number and Tools groups', () => {
  assert.deepEqual(Object.keys(HOME_RIBBON_GROUP_CONFIG), [
    'ribbon.start',
    'ribbon.start.history',
    'ribbon.start.format',
    'ribbon.start.layout',
    'ribbon.start.number',
    'ribbon.start.others',
  ])
  assert.deepEqual(HOME_RIBBON_GROUP_CONFIG['ribbon.start.history'], {
    title: 'Clipboard', hidden: false, order: 0,
  })
  assert.deepEqual(HOME_RIBBON_GROUP_CONFIG['ribbon.start.format'], {
    title: 'Font', hidden: false, order: 1,
  })
  assert.deepEqual(HOME_RIBBON_GROUP_CONFIG['ribbon.start.layout'], {
    title: 'Alignment', hidden: false, order: 2,
  })
  assert.deepEqual(HOME_RIBBON_GROUP_CONFIG['ribbon.start.number'], {
    title: 'Number', hidden: false, order: 3,
  })
  assert.deepEqual(HOME_RIBBON_GROUP_CONFIG['ribbon.start.others'], {
    title: 'Tools', hidden: false, order: 4,
  })
  assert.equal(Object.hasOwn(HOME_RIBBON_GROUP_CONFIG, 'ribbon.start.clipboard'), false)
})

test('Font family, font size and number format retain visible selected values', () => {
  assert.deepEqual(HOME_VALUE_SELECTOR_IDS, [
    'sheet.command.set-range-font-family',
    'sheet.command.set-range-fontsize',
    'sheet.operation.open.numfmt.panel',
  ])
  const config = buildHomeRibbonMenuConfig()
  HOME_VALUE_SELECTOR_IDS.forEach((commandId) => assert.equal(Object.hasOwn(config, commandId), false))
})

test('Alignment, wrapping, rotation, merge, text colour, fill and borders use native Univer 1.5.2 menu IDs', () => {
  const expected = [
    'sheet.command.set-horizontal-text-align',
    'sheet.command.set-vertical-text-align',
    'sheet.command.set-text-wrap',
    'sheet.command.set-shrink-to-fit',
    'sheet.command.set-text-rotation',
    'sheet.command.add-worksheet-merge',
    'sheet.command.add-worksheet-merge-all',
    'sheet.command.add-worksheet-merge-vertical',
    'sheet.command.add-worksheet-merge-horizontal',
    'sheet.command.remove-worksheet-merge',
    'sheet.command.set-range-text-color',
    'sheet.command.set-background-color',
    'sheet.command.set-border-basic',
  ]
  expected.forEach((commandId) => assert.ok(HOME_COMPACT_SELECTOR_IDS.includes(commandId), commandId))
})

test('Home Editing group exposes Excel actions without using duplicate placeholder controls', () => {
  assert.deepEqual(HOME_EDITING_COMMAND_IDS, [
    'sheet.command.sort-range-asc',
    'sheet.command.sort-range-desc',
    'sheet.command.smart-toggle-filter',
    'ui.operation.open-find-dialog',
  ])
  const config = buildHomeRibbonMenuConfig()
  assert.deepEqual(config['sheet.command.smart-toggle-filter'], {
    icon: 'gridnova-icon-filter',
    tooltip: 'Filter',
  })
  assert.equal(config['sheet.command.sort-range-asc'].tooltip, 'Sort A to Z')
  assert.equal(config['sheet.command.sort-range-desc'].tooltip, 'Sort Z to A')
  assert.equal(config['ui.operation.open-find-dialog'].tooltip, 'Find & Replace')
})

test('Format Painter points to the Univer desktop toolbar menu item ID', () => {
  assert.equal(HOME_FORMAT_PAINTER_MENU_ITEM_ID, 'ui.operation.activate-format-painter')
  assert.deepEqual(buildHomeRibbonMenuConfig()[HOME_FORMAT_PAINTER_MENU_ITEM_ID], {
    icon: 'gridnova-icon-paint',
    tooltip: 'Format Painter',
  })
})


test('UniverEditor imports the shared Home Ribbon group configuration before using it', async () => {
  const { readFile } = await import('node:fs/promises')
  const source = await readFile(new URL('../src/components/UniverEditor.jsx', import.meta.url), 'utf8')
  assert.match(source, /import\s*\{[^}]*HOME_RIBBON_GROUP_CONFIG[^}]*\}\s*from\s*['"]\.\.\/utils\/homeRibbonConfig\.js['"]/s)
  assert.match(source, /menu:\s*\{\s*\.\.\.HOME_RIBBON_GROUP_CONFIG/s)
  assert.match(source, /\.\.\.buildHomeRibbonMenuConfig\(\)/)
})
