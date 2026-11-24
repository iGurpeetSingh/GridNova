import test from 'node:test'
import assert from 'node:assert/strict'
import {
  HOME_COMPACT_SELECTOR_IDS,
  HOME_EDITING_COMMAND_IDS,
  HOME_FORMAT_PAINTER_MENU_ITEM_ID,
  HOME_HIDDEN_MENU_ITEM_IDS,
  HOME_MENU_ORDER_OVERRIDES,
  HOME_RIBBON_GROUP_CONFIG,
  HOME_RIBBON_ICON_OVERRIDES,
  HOME_VALUE_SELECTOR_IDS,
  HOME_VISIBLE_MENU_ITEM_IDS,
  buildHomeRibbonMenuConfig,
} from '../src/utils/homeRibbonConfig.js'

test('Home keeps the native Clipboard, Font, Alignment and Number groups', () => {
  assert.deepEqual(Object.keys(HOME_RIBBON_GROUP_CONFIG), [
    'ribbon.start',
    'ribbon.start.history',
    'ribbon.start.format',
    'ribbon.start.layout',
    'ribbon.start.number',
    'ribbon.start.others',
  ])
  assert.equal(HOME_RIBBON_GROUP_CONFIG['ribbon.start.history'].title, 'Clipboard')
  assert.equal(HOME_RIBBON_GROUP_CONFIG['ribbon.start.format'].title, 'Font')
  assert.equal(HOME_RIBBON_GROUP_CONFIG['ribbon.start.layout'].title, 'Alignment')
  assert.equal(HOME_RIBBON_GROUP_CONFIG['ribbon.start.number'].title, 'Number')
  assert.equal(HOME_RIBBON_GROUP_CONFIG['ribbon.start.others'].hidden, true)
})

test('Font, size and number format keep their native selector values', () => {
  assert.deepEqual(HOME_VALUE_SELECTOR_IDS, [
    'sheet.command.set-range-font-family',
    'sheet.command.set-range-fontsize',
    'sheet.operation.open.numfmt.panel',
  ])
})

test('Alignment shows only the reference row controls', () => {
  assert.deepEqual(HOME_VISIBLE_MENU_ITEM_IDS, [
    'sheet.command.set-horizontal-text-align',
    'sheet.command.set-text-wrap',
    'sheet.command.add-worksheet-merge',
  ])
  const config = buildHomeRibbonMenuConfig()
  HOME_VISIBLE_MENU_ITEM_IDS.forEach((menuId) => assert.equal(config[menuId].hidden, false))
  assert.deepEqual(HOME_COMPACT_SELECTOR_IDS, HOME_VISIBLE_MENU_ITEM_IDS)
})

test('Unneeded Home controls are hidden without hiding spreadsheet context-menu Cut and Copy', () => {
  const config = buildHomeRibbonMenuConfig()
  HOME_HIDDEN_MENU_ITEM_IDS.forEach((menuId) => assert.equal(config[menuId].hidden, true))
  assert.ok(HOME_HIDDEN_MENU_ITEM_IDS.includes('ui.command.clear-formatting'))
  assert.equal(Object.hasOwn(config, 'sheet.command.cut'), false)
  assert.equal(Object.hasOwn(config, 'sheet.command.copy'), false)
})

test('Home native toolbar icons match the supplied Single Line reference', () => {
  const expected = [
    HOME_FORMAT_PAINTER_MENU_ITEM_ID,
    'sheet.menu.paste',
    'sheet.command.set-range-bold',
    'sheet.command.set-range-italic',
    'sheet.command.set-range-underline',
    'sheet.command.set-range-stroke',
    'sheet.command.set-range-text-color',
    'sheet.command.set-background-color',
    'sheet.command.set-border-basic',
    'sheet.command.set-horizontal-text-align',
    'sheet.command.set-text-wrap',
    'sheet.command.add-worksheet-merge',
    'sheet.command.numfmt.set.currency',
    'sheet.command.numfmt.set.percent',
    'sheet.command.numfmt.add.decimal.command',
    'sheet.command.numfmt.subtract.decimal.command',
    'sheet.command.sort-range-asc',
    'sheet.command.smart-toggle-filter',
    'ui.operation.open-find-dialog',
  ]
  expected.forEach((menuId) => assert.ok(HOME_RIBBON_ICON_OVERRIDES[menuId], menuId))
  assert.equal(buildHomeRibbonMenuConfig()['sheet.menu.paste'].icon, 'gridnova-icon-paste')
})

test('Home command ordering follows the supplied one-line toolbar', () => {
  assert.equal(HOME_MENU_ORDER_OVERRIDES['univer.command.undo'], 0)
  assert.equal(HOME_MENU_ORDER_OVERRIDES['univer.command.redo'], 1)
  assert.equal(HOME_MENU_ORDER_OVERRIDES['sheet.menu.paste'], 2)
  assert.equal(HOME_MENU_ORDER_OVERRIDES[HOME_FORMAT_PAINTER_MENU_ITEM_ID], 3)
  assert.equal(HOME_MENU_ORDER_OVERRIDES['sheet.command.set-range-font-family'], 0)
  assert.equal(HOME_MENU_ORDER_OVERRIDES['sheet.command.set-range-fontsize'], 1)
  assert.equal(HOME_MENU_ORDER_OVERRIDES['sheet.command.set-horizontal-text-align'], 0)
  assert.equal(HOME_MENU_ORDER_OVERRIDES['sheet.command.add-worksheet-merge'], 2)
  assert.equal(HOME_MENU_ORDER_OVERRIDES['sheet.command.numfmt.set.currency'], 1)
})

test('Home Editing commands are limited to the requested compact actions', () => {
  assert.deepEqual(HOME_EDITING_COMMAND_IDS, [
    'sheet.command.sort-range-asc',
    'sheet.command.smart-toggle-filter',
    'ui.operation.open-find-dialog',
  ])
})

test('Format Painter retains its native Univer menu ID', () => {
  assert.equal(HOME_FORMAT_PAINTER_MENU_ITEM_ID, 'ui.operation.activate-format-painter')
  assert.equal(buildHomeRibbonMenuConfig()[HOME_FORMAT_PAINTER_MENU_ITEM_ID].icon, 'gridnova-icon-paint')
})

test('Home removes the old Styles and Sheet custom groups and keeps only Cells and Editing', async () => {
  const { readFile } = await import('node:fs/promises')
  const source = await readFile(new URL('../src/components/UniverEditor.jsx', import.meta.url), 'utf8')
  const homeStart = source.indexOf('const homeGroups = [')
  const homeEnd = source.indexOf('const nativeTabGroups = [', homeStart)
  assert.ok(homeStart >= 0 && homeEnd > homeStart)
  const homeGroups = source.slice(homeStart, homeEnd)
  assert.match(homeGroups, /id: 'cells'/)
  assert.match(homeGroups, /id: 'editing'/)
  assert.doesNotMatch(homeGroups, /id: 'styles'/)
  assert.doesNotMatch(homeGroups, /id: 'sheet'/)
  assert.doesNotMatch(homeGroups, /cell-style-|conditional-formatting|row-height|column-width|fill-down|fill-right/)
  assert.match(homeGroups, /icon: 'eraser'/)
  assert.match(homeGroups, /icon: 'calculator'/)
})

test('Home no longer appends the old duplicate Cut and Copy toolbar buttons', async () => {
  const { readFile } = await import('node:fs/promises')
  const source = await readFile(new URL('../src/components/UniverEditor.jsx', import.meta.url), 'utf8')
  assert.doesNotMatch(source, /const historyMenus = \[/)
  assert.doesNotMatch(source, /\.\.\.historyMenus/)
  assert.match(source, /return \[\.\.\.icons, \.\.\.drawingComponents, \.\.\.menus, \.\.\.homeGroupMenus, \.\.\.nativeTabGroupMenus\]/)
})
