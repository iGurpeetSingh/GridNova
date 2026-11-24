export const HOME_RIBBON_GROUP_CONFIG = Object.freeze({
  'ribbon.start': { title: 'Home', order: 1 },
  'ribbon.start.history': { title: 'Clipboard', hidden: false, order: 0 },
  'ribbon.start.format': { title: 'Font', hidden: false, order: 1 },
  'ribbon.start.layout': { title: 'Alignment', hidden: false, order: 2 },
  'ribbon.start.number': { title: 'Number', hidden: false, order: 3 },
  'ribbon.start.others': { title: 'Tools', hidden: true, order: 4 },
})

export const HOME_FORMAT_PAINTER_MENU_ITEM_ID = 'ui.operation.activate-format-painter'

export const HOME_HIDDEN_MENU_ITEM_IDS = Object.freeze([
  'ui.command.clear-formatting',
  'sheet.command.set-vertical-text-align',
  'sheet.command.set-shrink-to-fit',
  'sheet.command.set-text-rotation',
  'sheet.command.add-range-protection-from-toolbar',
  'base-ui.operation.toggle-shortcut-panel',
  'ui.operation.open-feature-search',
])

export const HOME_VISIBLE_MENU_ITEM_IDS = Object.freeze([
  'sheet.command.set-horizontal-text-align',
  'sheet.command.set-text-wrap',
  'sheet.command.add-worksheet-merge',
])

export const HOME_RIBBON_ICON_OVERRIDES = Object.freeze({
  [HOME_FORMAT_PAINTER_MENU_ITEM_ID]: ['paint', 'Format Painter'],
  'sheet.menu.paste': ['paste', 'Paste'],
  'sheet.command.set-range-bold': ['bold', 'Bold'],
  'sheet.command.set-range-italic': ['italic', 'Italic'],
  'sheet.command.set-range-underline': ['underline', 'Underline'],
  'sheet.command.set-range-stroke': ['strike', 'Strikethrough'],
  'sheet.command.set-range-text-color': ['textColor', 'Font color'],
  'sheet.command.set-background-color': ['fillColor', 'Fill color'],
  'sheet.command.set-border-basic': ['border', 'Borders'],
  'sheet.command.set-horizontal-text-align': ['alignLeftCompact', 'Alignment'],
  'sheet.command.set-text-wrap': ['wrap', 'Wrap text'],
  'sheet.command.add-worksheet-merge': ['merge', 'Merge cells'],
  'sheet.command.numfmt.set.currency': ['currency', 'Currency'],
  'sheet.command.numfmt.set.percent': ['percent', 'Percentage'],
  'sheet.command.numfmt.add.decimal.command': ['decimalIncrease', 'Increase decimal places'],
  'sheet.command.numfmt.subtract.decimal.command': ['decimalDecrease', 'Decrease decimal places'],
  'sheet.command.sort-range-asc': ['sort', 'Sort A to Z'],
  'sheet.command.smart-toggle-filter': ['filter', 'Filter'],
  'ui.operation.open-find-dialog': ['search', 'Find'],
})

export const HOME_MENU_ORDER_OVERRIDES = Object.freeze({
  'univer.command.undo': 0,
  'univer.command.redo': 1,
  'sheet.menu.paste': 2,
  [HOME_FORMAT_PAINTER_MENU_ITEM_ID]: 3,
  'sheet.command.set-range-font-family': 0,
  'sheet.command.set-range-fontsize': 1,
  'sheet.command.set-range-bold': 2,
  'sheet.command.set-range-italic': 3,
  'sheet.command.set-range-underline': 4,
  'sheet.command.set-range-stroke': 5,
  'sheet.command.set-border-basic': 6,
  'sheet.command.set-background-color': 7,
  'sheet.command.set-range-text-color': 8,
  'sheet.command.set-horizontal-text-align': 0,
  'sheet.command.set-text-wrap': 1,
  'sheet.command.add-worksheet-merge': 2,
  'sheet.operation.open.numfmt.panel': 0,
  'sheet.command.numfmt.set.currency': 1,
  'sheet.command.numfmt.set.percent': 2,
  'sheet.command.numfmt.add.decimal.command': 3,
  'sheet.command.numfmt.subtract.decimal.command': 4,
})

export const HOME_VALUE_SELECTOR_IDS = Object.freeze([
  'sheet.command.set-range-font-family',
  'sheet.command.set-range-fontsize',
  'sheet.operation.open.numfmt.panel',
])

export const HOME_COMPACT_SELECTOR_IDS = Object.freeze([
  'sheet.command.set-horizontal-text-align',
  'sheet.command.set-text-wrap',
  'sheet.command.add-worksheet-merge',
])

export const HOME_EDITING_COMMAND_IDS = Object.freeze([
  'sheet.command.sort-range-asc',
  'sheet.command.smart-toggle-filter',
  'ui.operation.open-find-dialog',
])

export function buildHomeRibbonMenuConfig(iconPrefix = 'gridnova-icon-') {
  const hiddenItems = Object.fromEntries(
    HOME_HIDDEN_MENU_ITEM_IDS.map((menuId) => [menuId, { hidden: true }]),
  )
  const visibleItems = Object.fromEntries(
    HOME_VISIBLE_MENU_ITEM_IDS.map((menuId) => [menuId, { hidden: false }]),
  )
  const icons = Object.fromEntries(
    Object.entries(HOME_RIBBON_ICON_OVERRIDES).map(([commandId, [iconName, tooltip]]) => [
      commandId,
      { icon: iconPrefix + iconName, tooltip },
    ]),
  )
  const ordered = Object.fromEntries(
    Object.entries(HOME_MENU_ORDER_OVERRIDES).map(([menuId, order]) => [menuId, { order }]),
  )

  return {
    ...HOME_RIBBON_GROUP_CONFIG,
    ...hiddenItems,
    ...visibleItems,
    ...icons,
    ...ordered,
  }
}
