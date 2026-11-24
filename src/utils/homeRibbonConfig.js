export const HOME_RIBBON_GROUP_CONFIG = Object.freeze({
  'ribbon.start': { title: 'Home', order: 1 },
  'ribbon.start.history': { title: 'Clipboard', hidden: false, order: 0 },
  'ribbon.start.format': { title: 'Font', hidden: false, order: 1 },
  'ribbon.start.layout': { title: 'Alignment', hidden: false, order: 2 },
  'ribbon.start.number': { title: 'Number', hidden: false, order: 3 },
  'ribbon.start.others': { title: 'Tools', hidden: false, order: 4 },
})

export const HOME_FORMAT_PAINTER_MENU_ITEM_ID = 'ui.operation.activate-format-painter'

export const HOME_RIBBON_ICON_OVERRIDES = Object.freeze({
  [HOME_FORMAT_PAINTER_MENU_ITEM_ID]: ['paint', 'Format Painter'],
  'sheet.command.smart-toggle-filter': ['filter', 'Filter'],
  'sheet.command.sort-range-asc': ['sort', 'Sort A to Z'],
  'sheet.command.sort-range-desc': ['sort', 'Sort Z to A'],
  'ui.operation.open-find-dialog': ['search', 'Find & Replace'],
})

export const HOME_VALUE_SELECTOR_IDS = Object.freeze([
  'sheet.command.set-range-font-family',
  'sheet.command.set-range-fontsize',
  'sheet.operation.open.numfmt.panel',
])

export const HOME_COMPACT_SELECTOR_IDS = Object.freeze([
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
])

export const HOME_EDITING_COMMAND_IDS = Object.freeze([
  'sheet.command.sort-range-asc',
  'sheet.command.sort-range-desc',
  'sheet.command.smart-toggle-filter',
  'ui.operation.open-find-dialog',
])

export function buildHomeRibbonMenuConfig(iconPrefix = 'gridnova-icon-') {
  return {
    ...HOME_RIBBON_GROUP_CONFIG,
    ...Object.fromEntries(
      Object.entries(HOME_RIBBON_ICON_OVERRIDES).map(([commandId, [iconName, tooltip]]) => [
        commandId,
        { icon: iconPrefix + iconName, tooltip },
      ]),
    ),
  }
}
