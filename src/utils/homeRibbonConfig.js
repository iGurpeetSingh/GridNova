export const HOME_RIBBON_ICON_OVERRIDES = Object.freeze({
  'sheet.command.smart-toggle-filter': ['filter', 'Toggle Filter'],
  'sheet.command.sort-range-asc': ['sort', 'Sort A to Z'],
  'sheet.command.sort-range-desc': ['sort', 'Sort Z to A'],
  'ui.operation.open-find-dialog': ['search', 'Find'],
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
  'sheet.command.set-range-text-color',
  'sheet.command.set-background-color',
  'sheet.command.set-border-basic',
])

export function buildHomeRibbonMenuConfig(iconPrefix = 'gridnova-icon-') {
  return Object.fromEntries(
    Object.entries(HOME_RIBBON_ICON_OVERRIDES).map(([commandId, [iconName, tooltip]]) => [
      commandId,
      { icon: iconPrefix + iconName, tooltip },
    ]),
  )
}
