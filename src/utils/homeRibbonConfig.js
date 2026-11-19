export const HOME_RIBBON_ICON_OVERRIDES = Object.freeze({
  'univer.command.undo': ['undo', 'Undo'],
  'univer.command.redo': ['redo', 'Redo'],
  'sheet.command.paste': ['paste', 'Paste'],
  'sheet.command.cut': ['cut', 'Cut'],
  'sheet.command.copy': ['copy', 'Copy'],
  'sheet.command.set-once-format-painter': ['paint', 'Format Painter'],
  'sheet.command.set-range-font-increase': ['fontIncrease', 'Increase Font Size'],
  'sheet.command.set-range-font-decrease': ['fontDecrease', 'Decrease Font Size'],
  'sheet.command.set-range-bold': ['bold', 'Bold'],
  'sheet.command.set-range-italic': ['italic', 'Italic'],
  'sheet.command.set-range-underline': ['underline', 'Underline'],
  'sheet.command.set-range-stroke': ['strike', 'Strikethrough'],
  'sheet.command.set-range-text-color': ['textColor', 'Font Color'],
  'sheet.command.set-background-color': ['fillColor', 'Fill Color'],
  'sheet.command.set-border-basic': ['border', 'Borders'],
  'sheet.command.set-horizontal-text-align': ['alignLeft', 'Horizontal Alignment'],
  'sheet.command.set-vertical-text-align': ['alignCenter', 'Vertical Alignment'],
  'sheet.command.set-text-wrap': ['wrap', 'Wrap Text'],
  'sheet.command.set-shrink-to-fit': ['shrink', 'Shrink to Fit'],
  'sheet.command.set-text-rotation': ['rotate', 'Text Rotation'],
  'sheet.command.add-worksheet-merge': ['merge', 'Merge Cells'],
  'sheet.command.add-worksheet-merge-all': ['merge', 'Merge All'],
  'sheet.command.add-worksheet-merge-vertical': ['merge', 'Merge Vertically'],
  'sheet.command.add-worksheet-merge-horizontal': ['merge', 'Merge Horizontally'],
  'sheet.command.remove-worksheet-merge': ['merge', 'Unmerge Cells'],
  'sheet.command.numfmt.set.currency': ['currency', 'Currency Format'],
  'sheet.command.numfmt.set.percent': ['percent', 'Percentage Format'],
  'sheet.command.numfmt.add.decimal.command': ['decimalIncrease', 'Increase Decimal Places'],
  'sheet.command.numfmt.subtract.decimal.command': ['decimalDecrease', 'Decrease Decimal Places'],
  'sheet.command.smart-toggle-filter': ['filter', 'Toggle Filter'],
  'sheet.command.sort-range-asc': ['sort', 'Sort A to Z'],
  'sheet.command.sort-range-desc': ['sort', 'Sort Z to A'],
  'ui.operation.open-find-dialog': ['search', 'Find'],
  'sheet.command.insert-row-before': ['insertRow', 'Insert Row'],
  'sheet.command.insert-col-before': ['insertColumn', 'Insert Column'],
  'sheet.command.remove-row-confirm': ['deleteRow', 'Delete Row'],
  'sheet.command.remove-col-confirm': ['deleteColumn', 'Delete Column'],
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
