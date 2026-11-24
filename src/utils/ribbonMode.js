export const RIBBON_MODE_STORAGE_PREFIX = 'gridnova:ribbon-mode-v8:'

/**
 * GridNova Single Line uses Univer's collapsed layout: ribbon groups and their
 * commands share one compact horizontal row below the tab strip.
 * GridNova's expanded Classic mode uses Univer's grid layout with two command
 * rows and visible captions.
 */
export const RIBBON_TYPE = Object.freeze({
  SINGLE_LINE: 'collapsed',
  CLASSIC: 'grid',
})

export function normalizeRibbonType(savedValue) {
  // "classic" was previously (incorrectly) stored as GridNova's Single Line
  // mode. Keep old workbook preferences compatible while correcting the
  // underlying Univer mode to "collapsed".
  if (savedValue === RIBBON_TYPE.CLASSIC || savedValue === 'grid') return RIBBON_TYPE.CLASSIC
  return RIBBON_TYPE.SINGLE_LINE
}

/**
 * CSS presentation class for the selected Univer ribbon type.
 */
export function getRibbonPresentationMode(value) {
  return value === RIBBON_TYPE.SINGLE_LINE ? 'single-line' : 'classic'
}

export function isRibbonType(value) {
  return value === RIBBON_TYPE.SINGLE_LINE || value === RIBBON_TYPE.CLASSIC
}
