export const RIBBON_MODE_STORAGE_PREFIX = 'gridnova:ribbon-mode-v9:'

/**
 * In the GridNova / Univer 1.5.2 build, Single Line keeps the full ribbon tab
 * strip visible and renders the selected tab's commands on one horizontal
 * toolbar row. Expanded Classic uses Univer's two-row grid layout.
 */
export const RIBBON_TYPE = Object.freeze({
  SINGLE_LINE: 'classic',
  CLASSIC: 'grid',
})

export function normalizeRibbonType(savedValue) {
  // Older GridNova releases persisted "collapsed" for the Single Line
  // preference. In Univer 1.5.2 that hides the normal tab strip, so migrate it
  // (and the older "classic" value) to the visible-tab Single Line layout.
  if (savedValue === RIBBON_TYPE.CLASSIC || savedValue === 'grid') return RIBBON_TYPE.CLASSIC
  return RIBBON_TYPE.SINGLE_LINE
}

/**
 * CSS presentation class for the selected Univer 1.5.2 ribbon type.
 */
export function getRibbonPresentationMode(value) {
  return value === RIBBON_TYPE.SINGLE_LINE ? 'single-line' : 'classic'
}

export function isRibbonType(value) {
  return value === RIBBON_TYPE.SINGLE_LINE || value === RIBBON_TYPE.CLASSIC
}
