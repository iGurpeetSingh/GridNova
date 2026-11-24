export const RIBBON_MODE_STORAGE_PREFIX = 'gridnova:ribbon-mode-v8:'

/**
 * GridNova's Single Line mode uses Univer's classic layout: the full tab strip
 * remains visible while the active tab's controls stay on one horizontal row.
 * GridNova's expanded Classic mode uses Univer's grid layout: each command
 * group gets two rows, visible command captions, group labels and More overflow.
 */
export const RIBBON_TYPE = Object.freeze({
  SINGLE_LINE: 'classic',
  CLASSIC: 'grid',
})

export function normalizeRibbonType(savedValue) {
  if (savedValue === RIBBON_TYPE.CLASSIC) return RIBBON_TYPE.CLASSIC
  return RIBBON_TYPE.SINGLE_LINE
}

/**
 * CSS presentation class for the selected Univer ribbon type.
 * Univer 1.5.2 uses "classic" for a tab row plus a separate toolbar row,
 * and "grid" for the expanded two-row command groups.
 */
export function getRibbonPresentationMode(value) {
  return value === RIBBON_TYPE.SINGLE_LINE ? 'single-line' : 'classic'
}

export function isRibbonType(value) {
  return value === RIBBON_TYPE.SINGLE_LINE || value === RIBBON_TYPE.CLASSIC
}
