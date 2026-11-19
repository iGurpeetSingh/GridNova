export const RIBBON_MODE_STORAGE_PREFIX = 'gridnova:ribbon-mode-v7:'

/**
 * Univer's "classic" type is GridNova's compact Single Line ribbon:
 * visible tabs plus one horizontal command strip with automatic More overflow.
 * Univer's "grid" type is GridNova's expanded Classic ribbon:
 * visible tabs plus the two-row grouped command ribbon.
 */
export const RIBBON_TYPE = Object.freeze({
  SINGLE_LINE: 'classic',
  CLASSIC: 'grid',
})

export function normalizeRibbonType(savedValue) {
  return savedValue === RIBBON_TYPE.CLASSIC
    ? RIBBON_TYPE.CLASSIC
    : RIBBON_TYPE.SINGLE_LINE
}

export function isRibbonType(value) {
  return value === RIBBON_TYPE.SINGLE_LINE || value === RIBBON_TYPE.CLASSIC
}
