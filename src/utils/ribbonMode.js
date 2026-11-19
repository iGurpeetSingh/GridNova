export const RIBBON_MODE_STORAGE_PREFIX = 'gridnova:ribbon-mode-v7:'

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
  return savedValue === RIBBON_TYPE.CLASSIC
    ? RIBBON_TYPE.CLASSIC
    : RIBBON_TYPE.SINGLE_LINE
}

export function isRibbonType(value) {
  return value === RIBBON_TYPE.SINGLE_LINE || value === RIBBON_TYPE.CLASSIC
}
