export const RIBBON_MODE_STORAGE_PREFIX = 'gridnova:ribbon-mode-v7:'

/**
 * Univer ribbon modes used by GridNova:
 * - collapsed: compact Single Line ribbon with the active tab's groups/commands
 *   on a single horizontal toolbar row and overflow handled by the engine.
 * - grid: expanded Classic ribbon with a tab strip and two-row command groups.
 *
 * The old GridNova preference used "classic" for Single Line. Treat it as a
 * legacy saved value so existing users keep their selected mode after upgrade.
 */
export const RIBBON_TYPE = Object.freeze({
  SINGLE_LINE: 'collapsed',
  CLASSIC: 'grid',
})

export function normalizeRibbonType(savedValue) {
  if (savedValue === RIBBON_TYPE.CLASSIC) return RIBBON_TYPE.CLASSIC
  return RIBBON_TYPE.SINGLE_LINE
}

export function isRibbonType(value) {
  return value === RIBBON_TYPE.SINGLE_LINE || value === RIBBON_TYPE.CLASSIC
}
