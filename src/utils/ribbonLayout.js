/**
 * Prepare the same working ribbon commands for both GridNova layouts.
 *
 * Univer's "classic" layout renders a single horizontal toolbar and calculates
 * overflow from real button widths. Grid coordinates must therefore be omitted.
 * Univer's "grid" layout uses a two-row command grid with visible captions.
 *
 * Explicitly unsupported commands are filtered at this boundary. A control
 * advertised as "(Unavailable)" must not appear as a clickable ribbon button.
 */
export function withRibbonGridLayouts(items = [], ribbonType = 'grid') {
  const availableItems = items.filter((item) => {
    const explicitlyUnavailable = typeof item.action === 'string'
      && item.action.endsWith('-unavailable')
    const visiblyUnavailable = typeof item.title === 'string'
      && /\(unavailable\)/i.test(item.title)
    return !explicitlyUnavailable && !visiblyUnavailable
  })

  if (ribbonType === 'classic') {
    return availableItems.map(({ gridLayout, ...item }) => item)
  }

  return availableItems.map((item, index) => {
    const configured = item.gridLayout || {}
    return {
      ...item,
      gridLayout: {
        ...configured,
        row: (index % 2) + 1,
        column: Math.floor(index / 2) + 1,
        showLabel: configured.showLabel ?? true,
      },
    }
  })
}
