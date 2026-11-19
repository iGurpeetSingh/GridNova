/**
 * Prepare the same ribbon commands for both GridNova layouts.
 *
 * Univer's "classic" layout renders a single horizontal toolbar and calculates
 * overflow from the real button widths. Grid coordinates must therefore be
 * omitted in this mode; retaining row 1/2 coordinates can hide or mis-order
 * commands in the one-line toolbar.
 *
 * Univer's "grid" layout uses a two-row command grid, so every item receives
 * an explicit row/column and a visible caption.
 */
export function withRibbonGridLayouts(items = [], ribbonType = 'grid') {
  if (ribbonType === 'classic') {
    return items.map(({ gridLayout, ...item }) => item)
  }

  return items.map((item, index) => {
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
