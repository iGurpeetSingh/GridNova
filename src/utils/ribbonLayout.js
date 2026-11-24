/**
 * Prepare the same ribbon commands for both GridNova layouts.
 *
 * Univer's "collapsed" layout places ribbon groups and their commands on one
 * compact horizontal row. Grid coordinates only apply to the "grid" layout,
 * so omit them in Single Line mode to keep the native compact placement and
 * responsive overflow behavior.
 *
 * Univer's "grid" layout uses a two-row command grid, so every item receives
 * an explicit row/column and a visible caption.
 */
export function withRibbonGridLayouts(items = [], ribbonType = 'grid') {
  if (ribbonType === 'collapsed') {
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
