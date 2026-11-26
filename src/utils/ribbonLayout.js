/**
 * Prepare the same ribbon commands for both GridNova layouts.
 *
 * In the pinned Univer 1.5.2 package, "classic" keeps the complete tab strip
 * visible and renders the selected tab's commands as a single horizontal
 * toolbar row. Grid coordinates must be omitted in this mode so native
 * toolbar ordering and width-based overflow remain in control.
 *
 * "grid" is GridNova's expanded two-row Ribbon, so every item receives an
 * explicit row/column and a visible caption.
 */
export function withRibbonGridLayouts(items = [], ribbonType = 'grid') {
  if (ribbonType === 'classic') {
    return items.map((item) => {
      if (item.keepGridLayoutInClassic) {
        const { keepGridLayoutInClassic, ...preserved } = item
        return preserved
      }
      const { gridLayout, ...nativeItem } = item
      return nativeItem
    })
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
