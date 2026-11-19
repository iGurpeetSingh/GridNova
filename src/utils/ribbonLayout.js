/**
 * Place a custom Ribbon group's commands in Univer's two-row grid and make
 * their captions visible. Univer defaults grid labels to hidden.
 */
export function withRibbonGridLayouts(items = []) {
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
