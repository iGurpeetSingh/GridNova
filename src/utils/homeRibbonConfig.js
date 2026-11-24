// GridNova removes the entire native Home ribbon: both its tab and its
// toolbar commands are hidden. Other workbook tabs are configured separately.
export const HOME_RIBBON_GROUP_CONFIG = Object.freeze({
  'ribbon.start': { title: 'Home', order: 1, hidden: true },
})
