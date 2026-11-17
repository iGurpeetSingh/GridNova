# GridNova

GridNova is a desktop-first spreadsheet workspace built with React, Vite, and Univer Sheets Core.

## Current implementation

- Dashboard with blank workbook and starter templates
- Recent workbook list
- Workbook editor with the Univer Sheets Core interface
- Responsive grid-style ribbon and built-in context menus
- Google Sheets-style column workflow: new sheets start with columns A–Z, inserted columns continue as AA, AB, AC…, and column-header right-click includes Insert Before/After, Delete selected columns, Hide/Show, Resize, and Fit to data
- Formula calculation, cell formatting, sheet tabs, row/column operations, and core spreadsheet editing through Univer
- Browser IndexedDB draft saving and reopen flow
- XLSX and CSV/TSV import, XLSX and active-sheet CSV export
- Chart preview from a selected range
- Basic grouped Pivot Summary creation on a new worksheet
- Phone view-only entry page by product decision

## Run locally

```bash
npm install
npm run dev
```

## Production build

```bash
npm run build
npm run preview
```

## Current storage and file limitations

- Google Drive is not connected yet. Workbook drafts currently live in the same browser profile using IndexedDB.
- The IndexedDB storage is a temporary development adapter. It is not cloud sync and does not move drafts between devices.
- XLSX conversion currently targets ordinary worksheets, values, formulas, basic formatting, row/column dimensions, and merged cells. It is not a lossless round-trip for every Excel object, macro, external link, embedded object, or advanced chart.
- Chart support currently provides a range-based preview panel, not an embedded editable chart object inside the grid.
- Pivot support currently creates a grouped summary worksheet. It is not a native, refreshable drag-and-drop PivotTable with slicers.
- Manual browser QA is still required for column-header context-menu behavior, formula/reference preservation after insert/delete, touch behavior, browser-specific IndexedDB quotas, large workbooks, and complex XLSX files.

## Architecture notes

- Spreadsheet rendering, Ribbon, core formula calculation, sheet tabs, and core context menus are provided by Univer Sheets Core.
- Application-level file and storage adapters live in `src/services/`.
- Google Drive should be added later behind a storage adapter so the editor does not depend on a specific cloud provider.
- The project intentionally does not include paid Univer Pro exchange, chart, pivot, or collaboration modules.

## Licensing

Review the license of every dependency before distributing the application. The open-source Univer packages used here are licensed separately from its commercial Pro modules. GridNova does not include commercial Univer modules.
