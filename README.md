# GridNova

GridNova is a React + Vite spreadsheet application powered by Univer 1.5.2. Workbook state is owned by the Univer worksheet facade; snapshots are persisted locally through IndexedDB.

## CSV and PDF import/export

The File, Data and Share ribbon tabs include CSV/PDF actions.

- **Data → Get Data → Get Data from CSV** opens a drag-and-drop/file-picker preview, detects comma/semicolon/tab delimiters, previews the first five rows and imports either by replacing the active sheet or writing from the selected cell.
- **Data → Get Data → Get Data from PDF** extracts selectable text from all PDF pages, groups positioned text into rows/cells and shows an import preview. Scanned image PDFs are rejected with an OCR requirement message; OCR is not bundled.
- **File → Export Active Sheet CSV** and **Share → Export as CSV** export the active worksheet's calculated cell values.
- **File → Export Active Sheet PDF** and **Share → Export as PDF** open a Portrait/Landscape selector and export the active worksheet with supported cell text, fill and alignment styling.

### Install and run

The dependencies are already declared in `package.json`. On a clean checkout run:

```bash
npm install
npm run dev
```

Run the automated tests and production build:

```bash
npm test
npm run build
```

The PDF.js worker is bundled by Vite and loaded only when a PDF is parsed.

### Implementation map

- `src/services/dataFileService.js` — Papa Parse CSV handling, PDF.js text extraction and jsPDF-AutoTable export.
- `src/components/DataFileImportModal.jsx` — upload/dropzone, preview and replace/insert selection.
- `src/components/PdfExportDialog.jsx` — PDF orientation selector.
- `src/components/UniverEditor.jsx` — Univer facade integration for selected-cell imports and worksheet expansion.
- `src/pages/WorkbookEditor.jsx` — ribbon actions, calculated-value exports, formula recalc and local workbook persistence.
- `tests/dataFileService.test.js` — CSV parsing/escaping and PDF text-grouping/style tests.

### Data and compatibility notes

Imports and exports run in the browser. CSV/PDF content is not uploaded to a server. Imported CSV/PDF values are inserted as text so a file cannot silently execute as a cell formula. CSV export escapes formula-like text fields while preserving numeric values, including negative numbers. Formula calculation is requested from the installed Univer formula service before export.

PDF import is intended for text-based PDFs whose text has meaningful coordinates. PDF table reconstruction uses positional heuristics (Y baselines and X gaps); unusual multi-column layouts, merged visual cells, rotated text, or complex financial statements may need manual review in the preview. Scanned/image-only PDFs require an external OCR pass first.

PDF export preserves supported cell bold/italic, text/fill colors, horizontal alignment and wrapping. It does not reproduce every Excel feature (for example, merged-cell geometry, conditional-format rendering, embedded drawings or print areas). Large files are subject to a 2,000,000-cell import limit to avoid exhausting browser memory.
