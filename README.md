# GridNova

GridNova is a desktop-first spreadsheet workspace built with React, Vite, and the open-source Univer Sheets Core.

## Current milestone

- Desktop dashboard shell
- Real workbook editor powered by Univer Sheets Core
- Workbook ribbon, formula bar, grid, sheet tabs, and core editing interactions supplied by the engine
- Tablet-friendly layout; phone opens a view-only notice by product decision

## Run locally

```bash
npm install
npm run dev
```

## Build

```bash
npm run build
npm run preview
```

## Architecture notes

- The spreadsheet engine is isolated in `src/components/UniverEditor.jsx`.
- The application shell and dashboard are independent from the engine.
- Google Drive integration is intentionally deferred until the editor milestone is complete.
- There is no persistent cloud save yet. Do not treat an open editor session as a saved workbook.

## Open-source licensing

The Univer open-source packages are distributed under Apache-2.0. Commercial Univer modules/features are not included in this initial setup. Review the license of every added dependency before shipping.
