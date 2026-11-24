import { lazy, Suspense } from 'react'
import { BrowserRouter, Navigate, Route, Routes } from 'react-router-dom'
import Dashboard from './pages/Dashboard.jsx'
import GridNovaLoader from './components/GridNovaLoader.jsx'

const WorkbookEditor = lazy(() => import('./pages/WorkbookEditor.jsx'))

export default function App() {
  return (
    <BrowserRouter>
      <Routes>
        <Route path="/" element={<Dashboard />} />
        <Route
          path="/workbook/:workbookId"
          element={(
            <Suspense fallback={<GridNovaLoader fullScreen label="Loading spreadsheet editor" />}>
              <WorkbookEditor />
            </Suspense>
          )}
        />
        <Route path="*" element={<Navigate to="/" replace />} />
      </Routes>
    </BrowserRouter>
  )
}
