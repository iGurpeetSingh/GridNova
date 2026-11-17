import { useMemo } from 'react'
import { useNavigate, useParams, useSearchParams } from 'react-router-dom'
import UniverEditor from '../components/UniverEditor.jsx'

export default function WorkbookEditor() {
  const navigate = useNavigate()
  const { workbookId } = useParams()
  const [searchParams] = useSearchParams()

  const title = useMemo(() => {
    const requestedTitle = searchParams.get('title')
    return requestedTitle?.trim() || 'Untitled workbook'
  }, [searchParams])

  return (
    <main className="editor-page">
      <header className="editor-appbar">
        <button className="editor-brand" onClick={() => navigate('/')} type="button" aria-label="Back to GridNova home">
          <span className="brand-mark small">G</span>
          <span>Grid<span>Nova</span></span>
        </button>
        <span className="appbar-separator" />
        <div className="workbook-heading">
          <strong>{title}</strong>
          <span>Workbook · {String(workbookId).slice(0, 8)}</span>
        </div>
        <div className="editor-appbar-spacer" />
        <span className="storage-notice"><span className="notice-dot" /> Cloud saving not connected yet</span>
        <button className="dashboard-return" type="button" onClick={() => navigate('/')}>Dashboard</button>
      </header>

      <section className="editor-workspace" aria-label="Spreadsheet editor">
        <UniverEditor />
      </section>

      <div className="mobile-editor-notice">
        <div className="mobile-notice-icon">▦</div>
        <h2>Open GridNova on a larger screen</h2>
        <p>The current GridNova build is optimized for desktop and tablet editing. Phone access is view-only in this product phase.</p>
        <button type="button" onClick={() => navigate('/')}>Back to dashboard</button>
      </div>
    </main>
  )
}
