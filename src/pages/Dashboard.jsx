import { useEffect, useState } from 'react'
import { useNavigate } from 'react-router-dom'
import { deleteWorkbook, listWorkbooks, saveWorkbook } from '../services/workbookStorage.js'

const templates = [
  { icon: '▦', title: 'Blank workbook', detail: 'Start with an empty spreadsheet', tone: 'green', ready: true },
  { icon: '↗', title: 'Monthly budget', detail: 'Organize income and expenses', tone: 'blue', ready: true },
  { icon: '◫', title: 'Project tracker', detail: 'Track tasks, owners and dates', tone: 'violet', ready: true },
  { icon: '▥', title: 'Sales report', detail: 'Review performance and totals', tone: 'amber', ready: true },
]

function makeId() {
  return globalThis.crypto?.randomUUID?.() ?? ('workbook-' + Date.now() + '-' + Math.random().toString(36).slice(2))
}

function makeTemplateSnapshot(title) {
  const sheetId = 'sheet1'
  let rows = []
  if (title === 'Monthly budget') {
    rows = [
      ['Category', 'Budget', 'Actual', 'Difference'],
      ['Housing', 15000, 14500, '=B2-C2'],
      ['Food', 8000, 7250, '=B3-C3'],
      ['Transport', 4000, 3600, '=B4-C4'],
      ['Utilities', 3000, 2850, '=B5-C5'],
      ['Entertainment', 2500, 2100, '=B6-C6'],
      ['Total', '=SUM(B2:B6)', '=SUM(C2:C6)', '=SUM(D2:D6)'],
    ]
  } else if (title === 'Project tracker') {
    rows = [
      ['Task', 'Owner', 'Status', 'Due date'],
      ['Research', 'Team', 'Complete', '2026-11-20'],
      ['Design', 'Product', 'In progress', '2026-11-25'],
      ['Development', 'Engineering', 'Not started', '2026-12-05'],
      ['QA', 'QA team', 'Not started', '2026-12-12'],
    ]
  } else if (title === 'Sales report') {
    rows = [
      ['Month', 'Revenue', 'Orders'],
      ['January', 125000, 42],
      ['February', 142000, 48],
      ['March', 138000, 45],
      ['April', 168000, 56],
      ['May', 179000, 61],
    ]
  }

  const cellData = {}
  rows.forEach((row, rowIndex) => {
    const cells = {}
    row.forEach((value, columnIndex) => {
      if (value !== '' && value !== null && value !== undefined) {
        cells[columnIndex] = typeof value === 'string' && value.startsWith('=')
          ? { f: value.slice(1), v: null }
          : { v: value }
      }
    })
    cellData[rowIndex] = cells
  })

  return {
    id: 'gridnova-template-' + makeId(),
    name: title,
    appVersion: '1.0.0',
    locale: 'enUS',
    styles: {},
    sheetOrder: [sheetId],
    sheets: {
      [sheetId]: {
        id: sheetId,
        name: 'Sheet1',
        tabColor: '',
        hidden: 0,
        rowCount: 1000,
        columnCount: 26,
        zoomRatio: 1,
        freeze: { startRow: -1, startColumn: -1, ySplit: 0, xSplit: 0 },
        scrollTop: 0,
        scrollLeft: 0,
        defaultColumnWidth: 100,
        defaultRowHeight: 23,
        mergeData: [],
        cellData,
        rowData: {},
        columnData: {},
        rowHeader: { width: 46, hidden: 0 },
        columnHeader: { height: 20, hidden: 0 },
        showGridlines: 1,
        rightToLeft: 0,
      },
    },
  }
}

function formatDate(timestamp) {
  if (!timestamp) return 'Just created'
  return new Date(timestamp).toLocaleString([], { month: 'short', day: 'numeric', hour: '2-digit', minute: '2-digit' })
}

export default function Dashboard() {
  const navigate = useNavigate()
  const [workbooks, setWorkbooks] = useState([])
  const [loadingFiles, setLoadingFiles] = useState(true)

  async function refreshWorkbooks() {
    try {
      setWorkbooks(await listWorkbooks())
    } catch (error) {
      console.error('Could not load GridNova workbook list:', error)
      setWorkbooks([])
    } finally {
      setLoadingFiles(false)
    }
  }

  useEffect(() => {
    refreshWorkbooks()
  }, [])

  async function openWorkbook(title = 'Untitled workbook', snapshot = null) {
    const id = makeId()
    if (snapshot) {
      snapshot.id = id
      snapshot.name = title
      try {
        await saveWorkbook({ id, name: title, snapshot })
      } catch (error) {
        window.alert(error instanceof Error ? error.message : 'Could not create the template workbook.')
        return
      }
    }
    navigate('/workbook/' + id + '?title=' + encodeURIComponent(title))
  }

  function openExisting(record) {
    navigate('/workbook/' + record.id + '?title=' + encodeURIComponent(record.name || 'Untitled workbook'))
  }

  async function removeWorkbook(event, id) {
    event.stopPropagation()
    if (!window.confirm('Move this workbook out of the GridNova recent list?')) return
    try {
      await deleteWorkbook(id)
      await refreshWorkbooks()
    } catch (error) {
      window.alert(error instanceof Error ? error.message : 'Could not remove workbook.')
    }
  }

  return (
    <main className="dashboard-page">
      <aside className="dashboard-sidebar">
        <a className="brand-lockup" href="/" aria-label="GridNova home">
          <span className="brand-mark">G</span>
          <span>Grid<span>Nova</span></span>
        </a>
        <div className="sidebar-section-label">WORKSPACE</div>
        <nav className="sidebar-nav" aria-label="Main navigation">
          <a className="sidebar-link active" href="/">
            <span className="nav-icon">⌂</span><span>Home</span>
          </a>
          <button className="sidebar-link" type="button" onClick={() => openWorkbook()}>
            <span className="nav-icon">▦</span><span>New workbook</span>
          </button>
          <a className="sidebar-link" href="#recent-workbooks">
            <span className="nav-icon">▤</span><span>My workbooks</span>
          </a>
        </nav>
        <div className="sidebar-footer">
          <div className="sidebar-status-dot" />
          <div>
            <strong>Browser draft storage</strong>
            <span>Google Drive integration later</span>
          </div>
        </div>
      </aside>

      <section className="dashboard-main">
        <header className="dashboard-topbar">
          <div className="breadcrumbs"><span>Workspace</span><span className="breadcrumb-divider">/</span><strong>Home</strong></div>
          <div className="topbar-actions">
            <span className="preview-pill"><span /> Development preview</span>
            <div className="avatar" aria-label="GridNova user">G</div>
          </div>
        </header>

        <div className="dashboard-content">
          <section className="welcome-row">
            <div>
              <p className="eyebrow">YOUR SPREADSHEET WORKSPACE</p>
              <h1>Good work starts with a <span>clear grid.</span></h1>
              <p className="welcome-copy">Create a workbook, organize information, and turn raw data into useful work.</p>
            </div>
            <button className="primary-button" onClick={() => openWorkbook()} type="button">
              <span className="button-plus">＋</span> New workbook
            </button>
          </section>

          <section className="template-section">
            <div className="section-heading">
              <div>
                <h2>Start something new</h2>
                <p>Choose a starting point for your next workbook.</p>
              </div>
              <span className="section-note">4 STARTING POINTS</span>
            </div>
            <div className="template-grid">
              {templates.map((template) => (
                <button
                  className="template-card template-card-ready"
                  key={template.title}
                  type="button"
                  onClick={() => openWorkbook(
                    template.title === 'Blank workbook' ? 'Untitled workbook' : template.title,
                    template.title === 'Blank workbook' ? null : makeTemplateSnapshot(template.title),
                  )}
                >
                  <span className={'template-preview ' + template.tone}>
                    <span className="preview-toolbar"><i /><i /><i /><i /></span>
                    <span className="preview-sheet">{Array.from({ length: 20 }, (_, index) => <i key={index} />)}</span>
                  </span>
                  <span className="template-meta">
                    <span className="template-title">{template.title}</span>
                    <span className="template-detail">{template.detail}</span>
                  </span>
                  <span className={'template-icon ' + template.tone}>{template.icon}</span>
                </button>
              ))}
            </div>
          </section>

          <section className="recent-section" id="recent-workbooks">
            <div className="section-heading recent-heading">
              <div>
                <h2>Recent workbooks</h2>
                <p>Open a workbook saved in this browser.</p>
              </div>
              <button className="refresh-files-button" type="button" onClick={refreshWorkbooks}>Refresh</button>
            </div>
            {loadingFiles ? (
              <div className="files-loading">Loading workbooks…</div>
            ) : workbooks.length ? (
              <div className="workbook-list">
                {workbooks.map((record) => (
                  <div className="workbook-list-row" key={record.id} role="button" tabIndex={0} onClick={() => openExisting(record)} onKeyDown={(event) => {
                    if (event.key === 'Enter' || event.key === ' ') openExisting(record)
                  }}>
                    <div className="workbook-file-icon">▦</div>
                    <div className="workbook-file-name">
                      <strong>{record.name || 'Untitled workbook'}</strong>
                      <span>{formatDate(record.updatedAt)}</span>
                    </div>
                    <span className="workbook-file-kind">Spreadsheet</span>
                    <button className="file-delete-button" type="button" aria-label={'Remove ' + record.name} onClick={(event) => removeWorkbook(event, record.id)}>×</button>
                  </div>
                ))}
              </div>
            ) : (
              <div className="empty-workbooks">
                <div className="empty-icon">▦</div>
                <h3>No workbooks yet</h3>
                <p>Create your first workbook to open the GridNova editor.</p>
                <button className="secondary-button" onClick={() => openWorkbook()} type="button">Create a blank workbook</button>
              </div>
            )}
          </section>

          <footer className="dashboard-footer">
            <span>GridNova</span>
            <span>Workbook drafts are currently stored in this browser.</span>
          </footer>
        </div>
      </section>
    </main>
  )
}
