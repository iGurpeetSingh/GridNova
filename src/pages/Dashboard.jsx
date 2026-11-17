import { useNavigate } from 'react-router-dom'

const templates = [
  { icon: '▦', title: 'Blank workbook', detail: 'Start with an empty spreadsheet', tone: 'green', ready: true },
  { icon: '↗', title: 'Monthly budget', detail: 'Organize income and expenses', tone: 'blue' },
  { icon: '◫', title: 'Project tracker', detail: 'Track tasks, owners and dates', tone: 'violet' },
  { icon: '▥', title: 'Sales report', detail: 'Review performance and totals', tone: 'amber' },
]

function makeId() {
  return globalThis.crypto?.randomUUID?.() ?? `workbook-${Date.now()}`
}

export default function Dashboard() {
  const navigate = useNavigate()

  function openWorkbook(title = 'Untitled workbook') {
    const id = makeId()
    navigate(`/workbook/${id}?title=${encodeURIComponent(title)}`)
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
          <button className="sidebar-link" type="button" onClick={() => window.alert('Workbook library will be connected when storage is implemented.')}>
            <span className="nav-icon">▤</span><span>My workbooks</span>
          </button>
        </nav>

        <div className="sidebar-footer">
          <div className="sidebar-status-dot" />
          <div>
            <strong>Editor preview</strong>
            <span>Cloud storage not connected</span>
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
              <span className="section-note">MORE TEMPLATES COMING LATER</span>
            </div>
            <div className="template-grid">
              {templates.map((template) => (
                <button
                  className={`template-card ${template.ready ? 'template-card-ready' : 'template-card-disabled'}`}
                  key={template.title}
                  type="button"
                  onClick={() => template.ready && openWorkbook(template.title === 'Blank workbook' ? 'Untitled workbook' : template.title)}
                  disabled={!template.ready}
                  aria-label={template.ready ? `Create ${template.title}` : `${template.title}, coming later`}
                >
                  <span className={`template-preview ${template.tone}`}>
                    <span className="preview-toolbar"><i /><i /><i /><i /></span>
                    <span className="preview-sheet">
                      {Array.from({ length: 20 }, (_, index) => <i key={index} />)}
                    </span>
                  </span>
                  <span className="template-meta">
                    <span className="template-title">{template.title}</span>
                    <span className="template-detail">{template.detail}</span>
                  </span>
                  <span className={`template-icon ${template.tone}`}>{template.icon}</span>
                </button>
              ))}
            </div>
          </section>

          <section className="recent-section">
            <div className="section-heading recent-heading">
              <div>
                <h2>Recent workbooks</h2>
                <p>Your recently opened workbooks will appear here.</p>
              </div>
            </div>
            <div className="empty-workbooks">
              <div className="empty-icon">▦</div>
              <h3>No workbooks yet</h3>
              <p>Create your first workbook to open the GridNova editor.</p>
              <button className="secondary-button" onClick={() => openWorkbook()} type="button">Create a blank workbook</button>
            </div>
          </section>

          <footer className="dashboard-footer">
            <span>GridNova</span>
            <span>Spreadsheet workspace · Early development</span>
          </footer>
        </div>
      </section>
    </main>
  )
}
