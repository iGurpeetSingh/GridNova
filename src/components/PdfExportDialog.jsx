import { useState } from 'react'

export default function PdfExportDialog({ onClose, onExport, busy = false }) {
  const [orientation, setOrientation] = useState('landscape')

  return (
    <div className="data-file-modal-backdrop" role="presentation" onMouseDown={(event) => {
      if (event.target === event.currentTarget && !busy) onClose()
    }}>
      <section className="data-file-modal data-file-export-modal" role="dialog" aria-modal="true" aria-labelledby="pdf-export-title">
        <header className="data-file-modal-header">
          <div>
            <span className="data-file-eyebrow">GRIDNOVA · FILE EXPORT</span>
            <h2 id="pdf-export-title">Export worksheet as PDF</h2>
            <p>Export calculated, displayed cell values with supported font, fill and alignment styling.</p>
          </div>
          <button type="button" className="data-file-close" onClick={onClose} disabled={busy} aria-label="Close PDF export dialog">×</button>
        </header>
        <div className="pdf-export-options">
          <span className="pdf-export-label">Page orientation</span>
          <label className={orientation === 'portrait' ? 'selected' : ''}>
            <input type="radio" name="pdf-export-orientation" value="portrait" checked={orientation === 'portrait'} onChange={() => setOrientation('portrait')} />
            <span className="pdf-paper-preview portrait" aria-hidden="true" />
            <span><strong>Portrait</strong><small>Vertical page</small></span>
          </label>
          <label className={orientation === 'landscape' ? 'selected' : ''}>
            <input type="radio" name="pdf-export-orientation" value="landscape" checked={orientation === 'landscape'} onChange={() => setOrientation('landscape')} />
            <span className="pdf-paper-preview landscape" aria-hidden="true" />
            <span><strong>Landscape</strong><small>Horizontal page</small></span>
          </label>
        </div>
        <footer className="data-file-modal-footer">
          <span>Large worksheets may create multiple PDF pages.</span>
          <div>
            <button type="button" className="data-file-cancel" onClick={onClose} disabled={busy}>Cancel</button>
            <button type="button" className="data-file-confirm" onClick={() => onExport(orientation)} disabled={busy}>{busy ? 'Exporting…' : 'Download PDF'}</button>
          </div>
        </footer>
      </section>
    </div>
  )
}
