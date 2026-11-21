import { useRef, useState } from 'react'
import { parseCsvFile, parsePdfFile } from '../services/dataFileService.js'

function formatDelimiter(value) {
  if (value === '\t') return 'Tab'
  if (value === ';') return 'Semicolon'
  return 'Comma'
}

export default function DataFileImportModal({ fileType, onClose, onImport }) {
  const inputRef = useRef(null)
  const [fileName, setFileName] = useState('')
  const [rows, setRows] = useState(null)
  const [metadata, setMetadata] = useState(null)
  const [mode, setMode] = useState('replace')
  const [busy, setBusy] = useState(false)
  const [progress, setProgress] = useState(0)
  const [error, setError] = useState('')
  const [dragging, setDragging] = useState(false)

  const isPdf = fileType === 'pdf'
  const accepted = isPdf ? '.pdf,application/pdf' : '.csv,.tsv,text/csv,text/tab-separated-values'
  const previewRows = rows?.slice(0, isPdf ? 10 : 5) || []

  async function loadFile(file) {
    if (!file) return
    const extension = String(file.name || '').split('.').pop().toLowerCase()
    const valid = isPdf
      ? extension === 'pdf' || file.type === 'application/pdf'
      : ['csv', 'tsv'].includes(extension)
    if (!valid) {
      setError(isPdf ? 'Only PDF files are supported here.' : 'Choose a .csv or .tsv file.')
      return
    }

    setBusy(true)
    setProgress(0)
    setError('')
    setRows(null)
    setMetadata(null)
    setFileName(file.name)
    try {
      const result = isPdf
        ? await parsePdfFile(file, { onProgress: setProgress })
        : await parseCsvFile(file, { onProgress: setProgress })
      setRows(result.rows)
      setMetadata(result)
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : 'The selected file could not be read.')
    } finally {
      setBusy(false)
    }
  }

  async function handleImport() {
    if (!rows?.length || busy) return
    setBusy(true)
    setError('')
    try {
      await onImport(rows, { mode, fileName })
      onClose()
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : 'Could not import this file into the worksheet.')
    } finally {
      setBusy(false)
    }
  }

  return (
    <div className="data-file-modal-backdrop" role="presentation" onMouseDown={(event) => {
      if (event.target === event.currentTarget && !busy) onClose()
    }}>
      <section className="data-file-modal" role="dialog" aria-modal="true" aria-labelledby="data-file-modal-title">
        <header className="data-file-modal-header">
          <div>
            <span className="data-file-eyebrow">GRIDNOVA · DATA IMPORT</span>
            <h2 id="data-file-modal-title">Get Data from {isPdf ? 'PDF' : 'CSV'}</h2>
            <p>{isPdf ? 'Extract selectable PDF text and arrange it into worksheet rows and columns.' : 'Import delimited data while preserving quoted commas, quotes and line breaks.'}</p>
          </div>
          <button type="button" className="data-file-close" onClick={onClose} disabled={busy} aria-label="Close import dialog">×</button>
        </header>

        {!rows && (
          <div
            className={'data-file-dropzone' + (dragging ? ' is-dragging' : '')}
            onDragOver={(event) => { event.preventDefault(); setDragging(true) }}
            onDragLeave={() => setDragging(false)}
            onDrop={(event) => {
              event.preventDefault()
              setDragging(false)
              loadFile(event.dataTransfer.files?.[0])
            }}
          >
            <div className="data-file-upload-icon">{isPdf ? 'PDF' : 'CSV'}</div>
            <strong>Drag and drop your {isPdf ? 'PDF' : 'CSV'} file here</strong>
            <span>or select a file from your computer</span>
            <button type="button" className="data-file-choose" onClick={() => inputRef.current?.click()} disabled={busy}>
              {busy ? 'Reading file…' : 'Choose file'}
            </button>
            <small>{isPdf ? 'Text-based PDFs are supported. Scanned image PDFs require OCR.' : 'CSV and TSV · comma, semicolon or tab delimiter'}</small>
            <input
              ref={inputRef}
              type="file"
              accept={accepted}
              className="visually-hidden"
              onChange={(event) => {
                const selected = event.target.files?.[0]
                event.target.value = ''
                loadFile(selected)
              }}
            />
          </div>
        )}

        {busy && (
          <div className="data-file-progress" role="status">
            <span className="loading-spinner" />
            <div><strong>{rows ? 'Writing data to the worksheet…' : 'Reading and processing file…'}</strong><span>{progress > 0 ? progress + '% complete' : 'Large files may take a little longer.'}</span></div>
          </div>
        )}

        {fileName && rows && !busy && (
          <div className="data-file-preview">
            <div className="data-file-preview-heading">
              <div><strong>{fileName}</strong><span>{rows.length.toLocaleString()} rows · {Math.max(1, ...rows.map((row) => row.length)).toLocaleString()} columns</span></div>
              <button type="button" onClick={() => { setRows(null); setFileName(''); setMetadata(null); setError('') }}>Choose another</button>
            </div>
            <div className="data-file-meta">
              {isPdf
                ? <span>{metadata?.pages || 1} PDF page(s) · {metadata?.extractedItems || 0} text blocks extracted</span>
                : <span>Detected delimiter: {formatDelimiter(metadata?.delimiter)}{metadata?.warnings?.length ? ' · ' + metadata.warnings.length + ' parsing warning(s)' : ''}</span>}
              <span>Preview: first {previewRows.length} rows</span>
            </div>
            <div className="data-file-table-wrap">
              <table className="data-file-table">
                <tbody>
                  {previewRows.map((row, rowIndex) => (
                    <tr key={rowIndex}>
                      <th>{rowIndex + 1}</th>
                      {row.map((value, columnIndex) => <td key={columnIndex}>{String(value ?? '')}</td>)}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            <fieldset className="data-file-import-mode">
              <legend>Import location</legend>
              <label><input type="radio" name="data-file-import-mode" checked={mode === 'replace'} onChange={() => setMode('replace')} /> <span><strong>Replace current sheet</strong><small>Clear current cells and place imported data at A1.</small></span></label>
              <label><input type="radio" name="data-file-import-mode" checked={mode === 'insert'} onChange={() => setMode('insert')} /> <span><strong>Insert at selected cell</strong><small>Keep existing data and insert starting at the current selection.</small></span></label>
            </fieldset>
          </div>
        )}

        {error && <div className="data-file-error" role="alert">{error}</div>}

        <footer className="data-file-modal-footer">
          <span>{rows ? 'Review the preview before importing.' : 'Your file is processed locally in this browser.'}</span>
          <div>
            <button type="button" className="data-file-cancel" onClick={onClose} disabled={busy}>Cancel</button>
            <button type="button" className="data-file-confirm" onClick={handleImport} disabled={!rows?.length || busy}>
              {busy && rows ? 'Importing…' : 'Import to sheet'}
            </button>
          </div>
        </footer>
      </section>
    </div>
  )
}
