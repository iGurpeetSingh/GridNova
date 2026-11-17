import { useCallback, useEffect, useRef, useState } from 'react'
import { useNavigate, useParams, useSearchParams } from 'react-router-dom'
import UniverEditor from '../components/UniverEditor.jsx'
import InsertToolsPanel from '../components/InsertToolsPanel.jsx'
import { exportCsvFile, exportExcelFile, importExcelFile } from '../services/excelFileService.js'
import { getWorkbook, saveWorkbook } from '../services/workbookStorage.js'

function safeFileName(value) {
  return String(value || 'GridNova Workbook').replace(/[\\/:*?"<>|]/g, '-').trim() || 'GridNova Workbook'
}

export default function WorkbookEditor() {
  const navigate = useNavigate()
  const { workbookId } = useParams()
  const [searchParams] = useSearchParams()
  const fileInputRef = useRef(null)
  const lastSavedRef = useRef('')
  const saveInFlightRef = useRef(null)
  const [title, setTitle] = useState(searchParams.get('title')?.trim() || 'Untitled workbook')
  const [initialSnapshot, setInitialSnapshot] = useState(null)
  const [loading, setLoading] = useState(true)
  const [editorApi, setEditorApi] = useState(null)
  const [saveState, setSaveState] = useState('Loading workbook')
  const [savedAt, setSavedAt] = useState(null)
  const [toolsOpen, setToolsOpen] = useState(false)
  const [notice, setNotice] = useState('')
  const [noticeType, setNoticeType] = useState('success')

  useEffect(() => {
    let mounted = true
    getWorkbook(workbookId)
      .then((record) => {
        if (!mounted) return
        if (record?.snapshot) {
          setInitialSnapshot(record.snapshot)
          setTitle(record.name || record.snapshot.name || 'Untitled workbook')
          lastSavedRef.current = JSON.stringify(record.snapshot)
          setSavedAt(record.updatedAt || null)
          setSaveState('Saved in this browser')
        } else {
          setSaveState('New workbook')
        }
      })
      .catch((error) => {
        console.error('Could not load GridNova workbook:', error)
        if (mounted) setSaveState('Storage unavailable')
      })
      .finally(() => {
        if (mounted) setLoading(false)
      })
    return () => { mounted = false }
  }, [workbookId])

  const saveNow = useCallback(async (options = {}) => {
    if (!editorApi) return false
    if (saveInFlightRef.current) return saveInFlightRef.current

    const task = (async () => {
      const snapshot = editorApi.getSnapshot()
      if (!snapshot) return false
      snapshot.name = title
      const serialized = JSON.stringify(snapshot)
      if (!options.force && serialized === lastSavedRef.current) return true

      setSaveState('Saving…')
      try {
        await saveWorkbook({
          id: workbookId,
          name: title,
          snapshot,
          updatedAt: Date.now(),
        })
        lastSavedRef.current = serialized
        setSavedAt(Date.now())
        setSaveState('Saved in this browser')
        return true
      } catch (error) {
        console.error('GridNova autosave failed:', error)
        setSaveState('Save failed')
        setNoticeType('error')
        setNotice(error instanceof Error ? error.message : 'Could not save this workbook in browser storage.')
        return false
      }
    })()

    saveInFlightRef.current = task
    try {
      return await task
    } finally {
      if (saveInFlightRef.current === task) saveInFlightRef.current = null
    }
  }, [editorApi, title, workbookId])

  useEffect(() => {
    if (!editorApi) return undefined
    const timer = window.setInterval(() => {
      saveNow()
    }, 1800)
    return () => window.clearInterval(timer)
  }, [editorApi, saveNow])

  useEffect(() => {
    if (!editorApi) return undefined
    const persistOnPageHide = () => {
      const snapshot = editorApi.getSnapshot()
      if (!snapshot) return
      snapshot.name = title
      saveWorkbook({ id: workbookId, name: title, snapshot }).catch((error) => {
        console.error('GridNova page-hide save failed:', error)
      })
    }
    window.addEventListener('pagehide', persistOnPageHide)
    return () => window.removeEventListener('pagehide', persistOnPageHide)
  }, [editorApi, title, workbookId])


  useEffect(() => {
    if (!editorApi) return undefined
    const handleKeyDown = (event) => {
      if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 's') {
        event.preventDefault()
        saveNow({ force: true })
      }
    }
    window.addEventListener('keydown', handleKeyDown)
    return () => window.removeEventListener('keydown', handleKeyDown)
  }, [editorApi, saveNow])

  async function handleImport(event) {
    const file = event.target.files?.[0]
    event.target.value = ''
    if (!file || !editorApi) return
    setNotice('')
    try {
      const importedSnapshot = await importExcelFile(file)
      importedSnapshot.id = 'gridnova-import-' + (globalThis.crypto?.randomUUID?.() || Date.now())
      editorApi.replaceSnapshot(importedSnapshot)
      const importedTitle = importedSnapshot.name || file.name.replace(/\\.(xlsx|xls|csv|tsv)$/i, '')
      setTitle(importedTitle)
      const importedCurrentSnapshot = editorApi.getSnapshot()
      importedCurrentSnapshot.name = importedTitle
      await saveWorkbook({ id: workbookId, name: importedTitle, snapshot: importedCurrentSnapshot })
      lastSavedRef.current = JSON.stringify(importedCurrentSnapshot)
      setSavedAt(Date.now())
      setNoticeType('success')
      setNotice('Imported ' + file.name + ' and saved the workbook in this browser.')
      setSaveState('Saved in this browser')
    } catch (error) {
      console.error('GridNova import failed:', error)
      setNoticeType('error')
      setNotice(error instanceof Error ? error.message : 'The selected file could not be imported.')
    }
  }

  async function handleExportExcel() {
    try {
      const snapshot = editorApi?.getSnapshot()
      if (!snapshot) throw new Error('The workbook is not ready yet.')
      snapshot.name = title
      await exportExcelFile(snapshot, safeFileName(title) + '.xlsx')
      setNoticeType('success')
      setNotice('Excel workbook exported.')
    } catch (error) {
      setNoticeType('error')
      setNotice(error instanceof Error ? error.message : 'Could not export the workbook.')
    }
  }

  function handleExportCsv() {
    try {
      const snapshot = editorApi?.getSnapshot()
      if (!snapshot) throw new Error('The workbook is not ready yet.')
      const sheetId = editorApi.getActiveSheet()?.getSheetId()
      exportCsvFile(snapshot, safeFileName(title) + '.csv', sheetId)
      setNoticeType('success')
      setNotice('Active worksheet exported as CSV.')
    } catch (error) {
      setNoticeType('error')
      setNotice(error instanceof Error ? error.message : 'Could not export CSV.')
    }
  }

  async function handleBackToDashboard() {
    await saveNow({ force: true })
    navigate('/')
  }

  return (
    <main className="editor-page">
      <header className="editor-appbar">
        <button className="editor-brand" onClick={handleBackToDashboard} type="button" aria-label="Back to GridNova dashboard">
          <span className="brand-mark small">G</span>
          <span>Grid<span>Nova</span></span>
        </button>
        <span className="appbar-separator" />
        <div className="workbook-heading">
          <input
            aria-label="Workbook name"
            value={title}
            onChange={(event) => setTitle(event.target.value)}
            onBlur={() => saveNow({ force: true })}
            maxLength={120}
          />
          <span>Workbook · {String(workbookId).slice(0, 8)}</span>
        </div>
        <div className="editor-appbar-spacer" />
        <span className={'storage-notice ' + (saveState === 'Save failed' ? 'storage-error' : '')}>
          <span className="notice-dot" /> {saveState}
          {savedAt ? ' · ' + new Date(savedAt).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }) : ''}
        </span>
        <div className="editor-action-group">
          <input ref={fileInputRef} className="visually-hidden" type="file" accept=".xlsx,.csv,.tsv" onChange={handleImport} />
          <button className="editor-action-button" type="button" onClick={() => fileInputRef.current?.click()}>Import</button>
          <button className="editor-action-button" type="button" onClick={handleExportCsv}>CSV</button>
          <button className="editor-action-button" type="button" onClick={handleExportExcel}>Export XLSX</button>
          <button className="editor-action-button insert-action" type="button" onClick={() => setToolsOpen(true)}>Insert tools</button>
          <button className="editor-save-button" type="button" onClick={() => saveNow({ force: true })}>Save now</button>
        </div>
        <button className="dashboard-return" type="button" onClick={handleBackToDashboard}>Dashboard</button>
      </header>

      {notice && <div className={'editor-notice ' + noticeType} role="status">{notice}<button type="button" onClick={() => setNotice('')} aria-label="Dismiss message">×</button></div>}

      <section className="editor-workspace" aria-label="Spreadsheet editor">
        {loading ? (
          <div className="editor-loading"><span className="loading-spinner" /> Loading workbook…</div>
        ) : (
          <UniverEditor
            workbookId={workbookId}
            title={title}
            initialSnapshot={initialSnapshot}
            onReady={setEditorApi}
          />
        )}
      </section>

      <div className="mobile-editor-notice">
        <div className="mobile-notice-icon">▦</div>
        <h2>Open GridNova on a larger screen</h2>
        <p>The current GridNova build is optimized for desktop and tablet editing. Phone access is view-only in this product phase.</p>
        <button type="button" onClick={() => navigate('/')}>Back to dashboard</button>
      </div>

      {toolsOpen && <InsertToolsPanel api={editorApi} onClose={() => setToolsOpen(false)} />}
    </main>
  )
}
