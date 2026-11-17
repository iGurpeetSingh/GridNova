import { useCallback, useEffect, useRef, useState } from 'react'
import { useNavigate, useParams, useSearchParams } from 'react-router-dom'
import { BookOpenRegular, CommentRegular, GridRegular, ShareRegular, TableRegular } from '@fluentui/react-icons'
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
  const [ribbonMode, setRibbonMode] = useState(() => {
    try { return window.localStorage.getItem('gridnova:ribbon-mode:' + workbookId) || 'classic' } catch { return 'classic' }
  })
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

  async function handleRibbonModeChange(nextMode) {
    if (nextMode === ribbonMode || !['classic', 'collapsed'].includes(nextMode)) return
    const snapshot = editorApi?.getSnapshot()
    if (snapshot) {
      snapshot.name = title
      setSaveState('Saving before changing ribbon…')
      try {
        await saveWorkbook({ id: workbookId, name: title, snapshot, updatedAt: Date.now() })
        lastSavedRef.current = JSON.stringify(snapshot)
        setInitialSnapshot(snapshot)
        setSavedAt(Date.now())
        setSaveState('Saved in this browser')
      } catch (error) {
        setNoticeType('error')
        setNotice(error instanceof Error ? error.message : 'Could not save before changing the ribbon.')
        return
      }
    }
    setRibbonMode(nextMode)
    try { window.localStorage.setItem('gridnova:ribbon-mode:' + workbookId, nextMode) } catch { /* preference storage is optional */ }
  }

  async function handleTopBarCommand(commandId, successMessage) {
    try {
      const result = await editorApi?.executeCommand(commandId)
      if (result === false) throw new Error('This command is not available in the current workbook state.')
      setNoticeType('success')
      setNotice(successMessage)
    } catch (error) {
      setNoticeType('error')
      setNotice(error instanceof Error ? error.message : 'The command could not be completed.')
    }
  }

  async function handleShareWorkbook() {
    try {
      const snapshot = editorApi?.getSnapshot()
      if (!snapshot) throw new Error('The workbook is not ready yet.')
      snapshot.name = title
      const fileName = safeFileName(title) + '.xlsx'
      const file = await exportExcelFile(snapshot, fileName, { download: false })
      if (typeof navigator.share === 'function' && typeof navigator.canShare === 'function' && navigator.canShare({ files: [file] })) {
        await navigator.share({ files: [file], title, text: 'GridNova workbook' })
        setNoticeType('success')
        setNotice('Workbook shared.')
      } else {
        await exportExcelFile(snapshot, fileName)
        setNoticeType('success')
        setNotice('File sharing is not supported here. A shareable Excel copy was downloaded.')
      }
    } catch (error) {
      if (error?.name === 'AbortError') return
      setNoticeType('error')
      setNotice(error instanceof Error ? error.message : 'Could not prepare the workbook for sharing.')
    }
  }

  async function handleRibbonAction(action) {
    if (action === 'insert-tools') {
      setToolsOpen(true)
      return
    }
    if (action === 'save') {
      await saveNow({ force: true })
      return
    }
    if (action === 'import') {
      fileInputRef.current?.click()
      return
    }
    if (action === 'export-xlsx') {
      await handleExportExcel()
      return
    }
    if (action === 'export-csv') {
      handleExportCsv()
      return
    }
    if (action === 'print') {
      try {
        const opened = await editorApi?.executeCommand('sheet.operation.print-open')
        if (opened === false) window.print()
      } catch {
        window.print()
      }
      return
    }
    if (action === 'copy-workbook-id') {
      try {
        await navigator.clipboard.writeText(workbookId)
        setNoticeType('success')
        setNotice('Workbook ID copied.')
      } catch {
        setNoticeType('error')
        setNotice('Clipboard access is unavailable in this browser.')
      }
      return
    }
    if (action === 'open-help') {
      window.open('https://docs.univer.ai/guides/sheets', '_blank', 'noopener,noreferrer')
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
        <div className="editor-cloud-actions">
          <button type="button" onClick={() => handleTopBarCommand('sheet.operation.toggle-comment-panel', 'Comments panel opened.')}><CommentRegular /> <span>Comments</span></button>
          <button type="button" onClick={() => handleTopBarCommand('univer.operation.toggle-edit-history', 'Edit history opened.')}><BookOpenRegular /> <span>Catch up</span></button>
          <span className="editor-mode-indicator">Editing</span>
          <button className="share-action" type="button" onClick={handleShareWorkbook}><ShareRegular /> <span>Share</span></button>
        </div>
        <div className="ribbon-mode-switch" role="group" aria-label="Ribbon display mode">
          <button type="button" className={ribbonMode === 'collapsed' ? 'active' : ''} aria-pressed={ribbonMode === 'collapsed'} onClick={() => handleRibbonModeChange('collapsed')} title="Single Line Ribbon">
            <TableRegular /> <span>Single Line</span>
          </button>
          <button type="button" className={ribbonMode === 'classic' ? 'active' : ''} aria-pressed={ribbonMode === 'classic'} onClick={() => handleRibbonModeChange('classic')} title="Classic Ribbon">
            <GridRegular /> <span>Classic</span>
          </button>
        </div>
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
            ribbonType={ribbonMode}
            onReady={setEditorApi}
            onRibbonAction={handleRibbonAction}
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
