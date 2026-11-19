import { useCallback, useEffect, useRef, useState } from 'react'
import { useNavigate, useParams, useSearchParams } from 'react-router-dom'
import { ArrowRedoRegular, ArrowUndoRegular, BookOpenRegular, CommentRegular, GridRegular, ShareRegular, TableRegular } from '@fluentui/react-icons'
import UniverEditor from '../components/UniverEditor.jsx'
import InsertToolsPanel from '../components/InsertToolsPanel.jsx'
import ReviewPanel from '../components/ReviewPanel.jsx'
import { exportCsvFile, exportExcelFile, importExcelFile } from '../services/excelFileService.js'
import { getWorkbook, saveWorkbook } from '../services/workbookStorage.js'
import { isRibbonType, normalizeRibbonType, RIBBON_MODE_STORAGE_PREFIX, RIBBON_TYPE } from '../utils/ribbonMode.js'

function safeFileName(value) {
  return String(value || 'GridNova Workbook').replace(/[\\/:*?"<>|]/g, '-').trim() || 'GridNova Workbook'
}

function makeBlankSnapshot(id, name) {
  const sheetId = 'sheet1'
  return {
    id,
    name,
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
        cellData: {},
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

function makeWorkbookId() {
  return globalThis.crypto?.randomUUID?.() ?? ('workbook-' + Date.now() + '-' + Math.random().toString(36).slice(2))
}

export default function WorkbookEditor() {
  const navigate = useNavigate()
  const { workbookId } = useParams()
  const [searchParams] = useSearchParams()
  const fileInputRef = useRef(null)
  const lastSavedRef = useRef('')
  const saveInFlightRef = useRef(null)
  const [title, setTitle] = useState(searchParams.get('title')?.trim() || 'Untitled workbook')
  const [isReadOnly, setIsReadOnly] = useState(() => {
    try { return window.localStorage.getItem('gridnova:editing-mode:' + workbookId) === 'view' } catch { return false }
  })
  const [ribbonMode, setRibbonMode] = useState(() => {
    try {
      return normalizeRibbonType(window.localStorage.getItem(RIBBON_MODE_STORAGE_PREFIX + workbookId))
    } catch { return RIBBON_TYPE.SINGLE_LINE }
  })
  const [formulaBarVisible, setFormulaBarVisible] = useState(() => {
    try { return window.localStorage.getItem('gridnova:formula-bar:' + workbookId) !== 'false' } catch { return true }
  })
  const [initialSnapshot, setInitialSnapshot] = useState(null)
  const [loading, setLoading] = useState(true)
  const [editorApi, setEditorApi] = useState(null)
  const [saveState, setSaveState] = useState('Loading workbook')
  const [savedAt, setSavedAt] = useState(null)
  const [toolsOpen, setToolsOpen] = useState(false)
  const [reviewPanelMode, setReviewPanelMode] = useState(null)
  const [activity, setActivity] = useState(() => {
    try {
      const stored = JSON.parse(window.localStorage.getItem('gridnova:activity:' + workbookId) || '[]')
      return Array.isArray(stored) ? stored : []
    } catch { return [] }
  })
  const [notice, setNotice] = useState('')
  const [noticeType, setNoticeType] = useState('success')

  useEffect(() => {
    try {
      const stored = JSON.parse(window.localStorage.getItem('gridnova:activity:' + workbookId) || '[]')
      setActivity(Array.isArray(stored) ? stored : [])
    } catch {
      setActivity([])
    }
    setReviewPanelMode(null)
    try { setIsReadOnly(window.localStorage.getItem('gridnova:editing-mode:' + workbookId) === 'view') } catch { setIsReadOnly(false) }
    try { setFormulaBarVisible(window.localStorage.getItem('gridnova:formula-bar:' + workbookId) !== 'false') } catch { setFormulaBarVisible(true) }
  }, [workbookId])

  useEffect(() => {
    if (!editorApi?.setFormulaBarVisible) return
    editorApi.setFormulaBarVisible(formulaBarVisible)
  }, [editorApi, formulaBarVisible])

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
    if (!editorApi || !isReadOnly) return
    const sheets = editorApi.getWorkbook(workbookId)?.getSheets?.() || [editorApi.getActiveSheet()].filter(Boolean)
    Promise.all(sheets.map((sheet) => sheet.getWorksheetPermission?.()?.setReadOnly?.()))
      .catch((error) => {
        setNoticeType('error')
        setNotice(error instanceof Error ? error.message : 'Could not restore viewing mode.')
      })
  }, [editorApi, isReadOnly, workbookId])

  useEffect(() => {
    try { window.localStorage.setItem('gridnova:editing-mode:' + workbookId, isReadOnly ? 'view' : 'edit') } catch { /* preference storage is optional */ }
  }, [workbookId, isReadOnly])

  useEffect(() => {
    if (!editorApi?.onCommandExecuted) return undefined
    const subscription = editorApi.onCommandExecuted((event) => {
      const activeRange = editorApi.getActiveSheet()?.getActiveRange?.()
      const command = String(event?.id || 'Workbook action')
      const entry = {
        id: globalThis.crypto?.randomUUID?.() || String(Date.now()),
        command,
        label: command.split('.').pop().replace(/[-_]/g, ' '),
        type: event?.type || 'COMMAND',
        cell: activeRange?.getA1Notation?.() || '',
        createdAt: Date.now(),
      }
      setActivity((previous) => {
        const next = [entry, ...previous].slice(0, 100)
        try { window.localStorage.setItem('gridnova:activity:' + workbookId, JSON.stringify(next)) } catch { /* activity history is best-effort */ }
        return next
      })
    })
    return () => subscription?.dispose?.()
  }, [editorApi, workbookId])

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

  function handleEditorReady(api) {
    setEditorApi(api)
  }

  function handleRibbonModeChange(nextMode) {
    if (nextMode === ribbonMode || !isRibbonType(nextMode)) return
    if (!editorApi?.setRibbonType) {
      setNoticeType('error')
      setNotice('Ribbon layout switching is unavailable until the spreadsheet engine is ready.')
      return
    }

    try {
      // Switch the layout on the existing Univer instance. This keeps the
      // active worksheet, current selection, formulas and undo stack intact.
      editorApi.setRibbonType(nextMode)
      setRibbonMode(nextMode)
      try {
        window.localStorage.setItem(RIBBON_MODE_STORAGE_PREFIX + workbookId, nextMode)
      } catch { /* preference storage is optional */ }
    } catch (error) {
      setNoticeType('error')
      setNotice(error instanceof Error ? error.message : 'Could not change the ribbon layout.')
    }
  }

  async function handleQuickAccessCommand(commandId) {
    try {
      const result = await editorApi?.executeCommand(commandId)
      if (result === false) throw new Error('This command is not available right now.')
    } catch (error) {
      setNoticeType('error')
      setNotice(error instanceof Error ? error.message : 'The command could not be completed.')
    }
  }

  async function handleTopBarCommand(commandId, fallbackMode) {
    try {
      const result = await editorApi?.executeCommand(commandId)
      if (result === false) setReviewPanelMode(fallbackMode)
    } catch {
      setReviewPanelMode(fallbackMode)
    }
  }

  async function handleToggleEditingMode() {
    const sheets = editorApi?.getWorkbook(workbookId)?.getSheets?.() || [editorApi?.getActiveSheet()].filter(Boolean)
    if (!sheets.length) return
    const nextReadOnly = !isReadOnly
    try {
      await Promise.all(sheets.map((sheet) => {
        const permission = sheet.getWorksheetPermission?.()
        return nextReadOnly ? permission?.setReadOnly?.() : permission?.setEditable?.()
      }))
      setIsReadOnly(nextReadOnly)
    } catch (error) {
      setNoticeType('error')
      setNotice(error instanceof Error ? error.message : 'Could not change workbook editing mode.')
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
    if (action === 'open-comments') {
      await handleTopBarCommand('sheet.operation.toggle-comment-panel', 'comments')
      return
    }
    if (action === 'open-note') {
      await handleTopBarCommand('sheet.operation.add-note-popup', 'comments')
      return
    }
    if (action === 'open-catch-up') {
      setReviewPanelMode('activity')
      return
    }
    if (action === 'toggle-formula-bar') {
      const nextVisible = !formulaBarVisible
      const applied = editorApi?.setFormulaBarVisible?.(nextVisible)
      if (!applied) {
        setNoticeType('error')
        setNotice('Formula Bar could not be updated because the editor toolbar is not ready.')
        return
      }
      setFormulaBarVisible(nextVisible)
      try { window.localStorage.setItem('gridnova:formula-bar:' + workbookId, String(nextVisible)) } catch { /* local preference is optional */ }
      setNoticeType('success')
      setNotice(nextVisible ? 'Formula Bar shown.' : 'Formula Bar hidden.')
      return
    }
    if (action === 'share-workbook') {
      await handleShareWorkbook()
      return
    }
    if (action === 'save') {
      await saveNow({ force: true })
      return
    }
    if (action === 'browse-workbooks') {
      await saveNow({ force: true })
      navigate('/')
      return
    }
    if (action === 'new-workbook') {
      const nextId = makeWorkbookId()
      const nextTitle = 'Untitled workbook'
      try {
        await saveNow({ force: true })
        const snapshot = makeBlankSnapshot(nextId, nextTitle)
        await saveWorkbook({ id: nextId, name: nextTitle, snapshot })
        navigate('/workbook/' + nextId + '?title=' + encodeURIComponent(nextTitle))
      } catch (error) {
        setNoticeType('error')
        setNotice(error instanceof Error ? error.message : 'Could not create a new workbook.')
      }
      return
    }
    if (action === 'import') {
      fileInputRef.current?.click()
      return
    }
    if (action === 'save-as') {
      try {
        const snapshot = editorApi?.getSnapshot()
        if (!snapshot) throw new Error('The workbook is not ready yet.')
        const nextId = makeWorkbookId()
        const nextTitle = title + ' (Copy)'
        const copy = JSON.parse(JSON.stringify(snapshot))
        copy.id = nextId
        copy.name = nextTitle
        await saveWorkbook({ id: nextId, name: nextTitle, snapshot: copy })
        navigate('/workbook/' + nextId + '?title=' + encodeURIComponent(nextTitle))
      } catch (error) {
        setNoticeType('error')
        setNotice(error instanceof Error ? error.message : 'Could not save a workbook copy.')
      }
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
      // Univer 1.5.2 does not ship a native print-preview command in the
      // installed presets; use the browser's real print pipeline instead.
      window.print()
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
    if (action === 'copy-workbook-link') {
      try {
        await navigator.clipboard.writeText(window.location.href)
        setNoticeType('success')
        setNotice('Local workbook link copied. This workbook is stored in this browser only; the link does not sync it to other devices.')
      } catch {
        setNoticeType('error')
        setNotice('Clipboard access is unavailable in this browser.')
      }
      return
    }
    const unavailableActions = {
      'manage-access-unavailable': 'Manage Access',
      'sheet-protection-unavailable': 'Sheet / Workbook Protection',
      'draw-tools-unavailable': 'Pen, Highlighter and Eraser',
      'shape-format-unavailable': 'Shape Fill, Outline and Properties',
      'insert-text-unavailable': 'Text Boxes and Icons',
      'data-outline-unavailable': 'Group, Ungroup and Subtotal',
      'paste-special-unavailable': 'Paste Special',
      'cell-styles-unavailable': 'Cell Styles',
      'format-cells-unavailable': 'Format Cells',
      'page-layout-unavailable': 'Page Layout settings',
      'review-tools-unavailable': 'Spelling / Show Changes',
      'view-tools-unavailable': 'Page Layout View',
      'fill-unavailable': 'Fill commands',
      'formula-feature-unavailable': 'Formula Auditing and Calculation Options',
      'insert-charts-unavailable': 'Native editable Column, Line, Pie and Area charts',
    }
    if (unavailableActions[action]) {
      setNoticeType('error')
      setNotice(unavailableActions[action] + ' is not implemented in this GridNova build.')
      return
    }
    if (action === 'normal-view-current') {
      setNoticeType('success')
      setNotice('Normal spreadsheet view is already active.')
      return
    }
    if (action === 'open-shortcuts') {
      window.alert([
        'GRIDNOVA KEYBOARD SHORTCUTS',
        '',
        'Ctrl / Cmd + C   Copy',
        'Ctrl / Cmd + X   Cut',
        'Ctrl / Cmd + V   Paste',
        'Ctrl / Cmd + Z   Undo',
        'Ctrl + Y         Redo (Windows/Linux)',
        'Cmd + Shift + Z  Redo (Mac)',
        'Ctrl / Cmd + F   Find',
        'Ctrl / Cmd + S   Save workbook',
        '',
        'Some shortcuts depend on which editor control is focused.',
      ].join('\\n'))
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
        <div className="editor-quick-access" aria-label="Quick Access Toolbar">
          <button type="button" onClick={() => handleQuickAccessCommand('univer.command.undo')} title="Undo" aria-label="Undo"><ArrowUndoRegular /></button>
          <button type="button" onClick={() => handleQuickAccessCommand('univer.command.redo')} title="Redo" aria-label="Redo"><ArrowRedoRegular /></button>
        </div>
        <div className="editor-appbar-spacer" />
        <div className="editor-cloud-actions">
          <button type="button" onClick={() => handleTopBarCommand('sheet.operation.toggle-comment-panel', 'comments')}><CommentRegular /> <span>Comments</span></button>
          <button type="button" onClick={() => setReviewPanelMode('activity')}><BookOpenRegular /> <span>Catch up</span></button>
          <button type="button" className="editor-mode-indicator" onClick={handleToggleEditingMode} aria-pressed={isReadOnly} title="Switch between editing and viewing">{isReadOnly ? 'Viewing' : 'Editing'} ▾</button>
          <button className="share-action" type="button" onClick={handleShareWorkbook}><ShareRegular /> <span>Share</span></button>
        </div>
        <div className="ribbon-mode-switch" role="group" aria-label="Ribbon display mode">
          <button type="button" className={ribbonMode === RIBBON_TYPE.SINGLE_LINE ? 'active' : ''} aria-pressed={ribbonMode === RIBBON_TYPE.SINGLE_LINE} onClick={() => handleRibbonModeChange(RIBBON_TYPE.SINGLE_LINE)} title="Single Line Ribbon">
            <TableRegular /> <span>Single Line</span>
          </button>
          <button type="button" className={ribbonMode === RIBBON_TYPE.CLASSIC ? 'active' : ''} aria-pressed={ribbonMode === RIBBON_TYPE.CLASSIC} onClick={() => handleRibbonModeChange(RIBBON_TYPE.CLASSIC)} title="Classic Ribbon">
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
            onReady={handleEditorReady}
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
      {reviewPanelMode && <ReviewPanel workbookId={workbookId} api={editorApi} mode={reviewPanelMode} activity={activity} onClose={() => setReviewPanelMode(null)} />}
    </main>
  )
}
