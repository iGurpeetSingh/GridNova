import { useEffect, useMemo, useState } from 'react'
import { createPortal } from 'react-dom'
import {
  AccessibilityRegular, AddRegular, ArrowClockwiseRegular, ArrowRightRegular,
  CheckmarkCircleRegular, CommentRegular, DeleteRegular, DismissRegular,
  EditRegular, EyeRegular, HistoryRegular, LockClosedRegular, LockOpenRegular,
  NoteRegular, ShieldCheckmarkRegular, SpellCheckRegular, TableRegular,
  TextColorRegular,
} from '@fluentui/react-icons'
import {
  addCustomDictionaryWord, calculateWorkbookStatistics, getCustomDictionary,
  getIgnoredWords, inspectWorkbookAccessibility, readAltText, saveIgnoredWords, scanSheetForSpelling, writeAltText,
} from '../utils/reviewTools.js'

function localKey(workbookId, suffix) { return 'gridnova:review:' + workbookId + ':' + suffix }
function readLocal(key, fallback) {
  try { const value = JSON.parse(window.localStorage.getItem(key) || 'null'); return value == null ? fallback : value }
  catch { return fallback }
}
function writeLocal(key, value) { try { window.localStorage.setItem(key, JSON.stringify(value)); return true } catch { return false } }
function activeAddress(api) {
  try { return api?.getActiveSheet?.()?.getActiveRange?.()?.getA1Notation?.() || 'A1' } catch { return 'A1' }
}
function commentText(comment) {
  try { const rich = comment?.getRichText?.(); if (rich?.toPlainText) return rich.toPlainText() } catch { /* stale comment */ }
  const data = comment?.getCommentData?.() || {}
  return data.text || data.content?.text || data.content || ''
}
function commentId(comment, fallback) {
  const data = comment?.getCommentData?.() || {}
  return data.id || data.commentId || data.threadId || fallback
}
function mapThread(comment, index) {
  const data = comment?.getCommentData?.() || {}
  const replies = comment?.getReplies?.() || []
  return {
    id: commentId(comment, 'thread-' + index),
    address: comment?.getRange?.()?.getA1Notation?.() || '',
    text: commentText(comment),
    author: data.personId || data.author || data.userName || 'GridNova user',
    createdAt: Number(data.createdAt || data.createTime || data.dateTime || Date.now()),
    resolved: !!(data.resolved || data.isResolved),
    replies: replies.map((reply, replyIndex) => ({
      id: commentId(reply, 'reply-' + index + '-' + replyIndex),
      text: commentText(reply),
      author: reply?.getCommentData?.()?.personId || 'GridNova user',
      createdAt: Number(reply?.getCommentData?.()?.createdAt || Date.now()),
      comment: reply,
    })),
    comment,
  }
}
function makeComment(univerApi, text) {
  const rich = univerApi.newRichText().insertText(text)
  return univerApi.newTheadComment().setContent(rich)
    .setId(globalThis.crypto?.randomUUID?.() || 'comment-' + Date.now())
    .setPersonId('gridnova-local-user').setDateTime(Date.now())
}
function timeLabel(value) {
  return new Date(value || Date.now()).toLocaleString([], { month: 'short', day: 'numeric', hour: 'numeric', minute: '2-digit' })
}
function digestFallback(value) {
  let hash = 2166136261
  for (let index = 0; index < value.length; index++) { hash ^= value.charCodeAt(index); hash = Math.imul(hash, 16777619) }
  return 'fallback-' + (hash >>> 0).toString(16)
}
async function digestPassword(value, salt) {
  const input = salt + ':' + value
  if (globalThis.crypto?.subtle && typeof TextEncoder !== 'undefined') {
    const bytes = await globalThis.crypto.subtle.digest('SHA-256', new TextEncoder().encode(input))
    return [...new Uint8Array(bytes)].map((byte) => byte.toString(16).padStart(2, '0')).join('')
  }
  return digestFallback(input)
}
function randomSalt() {
  const bytes = new Uint8Array(16)
  if (globalThis.crypto?.getRandomValues) globalThis.crypto.getRandomValues(bytes)
  else for (let i = 0; i < bytes.length; i++) bytes[i] = Math.floor(Math.random() * 256)
  return [...bytes].map((byte) => byte.toString(16).padStart(2, '0')).join('')
}
function PanelHeader({ title, subtitle, onClose }) {
  return <header className="review-panel-header"><div className="review-panel-heading"><p className="eyebrow">GRIDNOVA · REVIEW</p><h2>{title}</h2>{subtitle && <p className="review-panel-subtitle">{subtitle}</p>}</div><button type="button" className="panel-close" onMouseDown={(e) => e.preventDefault()} onClick={onClose} aria-label="Close Review pane"><DismissRegular /></button></header>
}
function SectionTitle({ children, count }) { return <div className="review-section-title"><strong>{children}</strong>{count !== undefined && <span>{count}</span>}</div> }

export default function ReviewPanel({
  workbookId, api, mode = 'comments', activity = [], changes = [], onClose,
  onNotice, onClearChanges, onRecordChange, selectedRange, isReadOnly = false,
}) {
  const [panelMode, setPanelMode] = useState(mode)
  const [commentFilter, setCommentFilter] = useState('All')
  const [threads, setThreads] = useState([])
  const [commentDraft, setCommentDraft] = useState('')
  const [replyDrafts, setReplyDrafts] = useState({})
  const [editingId, setEditingId] = useState(null)
  const [editingText, setEditingText] = useState('')
  const [busy, setBusy] = useState(false)
  const [noteText, setNoteText] = useState('')
  const [showAllNotes, setShowAllNotes] = useState(false)
  const [pinnedNotes, setPinnedNotes] = useState(() => readLocal(localKey(workbookId, 'pinned-notes'), []))
  const [spellingIssues, setSpellingIssues] = useState([])
  const [spellingIndex, setSpellingIndex] = useState(0)
  const [ignoredOnce, setIgnoredOnce] = useState([])
  const [replacement, setReplacement] = useState('')
  const [spellComplete, setSpellComplete] = useState(false)
  const [statistics, setStatistics] = useState(null)
  const [accessibilityIssues, setAccessibilityIssues] = useState([])
  const [altDrafts, setAltDrafts] = useState({})
  const [changeFilter, setChangeFilter] = useState('All Changes')
  const [sheetFilter, setSheetFilter] = useState('All')
  const [rangeName, setRangeName] = useState('')
  const [rangeAddress, setRangeAddress] = useState(activeAddress(api))
  const [rangeRules, setRangeRules] = useState(() => readLocal(localKey(workbookId, 'edit-ranges'), []))
  const [protectPassword, setProtectPassword] = useState('')
  const [permissions, setPermissions] = useState({ selectLocked: true, selectUnlocked: true, formatCells: false, insertDeleteRowsColumns: false, sort: true, useAutoFilter: true })
  const [workbookPassword, setWorkbookPassword] = useState('')

  const sheet = api?.getActiveSheet?.()
  const workbook = api?.getActiveWorkbook?.()
  const univerApi = api?.getUniverApi?.()
  const activeCell = activeAddress(api)
  const currentIssue = spellingIssues[spellingIndex] || null

  function notice(message, type = 'success') { onNotice?.(message, type) }
  function refreshThreads() {
    const list = (api?.getActiveSheet?.()?.getComments?.() || [])
      .filter((comment) => comment.getIsRoot?.())
      .map((comment, index) => mapThread(comment, index))
      .sort((a, b) => b.createdAt - a.createdAt)
    setThreads(list)
    return list
  }
  function openComments() { setPanelMode('comments'); refreshThreads() }

  useEffect(() => { setPanelMode(mode) }, [mode])
  useEffect(() => {
    if (panelMode === 'comments') refreshThreads()
    if (panelMode === 'notes') setNoteText(api?.getActiveSheet?.()?.getRange?.(activeCell)?.getNote?.()?.note || '')
    if (panelMode === 'statistics') setStatistics(calculateWorkbookStatistics(api?.getActiveWorkbook?.(), api?.getSnapshot?.()))
    if (panelMode === 'accessibility') setAccessibilityIssues(inspectWorkbookAccessibility(api?.getActiveWorkbook?.(), api?.getSnapshot?.(), workbookId))
    if (panelMode === 'protect-sheet') setRangeAddress(activeCell)
  }, [panelMode, workbookId, api, activeCell])

  const filteredThreads = useMemo(() => threads.filter((item) => {
    if (commentFilter === 'Open') return !item.resolved
    if (commentFilter === 'Resolved') return item.resolved
    return true
  }), [threads, commentFilter])
  const sheetNames = useMemo(() => [...new Set(changes.map((item) => item.sheetName).filter(Boolean))], [changes])
  const visibleChanges = useMemo(() => {
    let result = changes
    if (sheetFilter !== 'All') result = result.filter((item) => item.sheetName === sheetFilter)
    if (changeFilter === 'Selected Range Only' && selectedRange) {
      const firstRow = Math.min(selectedRange.startRow, selectedRange.endRow)
      const lastRow = Math.max(selectedRange.startRow, selectedRange.endRow)
      const firstCol = Math.min(selectedRange.startColumn ?? selectedRange.startCol, selectedRange.endColumn ?? selectedRange.endCol)
      const lastCol = Math.max(selectedRange.startColumn ?? selectedRange.startCol, selectedRange.endColumn ?? selectedRange.endCol)
      result = result.filter((item) => {
        const match = String(item.cell || '').toUpperCase().match(/^([A-Z]+)(\d+)$/)
        if (!match) return false
        let col = 0
        for (const char of match[1]) col = col * 26 + char.charCodeAt(0) - 64
        const row = Number(match[2]) - 1
        return row >= firstRow && row <= lastRow && col - 1 >= firstCol && col - 1 <= lastCol
      })
    }
    return result
  }, [changes, sheetFilter, changeFilter, selectedRange])

  async function addComment(event) {
    event?.preventDefault?.()
    if (isReadOnly) return notice('Switch to Editing mode before posting a comment.', 'error')
    const text = commentDraft.trim()
    if (!text || !sheet || !univerApi) return
    setBusy(true)
    try {
      if (!await sheet.getRange(activeCell).addCommentAsync(makeComment(univerApi, text))) throw new Error('The comment could not be posted.')
      setCommentDraft('')
      refreshThreads()
      onRecordChange?.({ action: 'New comment', sheetName: sheet.getSheetName(), cell: activeCell, description: 'Added comment in ' + activeCell })
    } catch (error) { notice(error instanceof Error ? error.message : 'Could not add comment.', 'error') }
    finally { setBusy(false) }
  }

  async function replyToThread(thread) {
    const text = String(replyDrafts[thread.id] || '').trim()
    if (!text || isReadOnly || !univerApi) return
    setBusy(true)
    try {
      const comment = sheet.getCommentById(thread.id)
      if (!comment || !await comment.replyAsync(makeComment(univerApi, text))) throw new Error('The reply could not be posted.')
      setReplyDrafts((current) => ({ ...current, [thread.id]: '' }))
      refreshThreads()
      onRecordChange?.({ action: 'Reply to comment', sheetName: sheet.getSheetName(), cell: thread.address, description: 'Replied to comment on ' + thread.address })
    } catch (error) { notice(error instanceof Error ? error.message : 'Could not post reply.', 'error') }
    finally { setBusy(false) }
  }

  async function updateComment(item, text) {
    if (isReadOnly || !text.trim()) return
    try {
      const rich = univerApi.newRichText().insertText(text.trim())
      if (!await item.comment.updateAsync(rich)) throw new Error('The comment could not be updated.')
      setEditingId(null); setEditingText(''); refreshThreads()
    } catch (error) { notice(error instanceof Error ? error.message : 'Could not update comment.', 'error') }
  }

  async function deleteComment(item) {
    if (isReadOnly || !window.confirm('Delete this comment or reply?')) return
    try {
      if (!await item.comment.deleteAsync()) throw new Error('The comment could not be deleted.')
      refreshThreads(); notice('Comment deleted.')
    } catch (error) { notice(error instanceof Error ? error.message : 'Could not delete comment.', 'error') }
  }

  async function resolveThread(thread, resolved) {
    if (isReadOnly) return
    try {
      const result = await sheet.getCommentById(thread.id)?.resolveAsync(resolved)
      if (result === false) throw new Error('Thread status was not changed.')
      refreshThreads()
    } catch (error) { notice(error instanceof Error ? error.message : 'Could not update thread status.', 'error') }
  }

  async function deleteThread(thread) {
    if (isReadOnly || !window.confirm('Delete the entire comment thread on ' + thread.address + '?')) return
    try {
      if (!await sheet.getCommentById(thread.id)?.deleteAsync()) throw new Error('The thread could not be deleted.')
      refreshThreads(); notice('Comment thread deleted.')
    } catch (error) { notice(error instanceof Error ? error.message : 'Could not delete thread.', 'error') }
  }

  function refreshSpelling() {
    const savedSheet = api?.getSnapshot?.()?.sheets?.[api?.getActiveSheet?.()?.getSheetId?.()]
    const issues = scanSheetForSpelling(savedSheet, activeAddress(api), { customWords: getCustomDictionary(), ignoredWords: [...getIgnoredWords(workbookId), ...ignoredOnce] })
    setSpellingIssues(issues); setSpellingIndex(0); setSpellComplete(issues.length === 0)
    setReplacement(issues[0]?.suggestions?.[0] || '')
    if (issues[0]) api?.selectCell?.(issues[0].address)
    return issues
  }
  useEffect(() => { if (panelMode === 'spelling') refreshSpelling() }, [panelMode, api, workbookId])
  function advanceSpelling(nextIndex) {
    if (!spellingIssues[nextIndex]) { setSpellComplete(true); return }
    setSpellingIndex(nextIndex)
    setReplacement(spellingIssues[nextIndex].suggestions?.[0] || '')
    api?.selectCell?.(spellingIssues[nextIndex].address)
  }
  function ignoreOnce() {
    if (!currentIssue) return
    setIgnoredOnce((current) => [...current, currentIssue.word.toLowerCase()])
    advanceSpelling(spellingIndex + 1)
  }
  function ignoreAll() {
    if (!currentIssue) return
    saveIgnoredWords(workbookId, [...getIgnoredWords(workbookId), currentIssue.word.toLowerCase()])
    advanceSpelling(spellingIndex + 1)
  }
  function addToDictionary() {
    if (!currentIssue) return
    addCustomDictionaryWord(currentIssue.word)
    refreshSpelling()
  }
  async function changeSpelling(changeAll = false) {
    if (!currentIssue || !replacement.trim() || isReadOnly) return
    const activeSheet = api?.getActiveSheet?.()
    if (!activeSheet) return
    const word = currentIssue.word
    const escaped = word.replace(/[.*+?^$()|[\]\\]/g, '\\$&')
    const pattern = new RegExp(escaped, 'gi')
    try {
      if (changeAll) {
        const cells = api.getSnapshot()?.sheets?.[activeSheet.getSheetId()]?.cellData || {}
        for (const [row, columns] of Object.entries(cells)) {
          for (const [column, cell] of Object.entries(columns || {})) {
            if (!cell || cell.f || cell.formula || typeof cell.v !== 'string') continue
            if (pattern.test(cell.v)) activeSheet.getRange(Number(row), Number(column)).setValue(cell.v.replace(pattern, replacement.trim()))
            pattern.lastIndex = 0
          }
        }
      } else {
        const value = currentIssue.value.slice(0, currentIssue.start) + replacement.trim() + currentIssue.value.slice(currentIssue.end)
        activeSheet.getRange(currentIssue.address).setValue(value)
      }
      onRecordChange?.({ action: changeAll ? 'Change All spelling' : 'Change spelling', sheetName: activeSheet.getSheetName(), cell: currentIssue.address, oldValue: currentIssue.value, newValue: replacement.trim() })
      notice(changeAll ? 'All occurrences of “' + word + '” were changed.' : 'Spelling changed in ' + currentIssue.address + '.')
      refreshSpelling()
    } catch (error) { notice(error instanceof Error ? error.message : 'Could not apply spelling change.', 'error') }
  }

  function saveNote() {
    if (isReadOnly) return notice('Switch to Editing mode before changing notes.', 'error')
    try {
      sheet.getRange(activeCell).createOrUpdateNote({ note: noteText, width: 210, height: 130, show: true })
      notice('Note saved in ' + activeCell + '.')
      onRecordChange?.({ action: 'Save note', sheetName: sheet.getSheetName(), cell: activeCell, description: 'Updated note in ' + activeCell })
    } catch (error) { notice(error instanceof Error ? error.message : 'Could not save note.', 'error') }
  }
  function deleteNote() {
    if (isReadOnly) return notice('Switch to Editing mode before changing notes.', 'error')
    sheet.getRange(activeCell).deleteNote(); setNoteText(''); notice('Note deleted.')
  }
  function toggleCurrentNote() {
    const note = sheet.getRange(activeCell).getNote()
    if (!note) return notice('There is no note in ' + activeCell + '.', 'error')
    sheet.getRange(activeCell).createOrUpdateNote({ note: note.note || '', width: note.width || 210, height: note.height || 130, show: note.show === false })
    notice(note.show === false ? 'Note shown.' : 'Note hidden.')
  }
  function togglePinnedNote() {
    const note = sheet.getRange(activeCell).getNote()
    if (!note) return notice('There is no note in ' + activeCell + '.', 'error')
    const isPinned = pinnedNotes.includes(activeCell)
    const next = isPinned ? pinnedNotes.filter((address) => address !== activeCell) : [...pinnedNotes, activeCell]
    setPinnedNotes(next)
    writeLocal(localKey(workbookId, 'pinned-notes'), next)
    sheet.getRange(activeCell).createOrUpdateNote({ note: note.note || '', width: note.width || 210, height: note.height || 130, show: !isPinned })
    notice(isPinned ? 'Note unpinned.' : 'Note pinned.')
  }
  function toggleAllNotes() {
    const notes = sheet.getNotes?.() || []
    const nextShow = !showAllNotes
    notes.forEach((note) => sheet.getRange(note.row, note.col).createOrUpdateNote({ note: note.note || '', width: note.width || 210, height: note.height || 130, show: nextShow }))
    setShowAllNotes(nextShow); notice(nextShow ? 'All notes shown.' : 'All notes hidden.')
  }
  async function convertNotesToComments() {
    if (isReadOnly || !univerApi) return
    const notes = sheet.getNotes?.() || []
    if (!notes.length) return notice('There are no notes to convert.')
    if (!window.confirm('Convert ' + notes.length + ' note(s) into threaded comments? The original notes will be removed.')) return
    setBusy(true)
    try {
      for (const item of notes) {
        await sheet.getRange(item.row, item.col).addCommentAsync(makeComment(univerApi, item.note || ''))
        sheet.getRange(item.row, item.col).deleteNote()
        onRecordChange?.({ action: 'Convert note to comment', sheetName: sheet.getSheetName(), cell: sheet.getRange(item.row, item.col).getA1Notation(), description: 'Converted note to comment' })
      }
      refreshThreads(); notice('Converted ' + notes.length + ' note(s) to comments.')
    } catch (error) { notice(error instanceof Error ? error.message : 'Could not convert notes.', 'error') }
    finally { setBusy(false) }
  }

  function recheckAccessibility() { setAccessibilityIssues(inspectWorkbookAccessibility(api?.getActiveWorkbook?.(), api?.getSnapshot?.(), workbookId)) }
  async function addAltText(issue) {
    const text = String(altDrafts[issue.id] || '').trim()
    if (!text) return
    writeAltText(workbookId, issue.objectId, text)
    setAltDrafts((current) => ({ ...current, [issue.id]: '' }))
    recheckAccessibility(); notice('Alt Text saved.')
  }
  function fixContrast(issue) {
    if (isReadOnly) return notice('Switch to Editing mode before formatting cells.', 'error')
    try {
      const range = sheet?.getRange?.(issue.address)
      if (!range?.setFontColor) throw new Error('Font-color editing is not available for this cell.')
      range.setFontColor('#1F2937'); recheckAccessibility(); notice('Text contrast improved in ' + issue.address + '.')
    } catch (error) { notice(error instanceof Error ? error.message : 'Could not fix text contrast.', 'error') }
  }
  function renameDefaultSheet(issue) {
    if (isReadOnly) return notice('Switch to Editing mode before renaming a worksheet.', 'error')
    const name = window.prompt('Rename worksheet', 'Data')
    if (!name?.trim()) return
    try { workbook.getSheets().find((item) => item.getSheetId() === issue.sheetId)?.setName(name.trim()); notice('Worksheet renamed to ' + name.trim() + '.') }
    catch (error) { notice(error instanceof Error ? error.message : 'Could not rename worksheet.', 'error') }
  }
  function unmergeIssue(issue) {
    if (isReadOnly) return notice('Switch to Editing mode before unmerging cells.', 'error')
    api?.selectCell?.(issue.address, issue.sheetId)
    api?.executeCommand?.('sheet.command.remove-worksheet-merge')
    notice('Merged range ' + issue.address + ' unmerged.')
  }

  async function protectSheet() {
    if (isReadOnly) return notice('Switch to Editing mode before changing protection.', 'error')
    const permission = sheet?.getWorksheetPermission?.()
    if (!permission) return notice('Worksheet protection is not available in this editor.', 'error')
    setBusy(true)
    try {
      if (!permission.isProtected()) await permission.protect({ name: 'GridNova worksheet protection' })
      const points = univerApi?.Enum?.WorksheetPermissionPoint || {}
      const rules = [['selectLocked', 'SelectProtectedCells'], ['selectUnlocked', 'SelectUnProtectedCells'], ['formatCells', 'SetCellStyle'], ['sort', 'Sort'], ['useAutoFilter', 'Filter']]
      for (const [field, key] of rules) if (points[key] !== undefined) await permission.setPoint(points[key], !!permissions[field])
      const rowKeys = ['InsertRow', 'DeleteRow', 'InsertColumn', 'DeleteColumn']
      for (const key of rowKeys) if (points[key] !== undefined) await permission.setPoint(points[key], !!permissions.insertDeleteRowsColumns)
      const salt = randomSalt()
      const hash = protectPassword ? await digestPassword(protectPassword, salt) : ''
      writeLocal(localKey(workbookId, 'sheet-password:' + sheet.getSheetId()), { salt, hash })
      setProtectPassword('')
      await permission.setMode('editable')
      notice('Sheet protection is enabled.')
      onRecordChange?.({ action: 'Protect sheet', sheetName: sheet.getSheetName(), description: 'Enabled worksheet protection' })
      api?.updateReviewCommandState?.('protect-sheet', 'Unprotect Sheet')
      onClose?.()
    } catch (error) { notice(error instanceof Error ? error.message : 'Could not protect this sheet.', 'error') }
    finally { setBusy(false) }
  }
  async function unprotectSheet() {
    if (isReadOnly) return notice('Switch to Editing mode before changing protection.', 'error')
    const permission = sheet?.getWorksheetPermission?.()
    if (!permission?.isProtected?.()) return notice('This sheet is already unprotected.')
    try {
      const saved = readLocal(localKey(workbookId, 'sheet-password:' + sheet.getSheetId()), null)
      if (saved?.hash) {
        const entered = window.prompt('Enter the sheet protection passcode')
        if (entered === null) return
        if (await digestPassword(entered, saved.salt) !== saved.hash) return notice('Incorrect passcode. The sheet is still protected.', 'error')
      }
      await permission.unprotect()
      writeLocal(localKey(workbookId, 'sheet-password:' + sheet.getSheetId()), null)
      await permission.setEditable()
      notice('Sheet protection removed.')
      onRecordChange?.({ action: 'Unprotect sheet', sheetName: sheet.getSheetName(), description: 'Removed worksheet protection' })
      api?.updateReviewCommandState?.('protect-sheet', 'Protect Sheet')
      onClose?.()
    } catch (error) { notice(error instanceof Error ? error.message : 'Could not unprotect this sheet.', 'error') }
  }

  async function createAllowEditRange() {
    const name = rangeName.trim()
    const address = rangeAddress.trim().toUpperCase()
    if (!name || !address) return notice('Enter a range name and cell range.', 'error')
    try {
      const permission = sheet.getRange(address).getRangePermission?.()
      if (!permission?.protect) throw new Error('Range protection is not available in this editor.')
      await permission.protect({ name, allowEdit: true })
      const next = [...rangeRules.filter((item) => item.sheetId !== sheet.getSheetId() || item.name !== name), { name, address, sheetId: sheet.getSheetId(), sheetName: sheet.getSheetName() }]
      setRangeRules(next); writeLocal(localKey(workbookId, 'edit-ranges'), next); setRangeName('')
      notice('Allow Edit Range “' + name + '” created for ' + address + '.')
      onRecordChange?.({ action: 'Allow Edit Range', sheetName: sheet.getSheetName(), cell: address, description: 'Created editable range ' + name })
    } catch (error) { notice(error instanceof Error ? error.message : 'Could not create the editable range.', 'error') }
  }
  async function removeAllowEditRange(rule) {
    try {
      const target = workbook.getSheets().find((item) => item.getSheetId() === rule.sheetId)
      await target?.getRange(rule.address)?.getRangePermission?.()?.unprotect?.()
      const next = rangeRules.filter((item) => item !== rule)
      setRangeRules(next); writeLocal(localKey(workbookId, 'edit-ranges'), next)
    } catch (error) { notice(error instanceof Error ? error.message : 'Could not remove the editable range.', 'error') }
  }

  async function protectWorkbookStructure() {
    if (isReadOnly) return notice('Switch to Editing mode before changing workbook protection.', 'error')
    const permission = workbook?.getWorkbookPermission?.()
    if (!permission) return notice('Workbook permissions are not available in this editor.', 'error')
    setBusy(true)
    try {
      const points = univerApi?.Enum?.WorkbookPermissionPoint || {}
      const salt = randomSalt()
      const hash = workbookPassword ? await digestPassword(workbookPassword, salt) : ''
      writeLocal(localKey(workbookId, 'workbook-password'), { salt, hash })
      for (const key of ['CreateSheet', 'DeleteSheet', 'RenameSheet', 'MoveSheet', 'CopySheet']) {
        if (points[key] !== undefined) await permission.setPoint(points[key], false)
      }
      setWorkbookPassword('')
      notice('Workbook structure is protected. Worksheet tabs cannot be added, deleted, renamed, moved or copied.')
      onRecordChange?.({ action: 'Protect workbook', description: 'Protected workbook structure' })
      api?.updateReviewCommandState?.('protect-workbook', 'Unprotect Workbook')
      onClose?.()
    } catch (error) { notice(error instanceof Error ? error.message : 'Could not protect workbook structure.', 'error') }
    finally { setBusy(false) }
  }
  async function unprotectWorkbookStructure() {
    if (isReadOnly) return notice('Switch to Editing mode before changing workbook protection.', 'error')
    try {
      const saved = readLocal(localKey(workbookId, 'workbook-password'), null)
      if (saved?.hash) {
        const entered = window.prompt('Enter the workbook protection passcode')
        if (entered === null) return
        if (await digestPassword(entered, saved.salt) !== saved.hash) return notice('Incorrect passcode. Workbook structure remains protected.', 'error')
      }
      const permission = workbook?.getWorkbookPermission?.()
      const points = univerApi?.Enum?.WorkbookPermissionPoint || {}
      for (const key of ['CreateSheet', 'DeleteSheet', 'RenameSheet', 'MoveSheet', 'CopySheet']) if (points[key] !== undefined) await permission?.setPoint?.(points[key], true)
      writeLocal(localKey(workbookId, 'workbook-password'), null)
      notice('Workbook structure protection removed.')
      onRecordChange?.({ action: 'Unprotect workbook', description: 'Removed workbook structure protection' })
      api?.updateReviewCommandState?.('protect-workbook', 'Protect Workbook')
      onClose?.()
    } catch (error) { notice(error instanceof Error ? error.message : 'Could not unprotect workbook.', 'error') }
  }

  async function revertChange(item) {
    if (isReadOnly || !item.cell) return
    try {
      const target = workbook?.getSheets?.().find((candidate) => candidate.getSheetId?.() === item.sheetId)
      if (!target) throw new Error('The worksheet for this change is no longer available.')
      target.getRange(item.cell).setValue(item.oldValue === '(blank)' ? '' : item.oldValue)
      notice('Reverted ' + item.sheetName + ' · ' + item.cell + ' to its previous value.')
    } catch (error) { notice(error instanceof Error ? error.message : 'Could not revert this change.', 'error') }
  }
  function jumpToCell(address, targetSheetId) {
    if (!address) return
    const target = workbook?.getSheets?.().find((item) => item.getSheetId?.() === targetSheetId) || sheet
    target?.activate?.()
    api?.selectCell?.(address, targetSheetId)
  }

  const panelTitles = {
    comments: 'Comments', notes: 'Notes', spelling: 'Spelling', statistics: 'Workbook Statistics',
    accessibility: 'Accessibility', changes: 'Show Changes',
    'protect-sheet': sheet?.getWorksheetPermission?.()?.isProtected?.() ? 'Unprotect Sheet' : 'Protect Sheet',
    'allow-ranges': 'Allow Edit Ranges', 'protect-workbook': 'Protect Workbook',
  }

  if (typeof document === 'undefined') return null

  return createPortal(
    <div className="review-panel-backdrop" role="presentation" onMouseDown={(event) => { if (event.target === event.currentTarget) onClose?.() }}>
      <aside className={'review-panel ' + (panelMode === 'statistics' ? 'review-panel-wide' : '')} role="dialog" aria-modal="true" aria-labelledby="review-panel-title">
        <PanelHeader title={panelTitles[panelMode] || 'Review'} subtitle={sheet?.getSheetName?.() || 'Active worksheet'} onClose={onClose} />

        {panelMode === 'comments' && <>
          <div className="review-panel-toolbar"><div className="review-filter-segment">{['All', 'Open', 'Resolved'].map((item) => <button key={item} type="button" onMouseDown={(e) => e.preventDefault()} className={commentFilter === item ? 'selected' : ''} onClick={() => setCommentFilter(item)}>{item}</button>)}</div><button type="button" className="review-quiet-button" onMouseDown={(e) => e.preventDefault()} onClick={refreshThreads}>Refresh</button></div>
          <form className="review-comment-form" onSubmit={addComment}>
            <div className="review-comment-context"><CommentRegular /><strong>New comment · {activeCell}</strong></div>
            <textarea value={commentDraft} onChange={(event) => setCommentDraft(event.target.value)} placeholder="Write a comment…" rows={3} disabled={isReadOnly} onKeyDown={(event) => { if ((event.ctrlKey || event.metaKey) && event.key === 'Enter') addComment(event) }} />
            <div className="review-form-footer"><span>Ctrl + Enter to post</span><button className="panel-primary" type="submit" disabled={busy || isReadOnly || !commentDraft.trim()}>Post</button></div>
          </form>
          <div className="review-comment-list">
            {filteredThreads.length ? filteredThreads.map((thread) => <article className={'review-thread-card ' + (thread.resolved ? 'resolved' : '')} key={thread.id}>
              <div className="review-thread-top"><button type="button" className="review-cell-link" onClick={() => jumpToCell(thread.address, sheet?.getSheetId?.())}>{thread.address || 'Cell'}</button><time>{timeLabel(thread.createdAt)}</time></div>
              <div className="review-thread-message"><span className="review-avatar">{String(thread.author || 'G').slice(0, 1).toUpperCase()}</span><div className="review-thread-body"><div className="review-thread-author">{thread.author}</div>{editingId === thread.id ? <div className="review-inline-edit"><textarea value={editingText} onChange={(e) => setEditingText(e.target.value)} rows={3} /><div><button type="button" onClick={() => updateComment(thread, editingText)}>Save</button><button type="button" onClick={() => setEditingId(null)}>Cancel</button></div></div> : <p>{thread.text}</p>}</div></div>
              {thread.resolved && <span className="review-resolved-badge">✓ Resolved</span>}
              {thread.replies.map((reply) => <div className="review-reply" key={reply.id}><div className="review-reply-meta"><strong>{reply.author}</strong><time>{timeLabel(reply.createdAt)}</time></div>{editingId === reply.id ? <div className="review-inline-edit"><textarea value={editingText} onChange={(e) => setEditingText(e.target.value)} rows={2} /><div><button type="button" onClick={() => updateComment(reply, editingText)}>Save</button><button type="button" onClick={() => setEditingId(null)}>Cancel</button></div></div> : <p>{reply.text}</p>}{!isReadOnly && <div className="review-message-actions"><button type="button" onClick={() => { setEditingId(reply.id); setEditingText(reply.text) }}>Edit</button><button type="button" onClick={() => deleteComment(reply)}>Delete</button></div>}</div>)}
              <div className="review-thread-actions"><button type="button" onClick={() => setReplyDrafts((current) => ({ ...current, [thread.id]: current[thread.id] === undefined ? '' : current[thread.id] }))}>Reply</button>{!isReadOnly && <><button type="button" onClick={() => { setEditingId(thread.id); setEditingText(thread.text) }}>Edit</button><button type="button" onClick={() => resolveThread(thread, !thread.resolved)}>{thread.resolved ? 'Reopen' : 'Resolve thread'}</button><button type="button" className="danger" onClick={() => deleteThread(thread)}>Delete</button></>}</div>
              {replyDrafts[thread.id] !== undefined && <form className="review-reply-form" onSubmit={(e) => { e.preventDefault(); replyToThread(thread) }}><textarea value={replyDrafts[thread.id]} onChange={(e) => setReplyDrafts((current) => ({ ...current, [thread.id]: e.target.value }))} placeholder="Write a reply…" rows={2} /><button className="panel-primary small" type="submit" disabled={busy || isReadOnly || !String(replyDrafts[thread.id] || '').trim()}>Reply</button></form>}
            </article>) : <p className="review-empty">No {commentFilter.toLowerCase()} comments on this worksheet. Add one above.</p>}
          </div>
          <div className="review-pane-footer"><span>{threads.length} thread(s)</span><button type="button" className="review-quiet-button" onClick={() => { setCommentDraft(''); document.querySelector('.review-comment-form textarea')?.focus() }}>+ New</button></div>
        </>}

        {panelMode === 'notes' && <div className="review-pane-scroll">
          <div className="review-note-editor"><div className="review-note-address"><NoteRegular /> Note for <strong>{activeCell}</strong></div><textarea value={noteText} onChange={(e) => setNoteText(e.target.value)} placeholder="Type a classic yellow sticky note…" rows={5} disabled={isReadOnly} /><div className="review-note-actions"><button type="button" className="panel-primary" disabled={isReadOnly} onClick={saveNote}>Save Note</button><button type="button" className="panel-secondary" disabled={isReadOnly} onClick={deleteNote}>Delete Note</button></div></div>
          <SectionTitle count={(sheet?.getNotes?.() || []).length}>Notes on this sheet</SectionTitle>
          <div className="review-note-command-list">
            <button type="button" onClick={() => setNoteText(sheet.getRange(activeCell).getNote()?.note || '')}><EditRegular /> New Note / Edit Note <span>{activeCell}</span></button>
            <button type="button" disabled={isReadOnly} onClick={deleteNote}><DeleteRegular /> Delete Note</button>
            <button type="button" onClick={toggleCurrentNote}><EyeRegular /> Show / Hide Note</button>
            <button type="button" onClick={togglePinnedNote}><NoteRegular /> {pinnedNotes.includes(activeCell) ? 'Unpin Note' : 'Pin Note'}</button>
            <button type="button" onClick={toggleAllNotes}><NoteRegular /> {showAllNotes ? 'Hide All Notes' : 'Show All Notes'}</button>
            <button type="button" disabled={isReadOnly || !(sheet?.getNotes?.() || []).length} onClick={convertNotesToComments}><CommentRegular /> Convert to Comments</button>
          </div>
          {(sheet?.getNotes?.() || []).map((item) => <button type="button" className="review-note-list-item" key={item.row + ':' + item.col} onClick={() => { const address = sheet.getRange(item.row, item.col).getA1Notation(); setNoteText(item.note || ''); api?.selectCell?.(address) }}><span className="review-note-dot" /><strong>{sheet.getRange(item.row, item.col).getA1Notation()}</strong><span>{item.note || 'Empty note'}</span></button>)}
        </div>}

        {panelMode === 'spelling' && <div className="review-pane-scroll">
          {spellComplete || !currentIssue ? <div className="review-success-state"><CheckmarkCircleRegular /><h3>Spell check complete.</h3><p>You're good to go!</p><button className="panel-primary" type="button" onClick={onClose}>Done</button></div> : <>
            <div className="review-spelling-progress"><span>Issue {spellingIndex + 1} of {spellingIssues.length}</span><span>{currentIssue.address}</span></div><SectionTitle>Not in dictionary</SectionTitle><div className="review-misspelled-word">{currentIssue.word}</div><p className="review-helper-copy">Cell content: <strong>{currentIssue.value}</strong></p><SectionTitle>Suggestions</SectionTitle>
            <div className="review-suggestion-list">{(currentIssue.suggestions || []).length ? currentIssue.suggestions.map((value) => <button type="button" key={value} className={replacement === value ? 'selected' : ''} onClick={() => setReplacement(value)}>{value}</button>) : <span className="review-empty-inline">No close dictionary match. Enter a correction below.</span>}</div>
            <label className="review-field-label">Replace with<input value={replacement} onChange={(e) => setReplacement(e.target.value)} /></label>
            <div className="review-spelling-actions"><button type="button" className="panel-primary" disabled={isReadOnly || !replacement.trim()} onClick={() => changeSpelling(false)}>Change</button><button type="button" className="panel-secondary" disabled={isReadOnly || !replacement.trim()} onClick={() => changeSpelling(true)}>Change All</button><button type="button" className="panel-secondary" onClick={ignoreOnce}>Ignore Once</button><button type="button" className="panel-secondary" onClick={ignoreAll}>Ignore All</button><button type="button" className="panel-secondary" onClick={addToDictionary}>Add to Dictionary</button></div>
            <div className="review-pane-footer"><span>{getCustomDictionary().length} custom words saved</span><button type="button" className="review-quiet-button" onClick={refreshSpelling}>Restart</button></div>
          </>}
        </div>}

        {panelMode === 'statistics' && <div className="review-pane-scroll review-statistics">
          <div className="review-stats-block"><SectionTitle>Current Sheet</SectionTitle>{statistics && Object.entries({ 'End of sheet': statistics.current.endOfSheet, 'Cells with data': statistics.current.cellsWithData, Tables: statistics.current.tables, Formulas: statistics.current.formulas, 'Charts / Objects': statistics.current.chartsAndObjects, Comments: statistics.current.comments, Notes: statistics.current.notes }).map(([label, value]) => <div className="review-stat-row" key={label}><span>{label}</span><strong>{value}</strong></div>)}</div>
          <div className="review-stats-block"><SectionTitle>Workbook</SectionTitle>{statistics && Object.entries({ 'Total sheets': statistics.workbook.totalSheets, 'Total cells with data': statistics.workbook.totalCellsWithData, 'Total formulas': statistics.workbook.totalFormulas, 'Total comments & notes': statistics.workbook.totalCommentsAndNotes, 'Protected sheets': statistics.workbook.protectedSheets }).map(([label, value]) => <div className="review-stat-row" key={label}><span>{label}</span><strong>{value}</strong></div>)}</div>
          <button type="button" className="panel-secondary" onClick={() => setStatistics(calculateWorkbookStatistics(api?.getActiveWorkbook?.(), api?.getSnapshot?.()))}>Refresh Statistics</button>
        </div>}

        {panelMode === 'accessibility' && <div className="review-pane-scroll">
          <div className="review-accessibility-summary"><AccessibilityRegular /><div><strong>{accessibilityIssues.length ? accessibilityIssues.length + ' issue(s) found' : 'No accessibility issues found'}</strong><span>Workbook scan · WCAG AA 4.5:1 target</span></div><button type="button" onClick={recheckAccessibility} aria-label="Re-check"><SearchRegular /></button></div>
          {accessibilityIssues.length ? accessibilityIssues.map((issue) => <article className="review-issue-card" key={issue.id}>
            <div className="review-issue-header"><span>{issue.category}</span><button type="button" onClick={() => jumpToCell(issue.address, issue.sheetId)}>{issue.address || 'Go to item'} <ArrowRightRegular /></button></div><strong>{issue.title}</strong><p>{issue.description}</p>
            {issue.action === 'alt-text' && <div className="review-inline-field"><input value={altDrafts[issue.id] || ''} onChange={(e) => setAltDrafts((current) => ({ ...current, [issue.id]: e.target.value }))} placeholder="Describe this object…" /><button type="button" disabled={!String(altDrafts[issue.id] || '').trim()} onClick={() => addAltText(issue)}>Save Alt Text</button></div>}
            {issue.action === 'fix-contrast' && <button type="button" className="review-issue-action" disabled={isReadOnly} onClick={() => fixContrast(issue)}><TextColorRegular /> Fix Contrast</button>}
            {issue.action === 'rename-sheet' && <button type="button" className="review-issue-action" disabled={isReadOnly} onClick={() => renameDefaultSheet(issue)}><EditRegular /> Rename</button>}
            {issue.action === 'unmerge' && <button type="button" className="review-issue-action" disabled={isReadOnly} onClick={() => unmergeIssue(issue)}><TableRegular /> Unmerge</button>}
          </article>) : <div className="review-success-inline"><CheckmarkCircleRegular /> No accessibility issues found</div>}
        </div>}

        {panelMode === 'changes' && <>
          <div className="review-changes-filters"><select value={changeFilter} onChange={(e) => setChangeFilter(e.target.value)} aria-label="Change filter"><option>All Changes</option><option>Selected Range Only</option></select><select value={sheetFilter} onChange={(e) => setSheetFilter(e.target.value)} aria-label="Sheet filter"><option>All</option>{sheetNames.map((name) => <option key={name}>{name}</option>)}</select></div>
          <div className="review-change-list">{visibleChanges.length ? visibleChanges.map((item) => <article className="review-change-card" key={item.id}><div className="review-change-meta"><strong>{item.sheetName}{item.cell ? ' · ' + item.cell : ''}</strong><time>{timeLabel(item.createdAt)}</time></div><p>{item.description || item.action}</p>{item.cell && <div className="review-change-values"><span>{item.oldValue || '(blank)'}</span><ArrowRightRegular /><strong>{item.newValue || '(blank)'}</strong></div>}<div className="review-change-actions"><button type="button" disabled={!item.cell} onClick={() => jumpToCell(item.cell, item.sheetId)}>Go to cell</button>{item.cell && <button type="button" disabled={isReadOnly} onClick={() => revertChange(item)}>Revert this cell</button>}</div></article>) : <p className="review-empty">No changes match the selected filters.</p>}</div>
          <div className="review-pane-footer"><span>{visibleChanges.length} change(s)</span><button type="button" className="review-quiet-button danger" onClick={() => { onClearChanges?.(); notice('Change log cleared.') }}>Clear Log</button></div>
        </>}

        {panelMode === 'protect-sheet' && <div className="review-pane-scroll">
          <p className="review-protect-copy">Protect the active sheet and choose which actions remain available.</p>
          <label className="review-field-label">Optional passcode<input type="password" value={protectPassword} onChange={(e) => setProtectPassword(e.target.value)} autoComplete="new-password" placeholder="Leave blank for no passcode" /></label>
          <SectionTitle>Allow users of this sheet to</SectionTitle>
          <div className="review-permission-list">{[['selectLocked', 'Select locked cells'], ['selectUnlocked', 'Select unlocked cells'], ['formatCells', 'Format cells'], ['insertDeleteRowsColumns', 'Insert / Delete rows & columns'], ['sort', 'Sort'], ['useAutoFilter', 'Use AutoFilter']].map(([key, label]) => <label key={key}><input type="checkbox" checked={!!permissions[key]} onChange={(e) => setPermissions((current) => ({ ...current, [key]: e.target.checked }))} />{label}</label>)}</div>
          {sheet?.getWorksheetPermission?.()?.isProtected?.() ? <button type="button" className="panel-primary review-full-button" disabled={busy || isReadOnly} onClick={unprotectSheet}><LockOpenRegular /> Unprotect Sheet</button> : <button type="button" className="panel-primary review-full-button" disabled={busy || isReadOnly} onClick={protectSheet}><LockClosedRegular /> Protect Sheet</button>}
        </div>}

        {panelMode === 'allow-ranges' && <div className="review-pane-scroll">
          <p className="review-protect-copy">Named ranges remain editable when the worksheet is protected.</p>
          <label className="review-field-label">Range name<input value={rangeName} onChange={(e) => setRangeName(e.target.value)} placeholder="InputRange" /></label>
          <label className="review-field-label">Cell range<input value={rangeAddress} onChange={(e) => setRangeAddress(e.target.value.toUpperCase())} placeholder="B2:D15" /></label>
          <button type="button" className="panel-primary review-full-button" onClick={createAllowEditRange} disabled={isReadOnly || !rangeName.trim() || !rangeAddress.trim()}><LockOpenRegular /> Allow Edit Range</button>
          <SectionTitle count={rangeRules.length}>Named edit ranges</SectionTitle>
          {rangeRules.map((item, index) => <div className="review-range-rule" key={item.sheetId + ':' + item.name + ':' + index}><div><strong>{item.name}</strong><span>{item.sheetName} · {item.address}</span></div><button type="button" onClick={() => removeAllowEditRange(item)}>Remove</button></div>)}
        </div>}

        {panelMode === 'protect-workbook' && <div className="review-pane-scroll">
          <p className="review-protect-copy">Workbook structure protection prevents adding, deleting, renaming, moving, and duplicating worksheets. Cell editing remains available unless a sheet is separately protected.</p>
          <label className="review-field-label">Optional passcode<input type="password" value={workbookPassword} onChange={(e) => setWorkbookPassword(e.target.value)} autoComplete="new-password" placeholder="Leave blank for no passcode" /></label>
          <div className="review-permission-list"><label><input type="checkbox" checked disabled /> Prevent adding sheets</label><label><input type="checkbox" checked disabled /> Prevent deleting sheets</label><label><input type="checkbox" checked disabled /> Prevent renaming sheets</label><label><input type="checkbox" checked disabled /> Prevent moving / duplicating sheets</label></div>
          <button type="button" className="panel-primary review-full-button" disabled={busy || isReadOnly} onClick={protectWorkbookStructure}><LockClosedRegular /> Protect Workbook</button>
          <button type="button" className="panel-secondary review-full-button" disabled={busy || isReadOnly} onClick={unprotectWorkbookStructure}><LockOpenRegular /> Unprotect Workbook</button>
        </div>}
      </aside>
    </div>,
    document.body,
  )
}
