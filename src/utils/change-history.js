import { diffWorkbookSnapshots, readChangeHistory, writeChangeHistory } from './reviewTools.js'

export function recordWorkbookChanges(workbookId, previousSnapshot, nextSnapshot, commandId) {
  const changes = diffWorkbookSnapshots(previousSnapshot, nextSnapshot, commandId)
  if (!changes.length) return []
  const current = readChangeHistory(workbookId)
  const next = [...changes, ...current].slice(0, 500)
  writeChangeHistory(workbookId, next)
  return changes
}

export function loadWorkbookChanges(workbookId) {
  return readChangeHistory(workbookId)
}

export function clearWorkbookChanges(workbookId) {
  writeChangeHistory(workbookId, [])
}

export function createReviewChange(workbookId, fields) {
  const entry = {
    id: globalThis.crypto?.randomUUID?.() || String(Date.now()),
    sheetId: fields.sheetId || '',
    sheetName: fields.sheetName || 'Worksheet',
    cell: fields.cell || '',
    action: fields.action || 'Workbook change',
    oldValue: fields.oldValue ?? '',
    newValue: fields.newValue ?? '',
    description: fields.description || fields.action || 'Workbook updated',
    createdAt: Date.now(),
  }
  const current = readChangeHistory(workbookId)
  writeChangeHistory(workbookId, [entry, ...current])
  return entry
}
