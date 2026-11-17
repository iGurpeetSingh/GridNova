import { useMemo, useState } from 'react'
import { CommentRegular, HistoryRegular, NoteRegular } from '@fluentui/react-icons'

function readStoredComments(workbookId) {
  try {
    const value = JSON.parse(window.localStorage.getItem('gridnova:comments:' + workbookId) || '[]')
    return Array.isArray(value) ? value : []
  } catch {
    return []
  }
}

export default function ReviewPanel({ workbookId, api, mode = 'comments', activity = [], onClose }) {
  const [tab, setTab] = useState(mode === 'activity' ? 'activity' : 'comments')
  const [comments, setComments] = useState(() => readStoredComments(workbookId))
  const [draft, setDraft] = useState('')
  const [kind, setKind] = useState('Comment')
  const activeCell = useMemo(() => {
    try {
      return api?.getActiveSheet?.()?.getActiveRange?.()?.getA1Notation?.() || 'A1'
    } catch {
      return 'A1'
    }
  }, [api, tab, comments.length])

  function persist(next) {
    setComments(next)
    try {
      window.localStorage.setItem('gridnova:comments:' + workbookId, JSON.stringify(next))
    } catch {
      // Keep the panel usable if browser storage is full or disabled.
    }
  }

  function addComment(event) {
    event.preventDefault()
    const text = draft.trim()
    if (!text) return
    const next = [{
      id: globalThis.crypto?.randomUUID?.() || String(Date.now()),
      cell: activeCell,
      kind,
      text,
      createdAt: Date.now(),
    }, ...comments]
    persist(next)
    setDraft('')
  }

  function deleteComment(id) {
    persist(comments.filter((item) => item.id !== id))
  }

  return (
    <div className="review-panel-backdrop" role="presentation" onMouseDown={(event) => {
      if (event.target === event.currentTarget) onClose?.()
    }}>
      <aside className="review-panel" role="dialog" aria-modal="true" aria-labelledby="review-panel-title">
        <header className="review-panel-header">
          <div>
            <p className="eyebrow">GRIDNOVA REVIEW</p>
            <h2 id="review-panel-title">{tab === 'activity' ? 'Catch up' : 'Comments & Notes'}</h2>
          </div>
          <button type="button" className="panel-close" onClick={onClose} aria-label="Close review panel">×</button>
        </header>

        <nav className="review-panel-tabs" aria-label="Review sections">
          <button type="button" className={tab === 'comments' ? 'selected' : ''} onClick={() => setTab('comments')}>
            <CommentRegular /> Comments
          </button>
          <button type="button" className={tab === 'activity' ? 'selected' : ''} onClick={() => setTab('activity')}>
            <HistoryRegular /> Catch up
          </button>
        </nav>

        {tab === 'comments' ? (
          <>
            <form className="review-comment-form" onSubmit={addComment}>
              <div className="review-comment-context">
                <span>{kind === 'Note' ? <NoteRegular /> : <CommentRegular />}</span>
                <strong>{kind} on {activeCell}</strong>
                <select value={kind} onChange={(event) => setKind(event.target.value)} aria-label="Comment type">
                  <option>Comment</option>
                  <option>Note</option>
                </select>
              </div>
              <textarea value={draft} onChange={(event) => setDraft(event.target.value)} placeholder="Write a comment or note…" rows={3} />
              <button className="panel-primary" type="submit" disabled={!draft.trim()}>Post</button>
              <p className="panel-hint">Comments and notes are stored locally in this browser for this workbook.</p>
            </form>
            <div className="review-comment-list">
              {comments.length ? comments.map((item) => (
                <article className="review-comment-card" key={item.id}>
                  <div className="review-comment-meta">
                    <strong>{item.kind || 'Comment'} · {item.cell}</strong>
                    <time>{new Date(item.createdAt).toLocaleString()}</time>
                  </div>
                  <p>{item.text}</p>
                  <button type="button" onClick={() => deleteComment(item.id)}>Delete</button>
                </article>
              )) : <p className="review-empty">No comments or notes yet. Add one for the currently selected cell.</p>}
            </div>
          </>
        ) : (
          <div className="review-activity-list">
            {activity.length ? activity.map((item) => (
              <article className="review-activity-item" key={item.id}>
                <span className="review-activity-dot" />
                <div>
                  <strong>{item.label || item.command || 'Workbook updated'}</strong>
                  <p>{item.cell ? 'Selection: ' + item.cell : 'Workbook action'}</p>
                  <time>{new Date(item.createdAt).toLocaleString()}</time>
                </div>
              </article>
            )) : <p className="review-empty">No recent workbook activity has been recorded in this browser yet.</p>}
          </div>
        )}
      </aside>
    </div>
  )
}
