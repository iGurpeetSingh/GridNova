import { useMemo, useState } from 'react'

function makeChartModel(matrix) {
  if (!Array.isArray(matrix) || matrix.length < 2) {
    throw new Error('Select a range with a header row and at least one data row, for example A1:B8.')
  }
  const dataRows = matrix.slice(1).filter((row) => row?.length)
  const labels = dataRows.map((row, index) => String(row?.[0] ?? 'Row ' + (index + 1)))
  const values = dataRows.map((row) => {
    const value = Number(row?.[1])
    return Number.isFinite(value) ? value : 0
  })
  if (!values.some((value) => value !== 0)) {
    throw new Error('The second column must contain numeric values to chart.')
  }
  return { title: String(matrix[0]?.[1] || 'Series'), labels, values }
}

function SvgChart({ model, type }) {
  const width = 560
  const height = 280
  const left = 40
  const right = 20
  const top = 24
  const bottom = 54
  const chartWidth = width - left - right
  const chartHeight = height - top - bottom
  const maximum = Math.max(1, ...model.values.map((value) => Math.max(0, value))) * 1.12
  const minimum = Math.min(0, ...model.values)
  const range = Math.max(1, maximum - minimum)
  const y = (value) => top + chartHeight - ((value - minimum) / range) * chartHeight
  const barWidth = Math.max(8, chartWidth / model.values.length * 0.58)
  const step = chartWidth / model.values.length
  const points = model.values.map((value, index) => (left + step * (index + 0.5)) + ',' + y(value)).join(' ')

  return (
    <svg className="data-chart" viewBox="0 0 560 280" role="img" aria-label={type + ' chart preview for ' + model.title}>
      {[0, 1, 2, 3].map((tick) => {
        const value = minimum + range * tick / 3
        const yPosition = y(value)
        return (
          <g key={tick}>
            <line x1={left} y1={yPosition} x2={width - right} y2={yPosition} stroke="#e9edf0" strokeDasharray="3 4" />
            <text x={left - 8} y={yPosition + 4} textAnchor="end" fontSize="9" fill="#8793a0">{Math.round(value * 100) / 100}</text>
          </g>
        )
      })}
      {type === 'line' ? (
        <>
          <polyline points={points} fill="none" stroke="#16834b" strokeWidth="3" strokeLinejoin="round" strokeLinecap="round" />
          {model.values.map((value, index) => (
            <circle key={model.labels[index]} cx={left + step * (index + 0.5)} cy={y(value)} r="4" fill="#fff" stroke="#16834b" strokeWidth="2.5" />
          ))}
        </>
      ) : (
        model.values.map((value, index) => {
          const barHeight = Math.abs(y(0) - y(value))
          const x = left + step * index + (step - barWidth) / 2
          return <rect key={model.labels[index]} x={x} y={Math.min(y(0), y(value))} width={barWidth} height={Math.max(1, barHeight)} rx="4" fill="#16834b" opacity={0.84 - (index % 4) * 0.08} />
        })
      )}
      {model.labels.map((label, index) => (
        <text key={label + index} x={left + step * (index + 0.5)} y={height - 27} textAnchor="middle" fontSize="9" fill="#778392">
          {label.length > 12 ? label.slice(0, 11) + '…' : label}
        </text>
      ))}
    </svg>
  )
}

export default function InsertToolsPanel({ api, onClose }) {
  const [activeTool, setActiveTool] = useState('chart')
  const [range, setRange] = useState('A1:B8')
  const [chartType, setChartType] = useState('column')
  const [chartModel, setChartModel] = useState(null)
  const [headers, setHeaders] = useState([])
  const [rowField, setRowField] = useState('')
  const [valueField, setValueField] = useState('')
  const [aggregation, setAggregation] = useState('SUM')
  const [message, setMessage] = useState('')
  const [error, setError] = useState('')

  const hasWorkbook = Boolean(api?.getActiveWorkbook?.())
  const cleanHeaders = useMemo(() => headers.map((header) => String(header)), [headers])

  function readMatrix() {
    if (!hasWorkbook) throw new Error('The workbook is not ready yet.')
    const sheet = api.getActiveSheet()
    if (!sheet) throw new Error('No active worksheet is selected.')
    return sheet.getRange(range.trim()).getValues()
  }

  function previewChart() {
    setError('')
    setMessage('')
    try {
      setChartModel(makeChartModel(readMatrix()))
    } catch (reason) {
      setChartModel(null)
      setError(reason instanceof Error ? reason.message : 'Could not read the selected range.')
    }
  }

  function inspectPivotRange() {
    setError('')
    setMessage('')
    try {
      const values = readMatrix()
      if (values.length < 2 || values[0].length < 2) {
        throw new Error('Pivot source needs a header row and at least two columns.')
      }
      const nextHeaders = values[0].map((header, index) => String(header || 'Column ' + (index + 1)))
      setHeaders(nextHeaders)
      setRowField(nextHeaders[0])
      setValueField(nextHeaders[1])
    } catch (reason) {
      setError(reason instanceof Error ? reason.message : 'Could not inspect the source data.')
    }
  }

  function createPivotSummary() {
    setError('')
    setMessage('')
    try {
      const matrix = readMatrix()
      if (matrix.length < 2) throw new Error('Add a header row and at least one record first.')
      const names = matrix[0].map((header, index) => String(header || 'Column ' + (index + 1)))
      const groupIndex = names.indexOf(rowField)
      const valueIndex = names.indexOf(valueField)
      if (groupIndex < 0 || valueIndex < 0) throw new Error('Choose valid row and value fields.')

      const grouped = new Map()
      matrix.slice(1).forEach((row) => {
        const key = String(row?.[groupIndex] ?? '(blank)')
        const raw = row?.[valueIndex]
        const numeric = Number(raw)
        const current = grouped.get(key) || { sum: 0, count: 0 }
        current.count += 1
        if (raw !== '' && raw !== null && raw !== undefined && Number.isFinite(numeric)) current.sum += numeric
        grouped.set(key, current)
      })

      const resultRows = Array.from(grouped.entries()).sort((a, b) => a[0].localeCompare(b[0])).map(([key, value]) => {
        const result = aggregation === 'COUNT' ? value.count : aggregation === 'AVERAGE'
          ? (value.count ? value.sum / value.count : 0)
          : value.sum
        return [key, result]
      })
      const workbook = api.getActiveWorkbook()
      const sheetNameBase = 'Pivot Summary'
      const existingNames = new Set(workbook.getSheets().map((sheet) => sheet.getName()))
      let sheetName = sheetNameBase
      let counter = 2
      while (existingNames.has(sheetName)) {
        sheetName = sheetNameBase + ' ' + counter
        counter += 1
      }

      const summarySheet = workbook.create(sheetName, Math.max(20, resultRows.length + 2), 4)
      summarySheet.getRange('A1:B1').setValues([[rowField, aggregation + ' of ' + valueField]])
      summarySheet.getRange('A1:B1').setFontWeight('bold')
      if (resultRows.length) {
        summarySheet.getRange(1, 0, resultRows.length, 2).setValues(resultRows)
      }
      summarySheet.getRange(1 + resultRows.length, 0, 1, 2).setValues([['Grand Total', aggregation === 'COUNT'
        ? matrix.length - 1
        : aggregation === 'AVERAGE'
          ? (resultRows.length ? resultRows.reduce((sum, row) => sum + Number(row[1] || 0), 0) / resultRows.length : 0)
          : resultRows.reduce((sum, row) => sum + Number(row[1] || 0), 0)]])
      summarySheet.activate()
      setMessage('Created ' + sheetName + ' as a new worksheet.')
    } catch (reason) {
      setError(reason instanceof Error ? reason.message : 'Could not create the pivot summary.')
    }
  }

  return (
    <div className="insert-tools-backdrop" role="presentation" onMouseDown={(event) => {
      if (event.target === event.currentTarget) onClose?.()
    }}>
      <aside className="insert-tools-panel" role="dialog" aria-modal="true" aria-labelledby="insert-tools-title">
        <div className="insert-tools-header">
          <div>
            <p className="eyebrow">WORKBOOK TOOLS</p>
            <h2 id="insert-tools-title">Insert &amp; analyze</h2>
          </div>
          <button type="button" className="panel-close" onClick={onClose} aria-label="Close tools">×</button>
        </div>

        <div className="insert-tool-tabs" role="tablist" aria-label="Insert tools">
          <button type="button" role="tab" aria-selected={activeTool === 'chart'} className={activeTool === 'chart' ? 'selected' : ''} onClick={() => { setActiveTool('chart'); setError(''); setMessage('') }}>Chart preview</button>
          <button type="button" role="tab" aria-selected={activeTool === 'pivot'} className={activeTool === 'pivot' ? 'selected' : ''} onClick={() => { setActiveTool('pivot'); setError(''); setMessage('') }}>Pivot summary</button>
        </div>

        <label className="panel-field">
          <span>Source range</span>
          <input value={range} onChange={(event) => setRange(event.target.value)} placeholder="A1:B8" />
          <small>Use the active worksheet. The first row should contain column headings.</small>
        </label>

        {activeTool === 'chart' ? (
          <div className="panel-tool-body">
            <label className="panel-field">
              <span>Chart type</span>
              <select value={chartType} onChange={(event) => setChartType(event.target.value)}>
                <option value="column">Column chart</option>
                <option value="line">Line chart</option>
              </select>
            </label>
            <button className="panel-primary" type="button" onClick={previewChart}>Preview chart</button>
            {chartModel && (
              <div className="chart-preview-card">
                <div className="chart-preview-title">
                  <strong>{chartModel.title}</strong>
                  <span>{chartType === 'line' ? 'Line chart' : 'Column chart'}</span>
                </div>
                <SvgChart model={chartModel} type={chartType} />
              </div>
            )}
            <p className="panel-hint">This free build provides a working chart preview from the selected range. Embedded, editable chart objects inside the worksheet are a separate advanced milestone.</p>
          </div>
        ) : (
          <div className="panel-tool-body">
            <button className="panel-secondary" type="button" onClick={inspectPivotRange}>Read source columns</button>
            {cleanHeaders.length > 1 && (
              <>
                <label className="panel-field">
                  <span>Rows grouped by</span>
                  <select value={rowField} onChange={(event) => setRowField(event.target.value)}>
                    {cleanHeaders.map((header) => <option key={header} value={header}>{header}</option>)}
                  </select>
                </label>
                <label className="panel-field">
                  <span>Values</span>
                  <select value={valueField} onChange={(event) => setValueField(event.target.value)}>
                    {cleanHeaders.map((header) => <option key={header} value={header}>{header}</option>)}
                  </select>
                </label>
                <label className="panel-field">
                  <span>Summarize values by</span>
                  <select value={aggregation} onChange={(event) => setAggregation(event.target.value)}>
                    <option value="SUM">Sum</option>
                    <option value="COUNT">Count records</option>
                    <option value="AVERAGE">Average</option>
                  </select>
                </label>
                <button className="panel-primary" type="button" onClick={createPivotSummary}>Create summary sheet</button>
              </>
            )}
            <p className="panel-hint">Creates a grouped summary on a new worksheet. This is a basic pivot-summary workflow; drag-and-drop pivot fields, slicers and refreshable PivotTable objects are not included in this build.</p>
          </div>
        )}

        {error && <div className="panel-message error" role="alert">{error}</div>}
        {message && <div className="panel-message success" role="status">{message}</div>}
      </aside>
    </div>
  )
}
