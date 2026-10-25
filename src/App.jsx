import { useMemo, useState } from 'react'
import * as XLSX from 'xlsx'
import './App.css'

const SHEETS = [
  'Daily Stock Report (Own)',
  'Daily Stock Report (Commercial)',
  'Free For Sale (Own)',
]

const clean = (value) =>
  String(value ?? '')
    .replace(/\s+/g, ' ')
    .trim()

const normalize = (value) => clean(value).toLowerCase()

function App() {
  const [workbook, setWorkbook] = useState(null)
  const [fileName, setFileName] = useState('')
  const [selectedSheet, setSelectedSheet] = useState(SHEETS[0])
  const [qualityFilter, setQualityFilter] = useState('')
  const [widthFilter, setWidthFilter] = useState('')
  const [igpFilter, setIgpFilter] = useState('')
  const [poFilter, setPoFilter] = useState('')
  const [error, setError] = useState('')

  const handleUpload = (event) => {
    const file = event.target.files?.[0]
    if (!file) return

    setError('')

    const reader = new FileReader()

    reader.onload = (e) => {
      try {
        const bytes = new Uint8Array(e.target.result)
        const newWorkbook = XLSX.read(bytes, { type: 'array' })

        const missingSheets = SHEETS.filter(
          (sheetName) => !newWorkbook.SheetNames.includes(sheetName),
        )

        if (missingSheets.length) {
          setError(
            `Excel file mein required sheet nahi mili: ${missingSheets.join(', ')}`,
          )
          return
        }

        setWorkbook(newWorkbook)
        setFileName(file.name)
        setSelectedSheet(SHEETS[0])
        setQualityFilter('')
        setWidthFilter('')
        setIgpFilter('')
        setPoFilter('')
      } catch {
        setError('Excel file read nahi ho saki. Please correct Excel file upload karein.')
      }
    }

    reader.onerror = () => {
      setError('File read nahi ho saki. Dobara upload karein.')
    }

    reader.readAsArrayBuffer(file)
  }

  const data = useMemo(() => {
    if (!workbook || !workbook.Sheets[selectedSheet]) {
      return { rows: [], columns: {}, stockColumns: null, headers: [] }
    }

    const sheet = workbook.Sheets[selectedSheet]
    const matrix = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      defval: '',
      raw: false,
    })

    const headerIndex = matrix.findIndex((row) => {
      const values = row.map(normalize)
      return (
        values.some((value) => /quality/.test(value)) &&
        values.some((value) => /width|witdh/.test(value))
      )
    })

    if (headerIndex === -1) {
      return { rows: [], columns: {}, stockColumns: null, headers: [] }
    }

    const headers = matrix[headerIndex].map(clean)

    const findColumn = (patterns) =>
      headers.findIndex((header) =>
        patterns.some((pattern) => pattern.test(header)),
      )

    const columns = {
      date: findColumn([/^date$/i]),
      quality: findColumn([/^quality$/i, /^quality\s*\/\s*construction$/i]),
      width: findColumn([/^width$/i, /^witdh$/i, /^g\.?\s*width$/i, /^g\.?\s*witdh$/i]),
      igp: findColumn([/\bigp\b/i, /\bgrn\b/i]),
      po: findColumn([/\bpo\b/i, /contract/i]),
    }

    const closingRowIndex = matrix.findIndex((row) =>
      row.some((cell) => /closing\s*balance/i.test(clean(cell))),
    )

    let stockColumns = null

    if (closingRowIndex !== -1) {
      const closingRow = matrix[closingRowIndex]
      const closingColumn = closingRow.findIndex((cell) =>
        /closing\s*balance/i.test(clean(cell)),
      )

      if (closingColumn !== -1) {
        stockColumns = {
          rolls: closingColumn,
          aGrade: closingColumn + 1,
          bGrade: closingColumn + 2,
        }
      }
    }

    const rows = matrix
      .slice(headerIndex + 1)
      .filter((row) => row.some((cell) => clean(cell) !== ''))
      .map((row, index) => ({
        id: index,
        values: row,
      }))

    return { rows, columns, stockColumns, headers }
  }, [workbook, selectedSheet])

  const getCell = (row, patterns) => {
    const headerIndex = data.headers?.findIndex((header) =>
      patterns.some((pattern) => pattern.test(clean(header))),
    )

    if (headerIndex === undefined || headerIndex === -1) return '—'

    const value = clean(row.values[headerIndex])
    return value || '—'
  }

  const getStockCell = (row, columnIndex) => {
    if (columnIndex === undefined || columnIndex === null) return '—'
    const value = clean(row.values[columnIndex])
    return value || '0'
  }

  const formatDate = (value) => {
    if (value === undefined || value === null || clean(value) === '') {
      return '—'
    }

    return clean(value)
  }

  const filteredRows = useMemo(() => {
    return data.rows.filter((row) => {
      const quality =
        data.columns.quality >= 0
          ? normalize(row.values[data.columns.quality])
          : ''
      const width =
        data.columns.width >= 0
          ? normalize(row.values[data.columns.width])
          : ''
      const igp =
        data.columns.igp >= 0
          ? normalize(row.values[data.columns.igp])
          : ''
      const po =
        data.columns.po >= 0
          ? normalize(row.values[data.columns.po])
          : ''

      return (
        (!qualityFilter || quality.includes(normalize(qualityFilter))) &&
        (!widthFilter || width.includes(normalize(widthFilter))) &&
        (!igpFilter || igp.includes(normalize(igpFilter))) &&
        (!poFilter || po.includes(normalize(poFilter)))
      )
    })
  }, [data, qualityFilter, widthFilter, igpFilter, poFilter])

  const clearFilters = () => {
    setQualityFilter('')
    setWidthFilter('')
    setIgpFilter('')
    setPoFilter('')
  }

  return (
    <main className="app">
      <header className="topbar">
        <div className="brand">
          <div className="brand-icon">S</div>
          <div>
            <h1>Stock Search</h1>
            <p>Inventory lookup dashboard</p>
          </div>
        </div>

        <div className="status">
          {workbook ? 'Excel Loaded' : 'Awaiting Excel'}
        </div>
      </header>

      <section className="intro">
        <span className="eyebrow">INVENTORY MANAGEMENT</span>
        <h2>Find your stock, faster.</h2>
        <p>
          Upload the latest stock report, select a sheet and search your
          inventory using the available filters.
        </p>
      </section>

      <section className="panel">
        <label htmlFor="excel-upload">Upload complete Excel stock file</label>
        <input
          id="excel-upload"
          type="file"
          accept=".xlsx,.xls"
          onChange={handleUpload}
        />

        {fileName && (
          <div className="file-info">
            <span>Uploaded file</span>
            <strong>{fileName}</strong>
          </div>
        )}

        {error && <div className="error-message">{error}</div>}

        <div className="section-heading">
          <div>
            <h3>Stock Sheet</h3>
            <p>Select the stock report you want to search.</p>
          </div>
        </div>

        <label htmlFor="sheet-select">Select sheet</label>
        <select
          id="sheet-select"
          value={selectedSheet}
          onChange={(event) => setSelectedSheet(event.target.value)}
        >
          {SHEETS.map((sheetName) => (
            <option key={sheetName} value={sheetName}>
              {sheetName}
            </option>
          ))}
        </select>

        <div className="section-heading">
          <div>
            <h3>Search Filters</h3>
            <p>Enter one or more values to narrow down the results.</p>
          </div>
          <button className="clear-button" onClick={clearFilters} type="button">
            Clear filters
          </button>
        </div>

        <div className="filters">
          <div className="field">
            <label htmlFor="quality-filter">Quality / Construction</label>
            <input
              id="quality-filter"
              value={qualityFilter}
              onChange={(event) => setQualityFilter(event.target.value)}
              placeholder="Enter quality"
            />
          </div>

          <div className="field">
            <label htmlFor="width-filter">Width</label>
            <input
              id="width-filter"
              value={widthFilter}
              onChange={(event) => setWidthFilter(event.target.value)}
              placeholder="Enter width"
            />
          </div>

          <div className="field">
            <label htmlFor="igp-filter">IGP / GRN</label>
            <input
              id="igp-filter"
              value={igpFilter}
              onChange={(event) => setIgpFilter(event.target.value)}
              placeholder="Search IGP / GRN"
            />
          </div>

          <div className="field">
            <label htmlFor="po-filter">PO / Contract</label>
            <input
              id="po-filter"
              value={poFilter}
              onChange={(event) => setPoFilter(event.target.value)}
              placeholder="Search PO / Contract"
            />
          </div>
        </div>

        <div className="notice">
          <strong>Quick tip</strong>
          <p>
            You can use any combination of filters. Stock amounts are shown
            directly on each matching card.
          </p>
        </div>
      </section>

      <section className="results-section">
        <div className="results-heading">
          <div>
            <h3>Stock Results</h3>
            <p>
              {workbook
                ? `Showing records from ${selectedSheet}`
                : 'Upload an Excel file to view your stock.'}
            </p>
          </div>
          <span className="count">{filteredRows.length} records</span>
        </div>

        {!workbook ? (
          <div className="empty-state">
            <div className="empty-icon">↥</div>
            <h3>Upload your Excel file</h3>
            <p>
              Select your updated stock workbook above. Your matching stock
              records will appear here.
            </p>
          </div>
        ) : filteredRows.length === 0 ? (
          <div className="empty-state">
            <div className="empty-icon">⌕</div>
            <h3>No matching records</h3>
            <p>
              Try changing your filters or clear them to view the available
              records.
            </p>
          </div>
        ) : (
          <div className="stock-list">
            {filteredRows.map((row) => {
              const quality =
                data.columns.quality >= 0
                  ? clean(row.values[data.columns.quality]) || 'Quality not available'
                  : 'Quality not available'

              const width =
                data.columns.width >= 0
                  ? clean(row.values[data.columns.width]) || '—'
                  : '—'

              const date =
                data.columns.date >= 0
                  ? formatDate(row.values[data.columns.date])
                  : '—'

              return (
                <article className="stock-card" key={row.id}>
                  <div className="stock-date">
                    <span>Date</span>
                    <strong>{date}</strong>
                  </div>

                  <div className="stock-card-title">
                    <div>
                      <span className="card-eyebrow">STOCK ITEM</span>
                      <h3>{quality}</h3>
                    </div>
                    <span className="width-tag">Width: {width}</span>
                  </div>

                  <div className="record-details">
                    <div className="detail-row">
                      <span className="detail-label">Supplier</span>
                      <span className="detail-value">
                        {getCell(row, [/^greige\s*supplier$/i])}
                      </span>
                    </div>

                    <div className="detail-row">
                      <span className="detail-label">Customer</span>
                      <span className="detail-value">
                        {getCell(row, [/^processing\s*customer$/i])}
                      </span>
                    </div>
                  </div>

                  <div className="stock-amounts">
                    <div className="amount-box rolls-box">
                      <span className="amount-label">Rolls / Bales</span>
                      <strong className="amount-value">
                        {getStockCell(row, data.stockColumns?.rolls)}
                      </strong>
                      <span className="amount-unit">R-B</span>
                    </div>

                    <div className="amount-box a-grade-box">
                      <span className="amount-label">A-Grade</span>
                      <strong className="amount-value">
                        {getStockCell(row, data.stockColumns?.aGrade)}
                      </strong>
                      <span className="amount-unit">MTR</span>
                    </div>

                    <div className="amount-box b-grade-box">
                      <span className="amount-label">B-Grade</span>
                      <strong className="amount-value">
                        {getStockCell(row, data.stockColumns?.bGrade)}
                      </strong>
                      <span className="amount-unit">MTR</span>
                    </div>
                  </div>

                  <div className="remarks-box">
                    <span className="remarks-label">Remarks</span>
                    <p>
                      {getCell(row, [/^remarks?$/i, /remarks?/i])}
                    </p>
                  </div>
                </article>
              )
            })}
          </div>
        )}
      </section>

      <footer>Stock Search Dashboard</footer>
    </main>
  )
}

export default App