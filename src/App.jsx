
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

const normalizeDate = (value) => {
  const text = clean(value)
  if (!text || text === '—') return ''

  const match = text.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})$/)

  if (match) {
    let [, day, month, year] = match
    if (year.length === 2) year = `20${year}`

    return `${day.padStart(2, '0')}-${month.padStart(2, '0')}-${year}`
  }

  const parsed = new Date(text)

  if (!Number.isNaN(parsed.getTime())) {
    const day = String(parsed.getDate()).padStart(2, '0')
    const month = String(parsed.getMonth() + 1).padStart(2, '0')
    const year = parsed.getFullYear()
    return `${day}-${month}-${year}`
  }

  return text
}

function App() {
  const [workbook, setWorkbook] = useState(null)
  const [fileName, setFileName] = useState('')
  const [selectedSheet, setSelectedSheet] = useState(SHEETS[0])
  const [qualityFilter, setQualityFilter] = useState('')
  const [widthFilter, setWidthFilter] = useState('')
  const [igpFilter, setIgpFilter] = useState('')
  const [poFilter, setPoFilter] = useState('')
  const [serialFilter, setSerialFilter] = useState('')
  const [lastReceivedFilter, setLastReceivedFilter] = useState('')
  const [error, setError] = useState('')
  const [hideZeroStock, setHideZeroStock] = useState(false)

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
        setSerialFilter('')
        setLastReceivedFilter('')
      } catch {
        setError(
          'Excel file read nahi ho saki. Please correct Excel file upload karein.',
        )
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
      lastReceived: findColumn([/^last\s*received$/i]),

      quality: findColumn([
        /^quality$/i,
        /^quality\s*\/\s*construction$/i,
      ]),

      width: findColumn([
        /^width$/i,
        /^witdh$/i,
        /^g\.?\s*width$/i,
        /^g\.?\s*witdh$/i,
      ]),

      igp: findColumn([/\bigp\b/i, /\bgrn\b/i]),
      po: findColumn([/\bpo\b/i, /contract/i]),
      serial: findColumn([/^sr\.?[-\s]*$/i, /^serial(?:\s*number)?$/i]),
      itemCode: findColumn([/^item\s*code$/i, /^item\s*#$/i]),

      // Excel ka Type column
      type: findColumn([/^type$/i]),

      weaveBlendSelvedge: findColumn([
        /^weave\s*,?\s*blend\s*,?\s*selvedge$/i,
        /^weave\s*blend\s*selvedge$/i,
      ]),
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
    const headerIndex = data.headers.findIndex((header) =>
      patterns.some((pattern) => pattern.test(clean(header))),
    )

    if (headerIndex === -1) return '—'

    const value = clean(row.values[headerIndex])
    return value || '—'
  }

  const getColumnCell = (row, columnIndex) => {
    if (columnIndex === undefined || columnIndex === null || columnIndex < 0) {
      return '—'
    }

    const value = clean(row.values[columnIndex])
    return value || '—'
  }

  const getStockCell = (row, columnIndex) => {
    if (columnIndex === undefined || columnIndex === null) return '—'

    const value = clean(row.values[columnIndex])
    return value || '0'
  }

  const filteredRows = useMemo(() => {
    return data.rows.filter((row) => {
      const valueAt = (column) =>
        column >= 0 ? clean(row.values[column]) : ''

      const quality = normalize(valueAt(data.columns.quality))
      const width = normalize(valueAt(data.columns.width))
      const igp = normalize(valueAt(data.columns.igp))
      const po = normalize(valueAt(data.columns.po))
      const serial = normalize(valueAt(data.columns.serial))

      const lastReceived = normalizeDate(valueAt(data.columns.lastReceived))
      const requestedDate = normalizeDate(lastReceivedFilter)

      const parseStock = (columnIndex) => {
        if (
          !data.stockColumns ||
          columnIndex === undefined ||
          columnIndex === null
        ) {
          return 0
        }

        const rawValue = clean(row.values[columnIndex]).replace(/,/g, '')
        const numberValue = Number(rawValue)

        return rawValue === '' || !Number.isFinite(numberValue)
          ? 0
          : numberValue
      }

      const rolls = parseStock(data.stockColumns?.rolls)
      const meters = parseStock(data.stockColumns?.aGrade)
      const hasStock = rolls !== 0 || meters !== 0

      return (
        (!hideZeroStock || hasStock) &&
        (!qualityFilter || quality.includes(normalize(qualityFilter))) &&
        (!widthFilter || width.includes(normalize(widthFilter))) &&
        (!igpFilter || igp.includes(normalize(igpFilter))) &&
        (!poFilter || po.includes(normalize(poFilter))) &&
        (!serialFilter || serial.includes(normalize(serialFilter))) &&
        (!lastReceivedFilter ||
          normalize(lastReceived) === normalize(requestedDate))
      )
    })
  }, [
    data,
    qualityFilter,
    widthFilter,
    igpFilter,
    poFilter,
    serialFilter,
    lastReceivedFilter,
    hideZeroStock,
  ])

  const clearFilters = () => {
    setQualityFilter('')
    setWidthFilter('')
    setIgpFilter('')
    setPoFilter('')
    setSerialFilter('')
    setLastReceivedFilter('')
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

      <section className="panel">
        <label htmlFor="excel-upload">
          Upload complete Excel stock file
        </label>

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

          <button
            className="clear-button"
            onClick={clearFilters}
            type="button"
          >
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

          <div className="field">
            <label htmlFor="serial-filter">Serial (Sr-)</label>
            <input
              id="serial-filter"
              value={serialFilter}
              onChange={(event) => setSerialFilter(event.target.value)}
              placeholder="Enter serial number"
            />
          </div>

          <div className="field">
            <label htmlFor="last-received-filter">
              Last Received (DD-MM-YYYY)
            </label>
            <input
              id="last-received-filter"
              value={lastReceivedFilter}
              onChange={(event) => setLastReceivedFilter(event.target.value)}
              placeholder="DD-MM-YYYY"
            />
          </div>
        </div>

        <div
          style={{
            display: 'flex',
            alignItems: 'center',
            gap: '8px',
            marginTop: '16px',
            fontSize: '13px',
            fontWeight: 700,
          }}
        >
          <input
            id="hide-zero-stock"
            type="checkbox"
            checked={hideZeroStock}
            onChange={(event) => setHideZeroStock(event.target.checked)}
            style={{ width: '16px', height: '16px', cursor: 'pointer' }}
          />

          <label htmlFor="hide-zero-stock" style={{ cursor: 'pointer' }}>
            Hide Zero Stock
          </label>
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
              const quality = getColumnCell(row, data.columns.quality)
              const width = getColumnCell(row, data.columns.width)

              const lastReceived = normalizeDate(
                getColumnCell(row, data.columns.lastReceived),
              )

              return (
                <article className="stock-card" key={row.id}>
                  <div className="stock-date">
                    <span>Last Received</span>
                    <strong>{lastReceived || '—'}</strong>
                  </div>

                  <div className="stock-card-title">
                    <div>
                      <span className="card-eyebrow">STOCK ITEM</span>
                      <h3>
                        {quality === '—' ? 'Quality not available' : quality}
                      </h3>
                    </div>

                    <span className="width-tag">Width: {width}</span>
                  </div>

                  <div className="record-details">
                    <div className="detail-row item-code-box">
                      <span className="detail-label">Item Code</span>
                      <span className="detail-value">
                        {getColumnCell(row, data.columns.itemCode)}
                      </span>
                    </div>

                    <div className="detail-row serial-box">
                      <span className="detail-label">Sr-</span>
                      <span className="detail-value">
                        {getColumnCell(row, data.columns.serial)}
                      </span>
                    </div>

                    <div className="detail-row weave-box">
                      <span className="detail-label">
                        Weave, Blend, Selvedge
                      </span>
                      <span className="detail-value">
                        {getColumnCell(row, data.columns.weaveBlendSelvedge)}
                      </span>
                    </div>

                    <div className="detail-row supplier-box">
                      <span className="detail-label">Supplier</span>
                      <span className="detail-value">
                        {getCell(row, [/^greige\s*supplier$/i])}
                      </span>
                    </div>

                    <div className="detail-row customer-box">
                      <span className="detail-label">Customer</span>
                      <span className="detail-value">
                        {getCell(row, [/^processing\s*customer$/i])}
                      </span>
                    </div>
                  </div>

                  <div className="stock-amounts">
                    {/* Type aur R-B ab ek hi box mein hain */}
                    <div className="amount-box rolls-box">
                      <span className="amount-label">Type / Rolls-Bales</span>

                      <strong className="amount-value">
                        {getColumnCell(row, data.columns.type)} /{' '}
                        {getStockCell(row, data.stockColumns?.rolls)}
                      </strong>

                      <span className="amount-unit">Type / R-B</span>
                    </div>

                    <div className="amount-box a-grade-box">
                      <span className="amount-label">A-Grade</span>

                      <strong className="amount-value">
                        {getStockCell(row, data.stockColumns?.aGrade)}
                      </strong>

                      <span className="amount-unit">MTR</span>
                    </div>
                  </div>

                  <div className="remarks-box">
                    <span className="remarks-label">Remarks</span>
                    <p>{getCell(row, [/^remarks?$/i, /remarks?/i])}</p>
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
