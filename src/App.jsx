import { useEffect, useMemo, useState } from 'react'
import * as XLSX from 'xlsx'
import { collection, doc, getDoc, getDocs, setDoc, deleteDoc } from 'firebase/firestore'
import { db } from './firebase.js'
import './App.css'

const UPLOAD_PASSWORD = "Stock@2025" // <-- Yahan apna password change kar sakte ho

const bytesToBase64 = (bytes) => {
  let binary = ''
  for (let i = 0; i < bytes.byteLength; i++) binary += String.fromCharCode(bytes[i])
  return btoa(binary)
}
const base64ToBytes = (base64) => {
  const binary = atob(base64)
  const bytes = new Uint8Array(binary.length)
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i)
  return bytes
}

const SHEETS = [
  'Daily Stock Report (Own)',
  'Daily Stock Report (Commercial)',
  'Free For Sale (Own)',
   'Free Of Cost (Own)',
]

const clean = (value) => String(value ?? '').replace(/\s+/g, ' ').trim()
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

const getWorkbookStockDate = (workbook) => {
  for (const sheetName of SHEETS) {
    const sheet = workbook.Sheets[sheetName]
    if (!sheet) continue
    const cell = sheet.A3
    if (!cell || cell.v === undefined || cell.v === null || clean(cell.v) === '') continue
    let value = cell.v
    if (cell.t === 'n' && typeof value === 'number') {
      const parsed = XLSX.SSF.parse_date_code(value)
      if (parsed) {
        value = `${String(parsed.d).padStart(2, '0')}-${String(parsed.m).padStart(2, '0')}-${parsed.y}`
        return value
      }
    }
    const formatted = cell.w ? clean(cell.w) : clean(value)
    const normalized = normalizeDate(formatted)
    if (normalized) return normalized
  }
  return ''
}

function App() {
  const [workbook, setWorkbook] = useState(null)
  const [fileName, setFileName] = useState('')
  const [stockUpdatedTill, setStockUpdatedTill] = useState('')
  const [selectedSheet, setSelectedSheet] = useState(SHEETS[0])
  const [qualityFilter, setQualityFilter] = useState('')
  const [widthFilter, setWidthFilter] = useState('')
  const [igpFilter, setIgpFilter] = useState('')
  const [poFilter, setPoFilter] = useState('')
  const [supplierFilter, setSupplierFilter] = useState('')
  const [customerFilter, setCustomerFilter] = useState('')
  const [serialFilter, setSerialFilter] = useState('')
  const [lastReceivedFilter, setLastReceivedFilter] = useState('')
  const [error, setError] = useState('')
  const [hideZeroStock, setHideZeroStock] = useState(false)
  const [isSaving, setIsSaving] = useState(false)
  const [isLoading, setIsLoading] = useState(true)

  useEffect(() => {
    let cancelled = false
    const loadSavedWorkbook = async () => {
      try {
        const currentRef = doc(db, 'stock', 'current')
        const currentSnap = await getDoc(currentRef)
        if (!currentSnap.exists()) { setIsLoading(false); return }
        const saved = currentSnap.data()
        const chunksSnap = await getDocs(collection(db, 'stockBackups', saved.uploadId, 'chunks'))
        const chunks = chunksSnap.docs.map((item) => item.data()).sort((a, b) => a.index - b.index)
        if (cancelled || chunks.length !== saved.chunkCount || chunks.length === 0) { setIsLoading(false); return }
        const base64 = chunks.map((item) => item.data).join('')
        const bytes = base64ToBytes(base64)
        const savedWorkbook = XLSX.read(bytes, { type: 'array', cellDates: true })
        if (cancelled) return
        setWorkbook(savedWorkbook)
        setFileName(saved.fileName || 'Saved Excel file')
        setStockUpdatedTill(saved.stockUpdatedTill || '')
      } catch (err) {
        console.error('Load failed:', err)
      } finally {
        if (!cancelled) setIsLoading(false)
      }
    }
    loadSavedWorkbook()
    return () => { cancelled = true }
  }, [])

  const handleUpload = async (event) => {
    const file = event.target.files?.[0]
    if (!file) return

    // --- PASSWORD CHECK ---
    const enteredPassword = prompt("Upload karne ke liye password likho:")
    if (enteredPassword === null) { // User ne cancel dabaya
      event.target.value = ''
      return
    }
    if (enteredPassword !== UPLOAD_PASSWORD) {
      setError("Ghalat password! Upload nahi ho sakta.")
      event.target.value = ''
      return
    }
    // --- END PASSWORD CHECK ---

    setError('')
    setIsSaving(true)
    const reader = new FileReader()
    reader.onload = async (e) => {
      try {
        const bytes = new Uint8Array(e.target.result)
        const newWorkbook = XLSX.read(bytes, { type: 'array', cellDates: true })
        const missingSheets = SHEETS.filter((sheetName) => !newWorkbook.SheetNames.includes(sheetName))
        if (missingSheets.length) {
          setError(`Excel file mein required sheet nahi mili: ${missingSheets.join(', ')}`)
          setIsSaving(false)
          event.target.value = ''
          return
        }
        const newStockDate = getWorkbookStockDate(newWorkbook)
        const base64Data = bytesToBase64(bytes)
        const uploadId = Date.now().toString()
        const CHUNK_SIZE = 800000
        const totalChunks = Math.ceil(base64Data.length / CHUNK_SIZE)

        const currentRef = doc(db, 'stock', 'current')
        const currentSnap = await getDoc(currentRef)
        if (currentSnap.exists()) {
          const oldId = currentSnap.data().uploadId
          try {
            const oldChunks = await getDocs(collection(db, 'stockBackups', oldId, 'chunks'))
            for (const d of oldChunks.docs) await deleteDoc(d.ref)
          } catch {}
        }

        for (let i = 0; i < totalChunks; i++) {
          const chunkData = base64Data.slice(i * CHUNK_SIZE, (i + 1) * CHUNK_SIZE)
          await setDoc(doc(db, 'stockBackups', uploadId, 'chunks', i.toString()), { index: i, data: chunkData })
        }
        await setDoc(doc(db, 'stock', 'current'), {
          uploadId, fileName: file.name, stockUpdatedTill: newStockDate, chunkCount: totalChunks, createdAt: new Date().toISOString(),
        })

        setWorkbook(newWorkbook)
        setFileName(file.name)
        setStockUpdatedTill(newStockDate)
        setSelectedSheet(SHEETS[0])
        setQualityFilter(''); setWidthFilter(''); setIgpFilter(''); setPoFilter(''); setSupplierFilter(''); setCustomerFilter(''); setSerialFilter(''); setLastReceivedFilter('')
        setError('File kamyabi se upload ho gayi!')
      } catch {
        setError('Excel file read nahi ho saki.')
      } finally {
        setIsSaving(false)
        event.target.value = ''
      }
    }
    reader.onerror = () => { setError('File read nahi ho saki.'); setIsSaving(false); event.target.value = '' }
    reader.readAsArrayBuffer(file)
  }

  const data = useMemo(() => {
    if (!workbook || !workbook.Sheets[selectedSheet]) return { rows: [], columns: {}, stockColumns: null, headers: [] }
    const sheet = workbook.Sheets[selectedSheet]
    const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: false })
    const headerIndex = matrix.findIndex((row) => {
      const values = row.map(normalize)
      return values.some((value) => /quality/.test(value)) && values.some((value) => /width|witdh/.test(value))
    })
    if (headerIndex === -1) return { rows: [], columns: {}, stockColumns: null, headers: [] }
    const headers = matrix[headerIndex].map(clean)
    const findColumn = (patterns) => headers.findIndex((header) => patterns.some((pattern) => pattern.test(header)))
    const columns = {
      lastReceived: findColumn([/^last\s*received$/i]),
      quality: findColumn([/^quality$/i, /^quality\s*\/\s*construction$/i]),
      width: findColumn([/^width$/i, /^witdh$/i, /^g\.?\s*width$/i, /^g\.?\s*witdh$/i]),
      igp: findColumn([/\bigp\b/i, /\bgrn\b/i]),
      po: findColumn([/\bpo\b/i, /contract/i]),
      supplier: findColumn([/^greige\s*supplier$/i]),
      customer: findColumn([/^processing\s*customer$/i]),
      serial: findColumn([/^sr\.?[-\s]*$/i, /^serial(?:\s*number)?$/i]),
      itemCode: findColumn([/^item\s*code$/i, /^item\s*#$/i]),
      type: findColumn([/^type$/i]),
      weaveBlendSelvedge: findColumn([/^weave\s*,?\s*blend\s*,?\s*selvedge$/i, /^weave\s*blend\s*selvedge$/i]),
    }
    const closingRowIndex = matrix.findIndex((row) => row.some((cell) => /closing\s*balance/i.test(clean(cell))))
    let stockColumns = null
    if (closingRowIndex !== -1) {
      const closingRow = matrix[closingRowIndex]
      const closingColumn = closingRow.findIndex((cell) => /closing\s*balance/i.test(clean(cell)))
      if (closingColumn !== -1) stockColumns = { rolls: closingColumn, aGrade: closingColumn + 1 }
    }
    const rows = matrix.slice(headerIndex + 1).filter((row) => {
      const quality = clean(row[columns.quality]); const width = clean(row[columns.width]); const weaveBlendSelvedge = clean(row[columns.weaveBlendSelvedge])
      return Boolean(quality || width || weaveBlendSelvedge)
    }).map((row, index) => ({ id: index, values: row }))
    return { rows, columns, stockColumns, headers }
  }, [workbook, selectedSheet])

  const getCell = (row, patterns) => {
    const headerIndex = data.headers.findIndex((header) => patterns.some((pattern) => pattern.test(clean(header))))
    if (headerIndex === -1) return '—'
    const value = clean(row.values[headerIndex])
    return value || '—'
  }
  const getColumnCell = (row, columnIndex) => {
    if (columnIndex === undefined || columnIndex === null || columnIndex < 0) return '—'
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
      const valueAt = (column) => column >= 0 ? clean(row.values[column]) : ''
      const quality = normalize(valueAt(data.columns.quality)), width = normalize(valueAt(data.columns.width)), igp = normalize(valueAt(data.columns.igp)), po = normalize(valueAt(data.columns.po)), supplier = normalize(valueAt(data.columns.supplier)), customer = normalize(valueAt(data.columns.customer)), serial = normalize(valueAt(data.columns.serial))
      const lastReceived = normalizeDate(valueAt(data.columns.lastReceived)), requestedDate = normalizeDate(lastReceivedFilter)
      const parseStock = (columnIndex) => {
        if (!data.stockColumns || columnIndex === undefined || columnIndex === null) return 0
        const rawValue = clean(row.values[columnIndex]).replace(/,/g, '')
        const numberValue = Number(rawValue)
        return rawValue === '' || !Number.isFinite(numberValue) ? 0 : numberValue
      }
      const rolls = parseStock(data.stockColumns?.rolls), meters = parseStock(data.stockColumns?.aGrade), hasStock = rolls !== 0 || meters !== 0
      return (!hideZeroStock || hasStock) && (!qualityFilter || quality.includes(normalize(qualityFilter))) && (!widthFilter || width.includes(normalize(widthFilter))) && (!igpFilter || igp.includes(normalize(igpFilter))) && (!poFilter || po.includes(normalize(poFilter))) && (!supplierFilter || supplier.includes(normalize(supplierFilter))) && (!customerFilter || customer.includes(normalize(customerFilter))) && (!serialFilter || serial.includes(normalize(serialFilter))) && (!lastReceivedFilter || normalize(lastReceived) === normalize(requestedDate))
    })
  }, [data, qualityFilter, widthFilter, igpFilter, poFilter, supplierFilter, customerFilter, serialFilter, lastReceivedFilter, hideZeroStock])

  const clearFilters = () => { setQualityFilter(''); setWidthFilter(''); setIgpFilter(''); setPoFilter(''); setSupplierFilter(''); setCustomerFilter(''); setSerialFilter(''); setLastReceivedFilter('') }

  if (isLoading) return <main className="app"><div className="empty-state"><h3>Loading saved stock...</h3><p>Firebase se purana data load ho raha hai</p></div></main>

  return (
    <main className="app">
      <header className="topbar">
        <div className="brand"><div className="brand-icon">S</div><div><h1>Stock Search</h1><p>Inventory lookup dashboard</p></div></div>
        <div className="status">{isSaving ? 'Saving...' : workbook ? 'Excel Loaded' : 'Awaiting Excel'}</div>
      </header>
      {workbook && <div className="stock-updated-till"><strong>Stock Updated Till:</strong> <span>{stockUpdatedTill || 'Date not found in Row 3'}</span></div>}
      <section className="panel">
        <label htmlFor="excel-upload">Upload complete Excel stock file</label>
        <input id="excel-upload" type="file" accept=".xlsx,.xls" onChange={handleUpload} disabled={isSaving} />
        {fileName && <div className="file-info"><span>Uploaded file</span><strong>{fileName}</strong></div>}
        {error && <div className="error-message">{error}</div>}
        <div className="section-heading"><div><h3>Stock Sheet</h3><p>Select the stock report you want to search.</p></div></div>
        <label htmlFor="sheet-select">Select sheet</label>
        <select id="sheet-select" value={selectedSheet} onChange={(event) => setSelectedSheet(event.target.value)}>{SHEETS.map((sheetName) => <option key={sheetName} value={sheetName}>{sheetName}</option>)}</select>
        <div className="section-heading"><div><h3>Search Filters</h3><p>Enter one or more values to narrow down the results.</p></div><button className="clear-button" onClick={clearFilters} type="button">Clear filters</button></div>
        <div className="filters">
          <div className="field"><label htmlFor="quality-filter">Quality / Construction</label><input id="quality-filter" value={qualityFilter} onChange={(event) => setQualityFilter(event.target.value)} placeholder="Enter quality" /></div>
          <div className="field"><label htmlFor="width-filter">Width</label><input id="width-filter" value={widthFilter} onChange={(event) => setWidthFilter(event.target.value)} placeholder="Enter width" /></div>
          <div className="field"><label htmlFor="igp-filter">IGP / GRN</label><input id="igp-filter" value={igpFilter} onChange={(event) => setIgpFilter(event.target.value)} placeholder="Search IGP / GRN" /></div>
          <div className="field"><label htmlFor="po-filter">PO / Contract</label><input id="po-filter" value={poFilter} onChange={(event) => setPoFilter(event.target.value)} placeholder="Search PO / Contract" /></div>
          <div className="field"><label htmlFor="supplier-filter">Supplier</label><input id="supplier-filter" value={supplierFilter} onChange={(event) => setSupplierFilter(event.target.value)} placeholder="Search supplier" /></div>
          <div className="field"><label htmlFor="customer-filter">Customer</label><input id="customer-filter" value={customerFilter} onChange={(event) => setCustomerFilter(event.target.value)} placeholder="Search customer" /></div>
          <div className="field"><label htmlFor="serial-filter">Serial (Sr-)</label><input id="serial-filter" value={serialFilter} onChange={(event) => setSerialFilter(event.target.value)} placeholder="Enter serial number" /></div>
          <div className="field"><label htmlFor="last-received-filter">Last Received (DD-MM-YYYY)</label><input id="last-received-filter" value={lastReceivedFilter} onChange={(event) => setLastReceivedFilter(event.target.value)} placeholder="DD-MM-YYYY" /></div>
        </div>
        <div style={{ display: 'flex', alignItems: 'center', gap: '8px', marginTop: '16px', fontSize: '13px', fontWeight: 700 }}>
          <input id="hide-zero-stock" type="checkbox" checked={hideZeroStock} onChange={(event) => setHideZeroStock(event.target.checked)} style={{ width: '16px', height: '16px', cursor: 'pointer' }} />
          <label htmlFor="hide-zero-stock" style={{ cursor: 'pointer' }}>Hide Zero Stock</label>
        </div>
      </section>
      <section className="results-section">
        <div className="results-heading"><div><h3>Stock Results</h3><p>{workbook ? `Showing records from ${selectedSheet}` : 'Upload an Excel file to view your stock.'}</p></div><span className="count">{filteredRows.length} records</span></div>
        {!workbook ? (
          <div className="empty-state"><div className="empty-icon">↥</div><h3>Upload your Excel file</h3><p>Select your updated stock workbook above.</p></div>
        ) : filteredRows.length === 0 ? (
          <div className="empty-state"><div className="empty-icon">⌕</div><h3>No matching records</h3><p>Try changing your filters.</p></div>
        ) : (
          <div className="stock-list">
            {filteredRows.map((row) => {
              const quality = getColumnCell(row, data.columns.quality), width = getColumnCell(row, data.columns.width), igpNumber = getColumnCell(row, data.columns.igp), contractNumber = getColumnCell(row, data.columns.po), lastReceived = normalizeDate(getColumnCell(row, data.columns.lastReceived))
              return (
                <article className="stock-card" key={row.id}>
                  <div className="stock-date"><span>Last Received</span><strong>{lastReceived || '—'}</strong></div>
                  <div className="stock-card-title"><div><span className="card-eyebrow">STOCK ITEM</span><h3>{quality === '—' ? 'Quality not available' : quality}</h3></div><span className="width-tag">Width: {width}</span></div>
                  <div className="record-details">
                    <div className="detail-row item-code-box"><span className="detail-label">Item Code</span><span className="detail-value">{getColumnCell(row, data.columns.itemCode)}</span></div>
                    <div className="detail-row serial-box"><span className="detail-label">Sr-</span><span className="detail-value">{getColumnCell(row, data.columns.serial)}</span></div>
                    <div className="detail-row weave-box"><span className="detail-label">Weave, Blend, Selvedge</span><span className="detail-value">{getColumnCell(row, data.columns.weaveBlendSelvedge)}</span></div>
                    <div className="detail-row supplier-box"><span className="detail-label">Supplier / IGP-GRN</span><span className="detail-value">{getCell(row, [/^greige\s*supplier$/i])} / {igpNumber}</span></div>
                    <div className="detail-row customer-box"><span className="detail-label">Customer / Contract</span><span className="detail-value">{getCell(row, [/^processing\s*customer$/i])} / {contractNumber}</span></div>
                  </div>
                  <div className="stock-amounts">
                    <div className="amount-box rolls-box"><span className="amount-label">Type / Rolls-Bales</span><strong className="amount-value">{getColumnCell(row, data.columns.type)} / {getStockCell(row, data.stockColumns?.rolls)}</strong><span className="amount-unit">Type / R-B</span></div>
                    <div className="amount-box a-grade-box"><span className="amount-label">A-Grade</span><strong className="amount-value">{getStockCell(row, data.stockColumns?.aGrade)}</strong><span className="amount-unit">MTR</span></div>
                  </div>
                  <div className="remarks-box"><span className="remarks-label">Remarks</span><p>{getCell(row, [/^remarks?$/i, /remarks?/i])}</p></div>
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