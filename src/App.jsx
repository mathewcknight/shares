import React, { useState, useCallback } from 'react'

// ---------------------------------------------------------------------------
// Parsing helpers
// ---------------------------------------------------------------------------

function getCellValue(sheet, row, col) {
  // row and col are 0-based
  const XLSX = window.XLSX
  const addr = XLSX.utils.encode_cell({ r: row, c: col })
  const cell = sheet[addr]
  return cell ? cell.v : null
}

function parseSummarySheet(sheet) {
  // Row 12 col F = r:11, c:5
  // Row 13 col F = r:12, c:5
  // Row 15 col F = r:14, c:5
  return {
    cashPosition:        getCellValue(sheet, 11, 5),
    domesticHoldingsValue: getCellValue(sheet, 12, 5),
    unsettledTrades:     getCellValue(sheet, 14, 5),
  }
}

function parseDomesticHoldingsSheet(sheet) {
  // Headers row 2 (index 1), data from row 3 (index 2)
  // Columns (0-based):
  //   0: Account, 1: Description, 2: Code, 3: Quantity, 4: Portfolio %,
  //   5: skip,    6: Market Price, 7: Market Price As At, 8: Average Price(1),
  //   9: skip,   10: Market Value, 11: Gain/Loss, 12: Gain/Loss %
  const rows = []
  let rowIndex = 2 // start at row 3 (0-based index 2)

  while (true) {
    const code        = getCellValue(sheet, rowIndex, 2)
    const description = getCellValue(sheet, rowIndex, 1)

    // Stop conditions
    if (code === null || code === undefined || code === '') break
    if (typeof description === 'string' && description.includes('Important Information')) break

    rows.push({
      account:          getCellValue(sheet, rowIndex, 0),
      description,
      code,
      quantity:         getCellValue(sheet, rowIndex, 3),
      portfolioPct:     getCellValue(sheet, rowIndex, 4),
      marketPrice:      getCellValue(sheet, rowIndex, 6),
      marketPriceAsAt:  getCellValue(sheet, rowIndex, 7),
      averagePrice:     getCellValue(sheet, rowIndex, 8),
      marketValue:      getCellValue(sheet, rowIndex, 10),
      gainLoss:         getCellValue(sheet, rowIndex, 11),
      gainLossPct:      getCellValue(sheet, rowIndex, 12),
    })

    rowIndex++
  }

  return rows
}

function parseWorkbook(arrayBuffer) {
  const XLSX = window.XLSX
  const workbook = XLSX.read(arrayBuffer, { type: 'array', cellDates: true })

  const summarySheet  = workbook.Sheets['Summary']
  const holdingsSheet = workbook.Sheets['Domestic Holdings']

  if (!summarySheet)  console.warn('Sheet "Summary" not found. Available:', workbook.SheetNames)
  if (!holdingsSheet) console.warn('Sheet "Domestic Holdings" not found. Available:', workbook.SheetNames)

  const summary  = summarySheet  ? parseSummarySheet(summarySheet)           : null
  const holdings = holdingsSheet ? parseDomesticHoldingsSheet(holdingsSheet) : []

  console.group('=== Stage 2 Parse Results ===')
  console.log('Summary:', summary)
  console.log('Holdings row count:', holdings.length)
  console.log('Holdings sample (first 3):', holdings.slice(0, 3))
  console.groupEnd()

  return { summary, holdings }
}

// ---------------------------------------------------------------------------
// App
// ---------------------------------------------------------------------------

export default function App() {
  const [fileName, setFileName]   = useState(null)
  const [isDragging, setIsDragging] = useState(false)
  const [parsed, setParsed]       = useState(null)
  const [parseError, setParseError] = useState(null)

  const handleFile = (file) => {
    if (!file || !file.name.endsWith('.xlsx')) return
    setFileName(file.name)
    setParseError(null)

    const reader = new FileReader()
    reader.onload = (e) => {
      try {
        const data = parseWorkbook(e.target.result)
        setParsed(data)
      } catch (err) {
        console.error('Parse error:', err)
        setParseError(err.message)
      }
    }
    reader.readAsArrayBuffer(file)
  }

  const handleDrop = useCallback((e) => {
    e.preventDefault()
    setIsDragging(false)
    handleFile(e.dataTransfer.files[0])
  }, [])

  const handleDragOver = useCallback((e) => {
    e.preventDefault()
    setIsDragging(true)
  }, [])

  const handleDragLeave = useCallback(() => {
    setIsDragging(false)
  }, [])

  const handleFileInput = (e) => handleFile(e.target.files[0])

  const parseSucceeded = parsed !== null

  return (
    <div className="min-h-screen bg-gray-950 text-gray-100" style={{ backgroundColor: '#0a0a0f' }}>
      {/* Header */}
      <header className="border-b border-gray-800 bg-gray-900 px-6 py-4">
        <div className="mx-auto max-w-7xl flex items-center gap-3">
          <div className="h-8 w-8 rounded-lg bg-blue-600 flex items-center justify-center text-sm font-bold">P</div>
          <h1 className="text-lg font-semibold tracking-wide text-white">Portfolio Dashboard</h1>
        </div>
      </header>

      {/* Main content */}
      <main className="flex min-h-[calc(100vh-57px)] items-center justify-center p-8">
        <div className="w-full max-w-lg space-y-6 text-center">
          <div>
            <h2 className="text-3xl font-bold text-white">Upload Your Portfolio</h2>
            <p className="mt-2 text-gray-400">Import your nabtrade Excel export to analyse your holdings</p>
          </div>

          {/* Drop zone */}
          <div
            onDrop={handleDrop}
            onDragOver={handleDragOver}
            onDragLeave={handleDragLeave}
            className="rounded-2xl border-2 border-dashed p-14 transition-all duration-200 cursor-default"
            style={{
              borderColor: isDragging ? '#3b82f6' : '#374151',
              backgroundColor: isDragging ? 'rgba(59,130,246,0.08)' : '#111827',
            }}
          >
            <div className="space-y-5">
              <div className="mx-auto flex h-16 w-16 items-center justify-center rounded-2xl bg-gray-800 text-3xl">
                📊
              </div>
              <div>
                <p className="text-base text-gray-300 font-medium">Drag &amp; drop your .xlsx file here</p>
                <p className="mt-1 text-sm text-gray-500">Supports nabtrade portfolio exports</p>
              </div>
              <div className="flex items-center gap-4">
                <div className="flex-1 border-t border-gray-700" />
                <span className="text-xs text-gray-500 uppercase tracking-widest">or</span>
                <div className="flex-1 border-t border-gray-700" />
              </div>
              <label className="inline-block cursor-pointer rounded-lg bg-blue-600 px-8 py-2.5 text-sm font-semibold text-white transition-colors hover:bg-blue-500 active:bg-blue-700">
                Browse File
                <input
                  type="file"
                  accept=".xlsx"
                  className="hidden"
                  onChange={handleFileInput}
                />
              </label>
            </div>
          </div>

          {/* Status */}
          {fileName && !parseError && (
            <div
              className="flex items-center gap-3 rounded-xl border px-4 py-3 text-left"
              style={{
                backgroundColor: parseSucceeded ? 'rgba(5,46,22,0.4)' : 'rgba(23,37,84,0.4)',
                borderColor:     parseSucceeded ? '#14532d'            : '#1e3a8a',
              }}
            >
              <div
                className="flex h-6 w-6 flex-shrink-0 items-center justify-center rounded-full text-xs text-white"
                style={{ backgroundColor: parseSucceeded ? '#16a34a' : '#2563eb' }}
              >
                {parseSucceeded ? '✓' : '…'}
              </div>
              <div>
                <p className="text-sm font-medium" style={{ color: parseSucceeded ? '#86efac' : '#93c5fd' }}>
                  {fileName}
                </p>
                <p className="text-xs text-gray-500">
                  {parseSucceeded
                    ? `Parsed — ${parsed.holdings.length} holdings · check console for details`
                    : 'Parsing…'}
                </p>
              </div>
            </div>
          )}

          {parseError && (
            <div
              className="flex items-center gap-3 rounded-xl border px-4 py-3 text-left"
              style={{ backgroundColor: 'rgba(69,10,10,0.4)', borderColor: '#7f1d1d' }}
            >
              <div className="flex h-6 w-6 flex-shrink-0 items-center justify-center rounded-full bg-red-600 text-xs text-white">!</div>
              <div>
                <p className="text-sm font-medium text-red-300">Parse failed</p>
                <p className="text-xs text-gray-500">{parseError}</p>
              </div>
            </div>
          )}

          <p className="text-xs text-gray-600">All processing happens locally in your browser. No data is uploaded.</p>
        </div>
      </main>
    </div>
  )
}
