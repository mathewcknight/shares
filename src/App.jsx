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
    cashPosition:          getCellValue(sheet, 11, 5),
    domesticHoldingsValue: getCellValue(sheet, 12, 5),
    unsettledTrades:       getCellValue(sheet, 14, 5),
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
      account:         getCellValue(sheet, rowIndex, 0),
      description,
      code,
      quantity:        getCellValue(sheet, rowIndex, 3),
      portfolioPct:    getCellValue(sheet, rowIndex, 4),
      marketPrice:     getCellValue(sheet, rowIndex, 6),
      marketPriceAsAt: getCellValue(sheet, rowIndex, 7),
      averagePrice:    getCellValue(sheet, rowIndex, 8),
      marketValue:     getCellValue(sheet, rowIndex, 10),
      gainLoss:        getCellValue(sheet, rowIndex, 11),
      gainLossPct:     getCellValue(sheet, rowIndex, 12),
    })

    rowIndex++
  }

  return rows
}

// Returns true if any string cell value in a row (across given col range)
// contains a stop sentinel for the Transactions sheet.
function rowHasTransactionSentinel(sheet, rowIndex, colCount) {
  const XLSX = window.XLSX
  for (let c = 0; c < colCount; c++) {
    const addr = XLSX.utils.encode_cell({ r: rowIndex, c })
    const cell = sheet[addr]
    if (!cell || typeof cell.v !== 'string') continue
    if (cell.v.includes('Important Information')) return true
    if (cell.v.startsWith('(1)')) return true
  }
  return false
}

function parseTransactionsSheet(sheet) {
  // Headers row 2 (index 1), data from row 3 (index 2)
  // Columns (0-based):
  //   0: Account,  1: Description,      2: Code,             3: Date,
  //   4: Movement Type,                 5: Confirmation Number,
  //   6: Exchange Currency,             7: Quantity,
  //   8: Transaction Price,             9: Value,
  //  10: Brokerage, 11: Other Fees,    12: Average Price,
  //  13: Multiplier,                   14: Settlement Amount (AUD)
  const rows = []
  let rowIndex = 2

  while (true) {
    const code = getCellValue(sheet, rowIndex, 2)

    // Stop if Code is null/empty
    if (code === null || code === undefined || code === '') break

    // Stop if any cell in the row contains a sentinel string
    if (rowHasTransactionSentinel(sheet, rowIndex, 15)) break

    rows.push({
      account:            getCellValue(sheet, rowIndex, 0),
      description:        getCellValue(sheet, rowIndex, 1),
      code,
      date:               getCellValue(sheet, rowIndex, 3),
      movementType:       getCellValue(sheet, rowIndex, 4),
      confirmationNumber: getCellValue(sheet, rowIndex, 5),
      exchangeCurrency:   getCellValue(sheet, rowIndex, 6),
      quantity:           getCellValue(sheet, rowIndex, 7),
      transactionPrice:   getCellValue(sheet, rowIndex, 8),
      value:              getCellValue(sheet, rowIndex, 9),
      brokerage:          getCellValue(sheet, rowIndex, 10),
      otherFees:          getCellValue(sheet, rowIndex, 11),
      averagePrice:       getCellValue(sheet, rowIndex, 12),
      multiplier:         getCellValue(sheet, rowIndex, 13),
      settlementAmount:   getCellValue(sheet, rowIndex, 14),
    })

    rowIndex++
  }

  return rows
}

function parseDividendsOrInterestSheet(sheet) {
  // Headers row 2 (index 1), data from row 3 (index 2)
  // Columns (0-based): 0: Account, 1: Date, 2: Description, 3: Value
  // Stop at any row whose Description contains "Important Information"
  const rows = []
  let rowIndex = 2

  while (true) {
    const description = getCellValue(sheet, rowIndex, 2)
    const account     = getCellValue(sheet, rowIndex, 0)

    // Stop if both account and description are empty (blank row)
    if ((account === null || account === '') &&
        (description === null || description === '')) break

    // Stop at Important Information sentinel
    if (typeof description === 'string' && description.includes('Important Information')) break
    if (typeof account === 'string' && account.includes('Important Information')) break

    rows.push({
      account,
      date:        getCellValue(sheet, rowIndex, 1),
      description,
      value:       getCellValue(sheet, rowIndex, 3),
    })

    rowIndex++
  }

  return rows
}

function parseWorkbook(arrayBuffer) {
  const XLSX = window.XLSX
  const workbook = XLSX.read(arrayBuffer, { type: 'array', cellDates: true })

  const sheetNames = workbook.SheetNames
  const summarySheet      = workbook.Sheets['Summary']
  const holdingsSheet     = workbook.Sheets['Domestic Holdings']
  const transactionsSheet = workbook.Sheets['Domestic Portfolio Transactions']
  const dividendsSheet    = workbook.Sheets['Domestic Dividends']
  const interestSheet     = workbook.Sheets['Interest']

  if (!summarySheet)      console.warn('Sheet "Summary" not found. Available:', sheetNames)
  if (!holdingsSheet)     console.warn('Sheet "Domestic Holdings" not found. Available:', sheetNames)
  if (!transactionsSheet) console.warn('Sheet "Domestic Portfolio Transactions" not found. Available:', sheetNames)
  if (!dividendsSheet)    console.warn('Sheet "Domestic Dividends" not found. Available:', sheetNames)
  if (!interestSheet)     console.warn('Sheet "Interest" not found. Available:', sheetNames)

  const summary      = summarySheet      ? parseSummarySheet(summarySheet)               : null
  const holdings     = holdingsSheet     ? parseDomesticHoldingsSheet(holdingsSheet)     : []
  const transactions = transactionsSheet ? parseTransactionsSheet(transactionsSheet)     : []
  const dividends    = dividendsSheet    ? parseDividendsOrInterestSheet(dividendsSheet) : []
  const interest     = interestSheet     ? parseDividendsOrInterestSheet(interestSheet)  : []

  console.group('=== Stage 3 Parse Results ===')
  console.log('Summary:', summary)
  console.log('Holdings row count:', holdings.length)
  console.log('Holdings sample (first 3):', holdings.slice(0, 3))
  console.log('Transactions row count:', transactions.length)
  console.log('Transactions sample (first 3):', transactions.slice(0, 3))
  console.log('Dividends row count:', dividends.length)
  console.log('Dividends sample (first 3):', dividends.slice(0, 3))
  console.log('Interest row count:', interest.length)
  console.log('Interest sample (first 3):', interest.slice(0, 3))
  console.groupEnd()

  return { summary, holdings, transactions, dividends, interest }
}

// ---------------------------------------------------------------------------
// Stage 4: Core calculations
// ---------------------------------------------------------------------------

// Movement type classification. nabtrade uses strings like "Buy", "Sell",
// "DRP" (Dividend Reinvestment Plan), etc.
function isBuyType(movementType) {
  if (!movementType) return false
  const t = String(movementType).toLowerCase()
  return t.includes('buy') || t.includes('drp') || t.includes('dividend reinvestment')
}

function isSellType(movementType) {
  if (!movementType) return false
  const t = String(movementType).toLowerCase()
  return t.includes('sell')
}

// Convert an Excel serial date or JS Date to a comparable number for sorting.
function toSortableDate(value) {
  if (!value) return 0
  if (value instanceof Date) return value.getTime()
  if (typeof value === 'number') return value // Excel serial — ordering is preserved as-is
  return new Date(value).getTime()
}

/**
 * calculateMetrics({ holdings, transactions, dividends })
 *
 * Returns:
 *   capitalInvested   – net cash deployed (sum of BUY+DRP settlements minus sum of SELL settlements)
 *   unrealisedPnL     – sum of Gain/Loss column from Holdings sheet
 *   realisedPnL       – realised gain/loss via running-average cost basis across all sells
 *   totalDividends    – sum of all Domestic Dividends values
 *   totalReturnDollars – unrealisedPnL + realisedPnL + totalDividends
 *   totalReturnPct    – totalReturnDollars / capitalInvested (null if capitalInvested === 0)
 *   costBasisByTicker – remaining average-cost basis per ticker after all transactions
 */
function calculateMetrics({ holdings, transactions, dividends }) {
  // Process transactions in chronological order so the running average cost
  // basis reflects the actual sequence of trades.
  const sorted = [...transactions].sort(
    (a, b) => toSortableDate(a.date) - toSortableDate(b.date)
  )

  // ------------------------------------------------------------------
  // Capital Invested
  //   = sum of abs(BUY/DRP settlementAmounts)
  //   - sum of abs(SELL settlementAmounts)
  // Settlement amounts in nabtrade exports are typically negative for
  // purchases (cash out) and positive for sales (cash in); abs() normalises
  // both so we work with magnitudes throughout.
  // ------------------------------------------------------------------
  let capitalInvested = 0
  for (const t of sorted) {
    const settlement = Math.abs(Number(t.settlementAmount) || 0)
    if (isBuyType(t.movementType))       capitalInvested += settlement
    else if (isSellType(t.movementType)) capitalInvested -= settlement
  }

  // ------------------------------------------------------------------
  // Unrealised P&L — straight sum of Gain/Loss from Holdings sheet
  // ------------------------------------------------------------------
  const unrealisedPnL = holdings.reduce(
    (sum, h) => sum + (Number(h.gainLoss) || 0), 0
  )

  // ------------------------------------------------------------------
  // Realised P&L — running-average cost basis per ticker
  //
  // State per ticker: { qty, totalCost }
  //   qty       = shares currently held (decreases on SELL)
  //   totalCost = total cost basis of those shares (avgCost = totalCost/qty)
  //
  // On BUY/DRP:
  //   qty       += trade qty
  //   totalCost += abs(settlementAmount)   ← includes brokerage
  //
  // On SELL:
  //   avgCost         = totalCost / qty
  //   realisedPnL    += abs(sellSettlement) - avgCost * sellQty
  //   qty            -= sellQty
  //   totalCost       = avgCost * remainingQty   ← avg cost unchanged
  // ------------------------------------------------------------------
  const costBasisByTicker = {}
  let realisedPnL = 0

  for (const t of sorted) {
    const code       = t.code
    const tradeQty   = Math.abs(Number(t.quantity)         || 0)
    const settlement = Math.abs(Number(t.settlementAmount) || 0)

    if (!costBasisByTicker[code]) costBasisByTicker[code] = { qty: 0, totalCost: 0 }
    const pos = costBasisByTicker[code]

    if (isBuyType(t.movementType)) {
      pos.qty       += tradeQty
      pos.totalCost += settlement

    } else if (isSellType(t.movementType)) {
      if (pos.qty > 0) {
        const avgCost      = pos.totalCost / pos.qty
        realisedPnL       += settlement - avgCost * tradeQty
        const remainingQty = Math.max(0, pos.qty - tradeQty)
        pos.qty            = remainingQty
        pos.totalCost      = avgCost * remainingQty
      } else {
        // Selling a position with no tracked cost basis (data before window)
        console.warn(`[Stage 4] SELL for ${code} but no cost basis tracked — skipping realised P&L for this trade`)
      }
    }
  }

  // ------------------------------------------------------------------
  // Total Dividends
  // ------------------------------------------------------------------
  const totalDividends = dividends.reduce(
    (sum, d) => sum + (Number(d.value) || 0), 0
  )

  // ------------------------------------------------------------------
  // Total Return
  // ------------------------------------------------------------------
  const totalReturnDollars = unrealisedPnL + realisedPnL + totalDividends
  const totalReturnPct     = capitalInvested !== 0
    ? totalReturnDollars / capitalInvested
    : null

  return {
    capitalInvested,
    unrealisedPnL,
    realisedPnL,
    totalDividends,
    totalReturnDollars,
    totalReturnPct,
    costBasisByTicker,
  }
}

function logMetrics(metrics) {
  const pct = (v) => v !== null ? (v * 100).toFixed(4) + '%' : 'n/a'
  const aud = (v) => '$' + v.toLocaleString('en-AU', { minimumFractionDigits: 2, maximumFractionDigits: 2 })

  console.group('=== Stage 4 Calculated Metrics ===')
  console.log('Capital Invested:    ', aud(metrics.capitalInvested))
  console.log('Unrealised P&L:      ', aud(metrics.unrealisedPnL))
  console.log('Realised P&L:        ', aud(metrics.realisedPnL))
  console.log('Total Dividends:     ', aud(metrics.totalDividends))
  console.log('Total Return $:      ', aud(metrics.totalReturnDollars))
  console.log('Total Return %:      ', pct(metrics.totalReturnPct))
  console.log('Cost basis by ticker:', metrics.costBasisByTicker)
  console.groupEnd()
}

// ---------------------------------------------------------------------------
// Stage 5: formatting helpers + summary card components
// ---------------------------------------------------------------------------

function fmtAUD(value) {
  if (value === null || value === undefined) return '—'
  const n = Number(value)
  if (!isFinite(n)) return '—'
  const abs = Math.abs(n)
  const str = abs.toLocaleString('en-AU', { minimumFractionDigits: 2, maximumFractionDigits: 2 })
  return (n < 0 ? '-$' : '$') + str
}

function fmtPct(value) {
  if (value === null || value === undefined) return '—'
  const n = Number(value)
  if (!isFinite(n)) return '—'
  return (n >= 0 ? '+' : '') + (n * 100).toFixed(2) + '%'
}

function pnlColor(value) {
  const n = Number(value)
  if (!isFinite(n)) return '#f3f4f6'
  return n >= 0 ? '#4ade80' : '#f87171'
}

function SummaryCard({ label, value, valueColor, sub, subColor }) {
  return (
    <div className="rounded-2xl border border-gray-800 bg-gray-900 p-5 flex-1 min-w-0">
      <p className="text-xs font-medium uppercase tracking-widest text-gray-500 truncate">{label}</p>
      <p className="mt-2 text-2xl font-bold truncate" style={{ color: valueColor || '#f3f4f6' }}>{value}</p>
      {sub && (
        <p className="mt-1 text-sm font-semibold" style={{ color: subColor || '#9ca3af' }}>{sub}</p>
      )}
    </div>
  )
}

function SummaryCards({ summary, metrics }) {
  const totalMarketValue =
    (Number(summary?.domesticHoldingsValue) || 0) +
    (Number(summary?.cashPosition)          || 0)

  return (
    <div className="flex gap-4">
      <SummaryCard
        label="Total Market Value"
        value={fmtAUD(totalMarketValue)}
      />
      <SummaryCard
        label="Capital Invested"
        value={fmtAUD(metrics.capitalInvested)}
      />
      <SummaryCard
        label="Unrealised P&L"
        value={fmtAUD(metrics.unrealisedPnL)}
        valueColor={pnlColor(metrics.unrealisedPnL)}
      />
      <SummaryCard
        label="Realised P&L"
        value={fmtAUD(metrics.realisedPnL)}
        valueColor={pnlColor(metrics.realisedPnL)}
      />
      <SummaryCard
        label="Total Return"
        value={fmtAUD(metrics.totalReturnDollars)}
        valueColor={pnlColor(metrics.totalReturnDollars)}
        sub={fmtPct(metrics.totalReturnPct)}
        subColor={pnlColor(metrics.totalReturnPct)}
      />
    </div>
  )
}

// ---------------------------------------------------------------------------
// App
// ---------------------------------------------------------------------------

export default function App() {
  const [fileName, setFileName]     = useState(null)
  const [isDragging, setIsDragging] = useState(false)
  const [parsed, setParsed]         = useState(null)
  const [parseError, setParseError] = useState(null)

  const handleFile = (file) => {
    if (!file || !file.name.endsWith('.xlsx')) return
    setFileName(file.name)
    setParseError(null)

    const reader = new FileReader()
    reader.onload = (e) => {
      try {
        const data = parseWorkbook(e.target.result)
        const metrics = calculateMetrics(data)
        logMetrics(metrics)
        setParsed({ ...data, metrics })
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

  return (
    <div className="min-h-screen bg-gray-950 text-gray-100" style={{ backgroundColor: '#0a0a0f' }}>
      {/* Header */}
      <header className="border-b border-gray-800 bg-gray-900 px-6 py-4">
        <div className="mx-auto max-w-7xl flex items-center gap-3">
          <div className="h-8 w-8 rounded-lg bg-blue-600 flex items-center justify-center text-sm font-bold">P</div>
          <h1 className="text-lg font-semibold tracking-wide text-white">Portfolio Dashboard</h1>
        </div>
      </header>

      {parsed ? (
        /* Dashboard — replaces upload zone once a file is loaded */
        <main className="mx-auto max-w-7xl px-6 py-8 space-y-6">
          <SummaryCards summary={parsed.summary} metrics={parsed.metrics} />
        </main>
      ) : (
        /* Upload screen */
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

            {/* Parse error */}
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
      )}
    </div>
  )
}
