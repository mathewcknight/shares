import React, { useState, useCallback } from 'react'
import {
  ComposedChart, AreaChart, Area,
  Line, Scatter, XAxis, YAxis, CartesianGrid,
  Tooltip, ResponsiveContainer, Legend, Brush,
} from 'recharts'

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
  const costBasisByTicker  = {}
  const realisedByTicker   = {} // { [code]: { pnl, totalInvested } }
  let realisedPnL = 0

  for (const t of sorted) {
    const code       = t.code
    const tradeQty   = Math.abs(Number(t.quantity)         || 0)
    const settlement = Math.abs(Number(t.settlementAmount) || 0)

    if (!costBasisByTicker[code]) costBasisByTicker[code] = { qty: 0, totalCost: 0 }
    if (!realisedByTicker[code])  realisedByTicker[code]  = { pnl: 0, totalInvested: 0, totalProceeds: 0 }
    const pos = costBasisByTicker[code]
    const rec = realisedByTicker[code]

    if (isBuyType(t.movementType)) {
      pos.qty            += tradeQty
      pos.totalCost      += settlement
      rec.totalInvested  += settlement

    } else if (isSellType(t.movementType)) {
      if (pos.qty > 0) {
        const avgCost      = pos.totalCost / pos.qty
        const tradePnL     = settlement - avgCost * tradeQty
        realisedPnL       += tradePnL
        rec.pnl            += tradePnL
        rec.totalProceeds  += settlement
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
    realisedByTicker,
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
// Stage 10: Holdings Table
// ---------------------------------------------------------------------------

function HoldingsTable({ holdings, metrics, transactions }) {
  const [sort, setSort] = React.useState({ col: 'returnPct', dir: 'desc' })

  const { costBasisByTicker, realisedByTicker } = metrics

  // description fallback: first transaction description per code
  const txDescMap = {}
  for (const t of transactions) {
    if (!txDescMap[t.code] && t.description) txDescMap[t.code] = t.description
  }

  const holdingsMap = {}
  for (const h of holdings) {
    if (h.code) holdingsMap[h.code] = h
  }

  // Union all codes seen in either transactions or current holdings
  const allCodes = new Set([
    ...Object.keys(realisedByTicker),
    ...Object.keys(holdingsMap),
  ])

  const rows = [...allCodes].map(code => {
    const isOpen = (costBasisByTicker[code]?.qty ?? 0) > 0.001
    const h = holdingsMap[code]
    const r = realisedByTicker[code]

    const rawDesc   = h?.description ?? txDescMap[code] ?? code
    const description = String(rawDesc).split(/\s+/).slice(0, 4).join(' ')

    const capitalInvested = r?.totalInvested ?? 0
    const currentValue    = isOpen ? (Number(h?.marketValue) || 0) : (r?.totalProceeds ?? 0)
    const pnl             = isOpen ? (Number(h?.gainLoss)    || 0) : (r?.pnl ?? 0)
    const returnPct       = capitalInvested > 0 ? pnl / capitalInvested : null
    const quantity        = isOpen ? (Number(h?.quantity)    || 0) : 0

    return { code, description, status: isOpen ? 'Open' : 'Closed', capitalInvested, currentValue, pnl, returnPct, quantity }
  })

  const handleSort = col =>
    setSort(s => ({ col, dir: s.col === col && s.dir === 'desc' ? 'asc' : 'desc' }))

  const sorted = [...rows].sort((a, b) => {
    const mul = sort.dir === 'desc' ? -1 : 1
    const av  = a[sort.col] ?? (sort.dir === 'desc' ? -Infinity : Infinity)
    const bv  = b[sort.col] ?? (sort.dir === 'desc' ? -Infinity : Infinity)
    return typeof av === 'string' ? mul * av.localeCompare(bv) : mul * (av - bv)
  })

  const Th = ({ col, children, right }) => (
    <th
      onClick={() => handleSort(col)}
      className={`py-2 pr-3 text-xs font-medium uppercase tracking-wider cursor-pointer select-none ${right ? 'text-right' : 'text-left'}`}
      style={{ color: sort.col === col ? '#d1d5db' : '#6b7280' }}
    >
      {children}
      <span style={{ marginLeft: 3, opacity: sort.col === col ? 1 : 0.3 }}>
        {sort.col === col ? (sort.dir === 'desc' ? '↓' : '↑') : '↕'}
      </span>
    </th>
  )

  return (
    <div className="rounded-2xl border border-gray-800 bg-gray-900 p-5 flex flex-col gap-3">
      <h3 className="text-sm font-semibold uppercase tracking-widest text-gray-400">All Holdings</h3>
      <div className="overflow-y-auto" style={{ maxHeight: 480 }}>
        <table className="w-full text-sm border-collapse">
          <thead className="sticky top-0 bg-gray-900 z-10">
            <tr>
              <Th col="code">Code</Th>
              <Th col="description">Description</Th>
              <Th col="status">Status</Th>
              <Th col="capitalInvested" right>Capital Invested</Th>
              <Th col="currentValue"    right>Current Value</Th>
              <Th col="pnl"             right>P&amp;L $</Th>
              <Th col="returnPct"       right>Return %</Th>
              <Th col="quantity"        right>Quantity</Th>
            </tr>
          </thead>
          <tbody>
            {sorted.map(r => (
              <tr key={r.code} style={{ backgroundColor: rowBg(r.pnl) }}>
                <td className="py-1.5 pr-3 font-mono font-semibold text-gray-200">{r.code}</td>
                <td className="py-1.5 pr-3 text-gray-400" style={{ maxWidth: 160, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{r.description}</td>
                <td className="py-1.5 pr-3">
                  <span className="text-xs px-1.5 py-0.5 rounded font-medium" style={{
                    background: r.status === 'Open' ? 'rgba(59,130,246,0.15)' : 'rgba(75,85,99,0.3)',
                    color:      r.status === 'Open' ? '#93c5fd' : '#9ca3af',
                  }}>{r.status}</span>
                </td>
                <td className="py-1.5 pr-3 text-right tabular-nums text-gray-300">{fmtAUD(r.capitalInvested)}</td>
                <td className="py-1.5 pr-3 text-right tabular-nums text-gray-300">{fmtAUD(r.currentValue)}</td>
                <td className="py-1.5 pr-3 text-right tabular-nums" style={{ color: r.pnl >= 0 ? '#4ade80' : '#f87171' }}>{fmtAUD(r.pnl)}</td>
                <td className="py-1.5 pr-3 text-right tabular-nums" style={{ color: (r.returnPct ?? 0) >= 0 ? '#4ade80' : '#f87171' }}>{fmtPct(r.returnPct)}</td>
                <td className="py-1.5 text-right tabular-nums text-gray-300">{r.quantity > 0 ? r.quantity.toLocaleString('en-AU') : '—'}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  )
}

// ---------------------------------------------------------------------------
// Stage 6: Position tables (Closed & Open)
// ---------------------------------------------------------------------------

function SortToggle({ value, onChange }) {
  return (
    <div className="flex gap-1">
      {['$', '%'].map(v => (
        <button
          key={v}
          onClick={() => onChange(v)}
          className="px-2.5 py-0.5 rounded text-xs font-semibold transition-colors"
          style={{
            background: value === v ? '#3b82f6' : '#1f2937',
            color:      value === v ? '#fff'    : '#6b7280',
          }}
        >
          {v}
        </button>
      ))}
    </div>
  )
}

function rowBg(pnl) {
  const n = Number(pnl)
  if (!isFinite(n) || n === 0) return 'transparent'
  return n > 0 ? 'rgba(74,222,128,0.07)' : 'rgba(248,113,113,0.07)'
}

function ClosedPositionsTable({ metrics }) {
  const [sortBy, setSortBy] = React.useState('$')

  const { costBasisByTicker, realisedByTicker } = metrics

  // Closed = qty rounds to zero after all transactions
  const rows = Object.entries(realisedByTicker)
    .filter(([code]) => (costBasisByTicker[code]?.qty ?? 0) < 0.001)
    .map(([code, { pnl, totalInvested }]) => ({
      code,
      pnl,
      pct: totalInvested > 0 ? pnl / totalInvested : null,
    }))
    .sort((a, b) => sortBy === '$'
      ? b.pnl - a.pnl
      : (b.pct ?? -Infinity) - (a.pct ?? -Infinity))

  return (
    <div className="rounded-2xl border border-gray-800 bg-gray-900 p-5 flex flex-col gap-3 min-w-0">
      <div className="flex items-center justify-between">
        <h3 className="text-sm font-semibold uppercase tracking-widest text-gray-400">Closed Positions</h3>
        <SortToggle value={sortBy} onChange={setSortBy} />
      </div>
      {rows.length === 0 ? (
        <p className="text-xs text-gray-600 italic">No closed positions found.</p>
      ) : (
        <div className="overflow-y-auto" style={{ maxHeight: 380 }}>
          <table className="w-full text-sm border-collapse">
            <thead>
              <tr className="text-xs uppercase tracking-wider text-gray-500">
                <th className="text-left py-1.5 pr-4 font-medium">Code</th>
                <th className="text-right py-1.5 pr-4 font-medium">Realised P&amp;L $</th>
                <th className="text-right py-1.5 font-medium">Realised P&amp;L %</th>
              </tr>
            </thead>
            <tbody>
              {rows.map(r => (
                <tr key={r.code} style={{ backgroundColor: rowBg(r.pnl) }}>
                  <td className="py-1.5 pr-4 font-mono font-semibold text-gray-200">{r.code}</td>
                  <td className="py-1.5 pr-4 text-right tabular-nums" style={{ color: r.pnl >= 0 ? '#4ade80' : '#f87171' }}>
                    {fmtAUD(r.pnl)}
                  </td>
                  <td className="py-1.5 text-right tabular-nums" style={{ color: r.pnl >= 0 ? '#4ade80' : '#f87171' }}>
                    {fmtPct(r.pct)}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  )
}

function OpenPositionsTable({ holdings }) {
  const [sortBy, setSortBy] = React.useState('$')

  const rows = [...holdings]
    .filter(h => h.code && h.gainLoss != null)
    .map(h => ({
      code: h.code,
      pnl:  Number(h.gainLoss)    || 0,
      pct:  h.gainLossPct != null ? Number(h.gainLossPct) : null,
    }))
    .sort((a, b) => sortBy === '$'
      ? b.pnl - a.pnl
      : (b.pct ?? -Infinity) - (a.pct ?? -Infinity))

  return (
    <div className="rounded-2xl border border-gray-800 bg-gray-900 p-5 flex flex-col gap-3 min-w-0">
      <div className="flex items-center justify-between">
        <h3 className="text-sm font-semibold uppercase tracking-widest text-gray-400">Open Positions</h3>
        <SortToggle value={sortBy} onChange={setSortBy} />
      </div>
      {rows.length === 0 ? (
        <p className="text-xs text-gray-600 italic">No open positions found.</p>
      ) : (
        <div className="overflow-y-auto" style={{ maxHeight: 380 }}>
          <table className="w-full text-sm border-collapse">
            <thead>
              <tr className="text-xs uppercase tracking-wider text-gray-500">
                <th className="text-left py-1.5 pr-4 font-medium">Code</th>
                <th className="text-right py-1.5 pr-4 font-medium">Unrealised P&amp;L $</th>
                <th className="text-right py-1.5 font-medium">Unrealised P&amp;L %</th>
              </tr>
            </thead>
            <tbody>
              {rows.map(r => (
                <tr key={r.code} style={{ backgroundColor: rowBg(r.pnl) }}>
                  <td className="py-1.5 pr-4 font-mono font-semibold text-gray-200">{r.code}</td>
                  <td className="py-1.5 pr-4 text-right tabular-nums" style={{ color: r.pnl >= 0 ? '#4ade80' : '#f87171' }}>
                    {fmtAUD(r.pnl)}
                  </td>
                  <td className="py-1.5 text-right tabular-nums" style={{ color: r.pnl >= 0 ? '#4ade80' : '#f87171' }}>
                    {fmtPct(r.pct)}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  )
}

// ---------------------------------------------------------------------------
// Stage 7: Cumulative Capital Deployed chart
// ---------------------------------------------------------------------------

function fmtDate(value) {
  if (!value) return '—'
  if (value instanceof Date) {
    return value.toLocaleDateString('en-AU', { day: '2-digit', month: '2-digit', year: 'numeric' })
  }
  // Excel serial number (days since 1900-01-01, with Lotus 1-2-3 leap-year bug)
  if (typeof value === 'number') {
    const d = new Date(Math.round((value - 25569) * 86400 * 1000))
    return d.toLocaleDateString('en-AU', { day: '2-digit', month: '2-digit', year: 'numeric' })
  }
  return String(value)
}

// Fine-grained movement-type classifiers used only by the chart.
// (isBuyType / isSellType in the summary section remain unchanged.)
function isChartBuyOnly(mt) {
  if (!mt) return false
  const t = String(mt).toLowerCase()
  return t.includes('buy') && !t.includes('drp') && !t.includes('reinvestment')
}
function isChartDRP(mt) {
  if (!mt) return false
  const t = String(mt).toLowerCase()
  return t.includes('drp') || t.includes('dividend reinvestment')
}
function isChartCashDiv(mt) {
  if (!mt) return false
  const t = String(mt).toLowerCase()
  return t.includes('dividend') && !t.includes('reinvestment') && !t.includes('drp')
}
function isChartSell(mt) {
  if (!mt) return false
  return String(mt).toLowerCase().includes('sell')
}

/**
 * Returns { lineData, buyPoints, drpPoints, cashDivPoints, sellPoints }.
 * lineData       – { x (ms), y (cumulative AUD) } — steps change on BUY/DRP/SELL only
 * *Points        – per-transaction markers with metadata + normalised size
 * Cash dividends don't move cumulative capital; their marker y = current cumulative.
 */
function buildCapitalDeployedSeries(transactions) {
  const relevant = transactions
    .filter(t => {
      const mt = t.movementType
      return isChartBuyOnly(mt) || isChartDRP(mt) || isChartSell(mt) || isChartCashDiv(mt)
    })
    .sort((a, b) => toSortableDate(a.date) - toSortableDate(b.date))

  // Normalise dot sizes across all trades (min 4, max 20)
  const settlements = relevant.map(t => Math.abs(Number(t.settlementAmount) || 0))
  const minS   = settlements.length ? Math.min(...settlements) : 0
  const maxS   = settlements.length ? Math.max(...settlements) : 1
  const sRange = maxS - minS || 1
  const normSize = s => 4 + ((s - minS) / sRange) * 16

  let cumulative = 0
  const lineData     = []
  const buyPoints    = []
  const drpPoints    = []
  const cashDivPoints = []
  const sellPoints   = []

  for (const t of relevant) {
    const mt         = t.movementType
    const settlement = Math.abs(Number(t.settlementAmount) || 0)

    if      (isChartBuyOnly(mt)) cumulative += settlement
    else if (isChartDRP(mt))     cumulative += settlement
    else if (isChartSell(mt))    cumulative -= settlement
    // cash dividends: no change to cumulative

    const x    = toSortableDate(t.date)
    const base = {
      x,
      y:               cumulative,
      dateStr:         fmtDate(t.date),
      code:            t.code,
      movementType:    mt,
      quantity:        Math.abs(Number(t.quantity) || 0),
      transactionPrice: t.transactionPrice,
      settlementAmount: t.settlementAmount,
      size:            normSize(settlement),
    }

    // Only capital-moving events update the step line
    if (!isChartCashDiv(mt)) lineData.push({ x, y: cumulative })

    if      (isChartBuyOnly(mt)) buyPoints.push(base)
    else if (isChartDRP(mt))     drpPoints.push(base)
    else if (isChartCashDiv(mt)) cashDivPoints.push(base)
    else                         sellPoints.push(base)
  }

  return { lineData, buyPoints, drpPoints, cashDivPoints, sellPoints }
}

// Diamond marker — symmetric rhombus, sized from payload.size
function DiamondShape({ cx, cy, payload, fill, stroke = 'none' }) {
  if (cx == null || cy == null) return null
  const s  = payload?.size ?? 8
  const hw = s * 0.55
  const hh = s * 0.75
  const pts = `${cx},${cy - hh} ${cx + hw},${cy} ${cx},${cy + hh} ${cx - hw},${cy}`
  return <polygon points={pts} fill={fill} stroke={stroke} strokeWidth={1} opacity={0.9} />
}

function BuyDiamond(props)     { return <DiamondShape {...props} fill="#4ade80" /> }
function SellDiamond(props)    { return <DiamondShape {...props} fill="#f87171" /> }
function DRPDiamond(props)     { return <DiamondShape {...props} fill="#60a5fa" /> }
function CashDivDiamond(props) { return <DiamondShape {...props} fill="#111827" stroke="#d1d5db" /> }

function ChartTooltip({ active, payload }) {
  if (!active || !payload?.length) return null
  const scatter = payload.find(p => p.payload?.code)
  if (!scatter) return null
  const d = scatter.payload
  return (
    <div style={{
      background: '#1f2937', border: '1px solid #374151',
      borderRadius: 8, padding: '10px 14px',
      fontSize: 12, color: '#f3f4f6', lineHeight: 1.8,
    }}>
      <p style={{ fontWeight: 700, marginBottom: 4 }}>{d.dateStr}</p>
      <p>Code: <strong>{d.code}</strong></p>
      <p>Type: {d.movementType}</p>
      <p>Qty: {d.quantity?.toLocaleString('en-AU')}</p>
      <p>Price: {fmtAUD(d.transactionPrice)}</p>
      <p>Settlement: {fmtAUD(d.settlementAmount)}</p>
    </div>
  )
}

function PortfolioValueTooltip({ active, payload }) {
  if (!active || !payload?.length) return null

  // Scatter point hover — show transaction details
  const scatter = payload.find(p => p.payload?.code)
  if (scatter) {
    const d = scatter.payload
    return (
      <div style={{ background: '#1f2937', border: '1px solid #374151', borderRadius: 8, padding: '10px 14px', fontSize: 12, color: '#f3f4f6', lineHeight: 1.8 }}>
        <p style={{ fontWeight: 700, marginBottom: 4 }}>{d.dateStr}</p>
        <p>Code: <strong>{d.code}</strong></p>
        <p>Type: {d.movementType}</p>
        <p>Qty: {d.quantity?.toLocaleString('en-AU')}</p>
        <p>Price: {fmtAUD(d.transactionPrice)}</p>
        <p>Settlement: {fmtAUD(d.settlementAmount)}</p>
      </div>
    )
  }

  // Line hover — show portfolio value and benchmark
  const pt = payload.find(p => p.dataKey === 'y' || p.dataKey === 'bench')
  if (!pt) return null
  const d = pt.payload
  return (
    <div style={{ background: '#1f2937', border: '1px solid #374151', borderRadius: 8, padding: '10px 14px', fontSize: 12, color: '#f3f4f6', lineHeight: 1.9 }}>
      <p style={{ fontWeight: 700, marginBottom: 4 }}>{fmtDate(new Date(d.x))}</p>
      <p style={{ color: '#a78bfa' }}>Portfolio Value: <strong>{fmtAUD(d.y)}</strong></p>
      {d.bench != null && (
        <p style={{ color: '#fb923c' }}>VGS Benchmark: <strong>{fmtAUD(d.bench)}</strong></p>
      )}
    </div>
  )
}

function CapitalDeployedChart({ transactions }) {
  const { lineData, buyPoints, drpPoints, cashDivPoints, sellPoints } =
    buildCapitalDeployedSeries(transactions)

  if (!lineData.length) return null

  const yTickFmt = v => '$' + (v >= 1_000_000
    ? (v / 1_000_000).toFixed(1) + 'm'
    : v >= 1000 ? (v / 1000).toFixed(0) + 'k' : v.toFixed(0))

  return (
    <div className="rounded-2xl border border-gray-800 bg-gray-900 p-6">
      <h3 className="mb-6 text-sm font-semibold uppercase tracking-widest text-gray-400">
        Cumulative Capital Deployed Over Time
      </h3>
      <ResponsiveContainer width="100%" height={420}>
        <ComposedChart margin={{ top: 10, right: 20, left: 10, bottom: 10 }}>
          <CartesianGrid strokeDasharray="3 3" stroke="#1f2937" />
          <XAxis
            dataKey="x"
            type="number"
            domain={['dataMin', 'dataMax']}
            scale="time"
            tickFormatter={v => fmtDate(new Date(v))}
            tick={{ fill: '#6b7280', fontSize: 11 }}
            axisLine={{ stroke: '#374151' }}
            tickLine={false}
            tickCount={8}
          />
          <YAxis
            domain={[0, 'auto']}
            tickFormatter={yTickFmt}
            tick={{ fill: '#6b7280', fontSize: 11 }}
            axisLine={{ stroke: '#374151' }}
            tickLine={false}
            width={60}
          />
          <Tooltip content={<ChartTooltip />} />
          <Legend
            formatter={v => <span style={{ color: '#9ca3af', fontSize: 12 }}>{v}</span>}
          />
          <Line
            data={lineData}
            dataKey="y"
            type="stepAfter"
            stroke="#3b82f6"
            strokeWidth={2}
            dot={false}
            activeDot={false}
            name="Cumulative Capital"
            legendType="line"
          />
          <Scatter data={buyPoints}     dataKey="y" shape={<BuyDiamond />}     name="BUY"          legendType="diamond" />
          <Scatter data={sellPoints}    dataKey="y" shape={<SellDiamond />}    name="SELL"         legendType="diamond" />
          <Scatter data={cashDivPoints} dataKey="y" shape={<CashDivDiamond />} name="Cash Dividend" legendType="diamond" />
          <Scatter data={drpPoints}     dataKey="y" shape={<DRPDiamond />}     name="DRP"          legendType="diamond" />
          <Brush
            dataKey="x"
            data={lineData}
            tickFormatter={v => fmtDate(new Date(v))}
            height={28}
            stroke="#374151"
            fill="#111827"
            travellerWidth={6}
          />
        </ComposedChart>
      </ResponsiveContainer>
    </div>
  )
}

// ---------------------------------------------------------------------------
// Stage 8: Returns Breakdown series (data only)
// ---------------------------------------------------------------------------

/**
 * buildReturnComponentSeries(transactions, dividends, priceData)
 *
 * At each month-end timestamp present in priceData returns:
 *   x          – timestamp ms (for Recharts)
 *   unrealised – Σ (qty × price − costBasis) for all held tickers with price data
 *   realised   – cumulative realised P&L from all SELLs up to this date
 *   dividends  – cumulative cash dividends (Domestic Dividends sheet) up to this date
 *   drp        – cumulative DRP settlement amounts up to this date
 *
 * Uses the same running-pointer technique as buildPortfolioValueSeries.
 * Trims leading all-zero rows. Logs first and last 3 rows to console.
 */
function buildReturnComponentSeries(transactions, dividends, priceData) {
  if (!priceData || !Object.keys(priceData).length) return []

  // ── 1. Union of all month timestamps ───────────────────────────────────────
  const allMs = new Set()
  for (const priceMap of Object.values(priceData)) {
    for (const ms of Object.keys(priceMap)) allMs.add(Number(ms))
  }
  const sortedMonths = [...allMs].sort((a, b) => a - b)
  if (!sortedMonths.length) return []

  // ── 2. Sort inputs chronologically ─────────────────────────────────────────
  const sortedTx = [...transactions].sort(
    (a, b) => toSortableDate(a.date) - toSortableDate(b.date),
  )
  const sortedDiv = [...dividends].sort(
    (a, b) => toSortableDate(a.date) - toSortableDate(b.date),
  )

  // ── 3. Walk months with running state ──────────────────────────────────────
  const holdings   = {} // yahooTicker → { qty, totalCost }
  let cumRealised  = 0
  let cumDrp       = 0
  let cumDividends = 0
  let txIdx  = 0
  let divIdx = 0
  const result = []

  for (const ms of sortedMonths) {
    // Advance transaction pointer: apply all tx with date ≤ ms
    while (txIdx < sortedTx.length) {
      const t = sortedTx[txIdx]
      if (toSortableDate(t.date) > ms) break

      const ticker     = nabtradeTicker(t.code)
      const tradeQty   = Math.abs(Number(t.quantity)         || 0)
      const settlement = Math.abs(Number(t.settlementAmount) || 0)

      if (ticker) {
        if (!holdings[ticker]) holdings[ticker] = { qty: 0, totalCost: 0 }
        const pos = holdings[ticker]

        if (isChartBuyOnly(t.movementType)) {
          pos.qty       += tradeQty
          pos.totalCost += settlement

        } else if (isChartDRP(t.movementType)) {
          pos.qty       += tradeQty
          pos.totalCost += settlement
          cumDrp        += settlement

        } else if (isChartSell(t.movementType)) {
          if (pos.qty > 0) {
            const avgCost      = pos.totalCost / pos.qty
            cumRealised       += settlement - avgCost * tradeQty
            const remainingQty = Math.max(0, pos.qty - tradeQty)
            pos.qty            = remainingQty
            pos.totalCost      = avgCost * remainingQty
          }
        }
      }
      txIdx++
    }

    // Advance dividend pointer: accumulate all dividends with date ≤ ms
    while (divIdx < sortedDiv.length) {
      const d = sortedDiv[divIdx]
      if (toSortableDate(d.date) > ms) break
      cumDividends += Number(d.value) || 0
      divIdx++
    }

    // Unrealised: Σ (qty × price − costBasis) for tickers with price data
    let unrealised = 0
    for (const [ticker, { qty, totalCost }] of Object.entries(holdings)) {
      if (qty <= 0) continue
      const price = priceData[ticker]?.[ms]
      if (price != null) unrealised += qty * price - totalCost
    }

    result.push({ x: ms, unrealised, realised: cumRealised, dividends: cumDividends, drp: cumDrp })
  }

  const isNearZero = r =>
    Math.abs(r.unrealised) + Math.abs(r.realised) + Math.abs(r.dividends) + Math.abs(r.drp) < 1

  // Trim leading all-zero rows
  const firstActive = result.findIndex(r => !isNearZero(r))
  const afterLeadTrim = firstActive >= 0 ? result.slice(firstActive) : result

  // Trim trailing near-zero rows (missing prices at tail)
  let lastActive = afterLeadTrim.length - 1
  while (lastActive > 0 && isNearZero(afterLeadTrim[lastActive])) lastActive--
  const trimmed = afterLeadTrim.slice(0, lastActive + 1)

  // Console diagnostics
  const fmt = r => `  ${new Date(r.x).toLocaleDateString('en-AU')}` +
    `  unreal=${r.unrealised.toFixed(0)}` +
    `  real=${r.realised.toFixed(0)}` +
    `  div=${r.dividends.toFixed(0)}` +
    `  drp=${r.drp.toFixed(0)}`
  console.group('[Stage 8] buildReturnComponentSeries')
  console.log(`Total rows: ${trimmed.length}`)
  trimmed.slice(0, 3).forEach(r => console.log('First:', fmt(r)))
  trimmed.slice(-3).forEach(r => console.log('Last: ', fmt(r)))
  console.groupEnd()

  return trimmed
}

// ---------------------------------------------------------------------------
// Stage 7: Portfolio Market Value chart
// ---------------------------------------------------------------------------

/**
 * For each month-end timestamp present in priceData, replay all transactions
 * up to that point to get current holdings, then multiply by closing price.
 * Tickers with no price data are silently skipped.
 *
 * Returns { lineData, buyPoints, drpPoints, cashDivPoints, sellPoints }
 * lineData   – { x (ms), y (AUD portfolio value) } trimmed to first non-zero
 * *Points    – transaction markers placed at nearest prior monthly y value
 */
function buildPortfolioValueSeries(transactions, priceData) {
  const empty = { lineData: [], buyPoints: [], drpPoints: [], cashDivPoints: [], sellPoints: [] }

  if (!priceData || !Object.keys(priceData).length) return empty

  // ── 1. Union of all month timestamps across loaded tickers ─────────────────
  const allMs = new Set()
  for (const priceMap of Object.values(priceData)) {
    for (const ms of Object.keys(priceMap)) allMs.add(Number(ms))
  }
  const sortedMonths = [...allMs].sort((a, b) => a - b)
  if (!sortedMonths.length) return empty

  // ── 2. Walk transactions (chronological) with a running holdings map ───────
  const sorted = [...transactions].sort(
    (a, b) => toSortableDate(a.date) - toSortableDate(b.date),
  )

  const holdings = {} // yahooTicker → qty
  let txIdx = 0
  const rawLine = []

  for (const ms of sortedMonths) {
    // Advance pointer: apply every transaction with date ≤ ms
    while (txIdx < sorted.length) {
      const t = sorted[txIdx]
      if (toSortableDate(t.date) > ms) break
      const ticker = nabtradeTicker(t.code)
      if (ticker) {
        if (!holdings[ticker]) holdings[ticker] = 0
        const qty = Math.abs(Number(t.quantity) || 0)
        if (isBuyType(t.movementType))        holdings[ticker] += qty
        else if (isSellType(t.movementType))  holdings[ticker] = Math.max(0, holdings[ticker] - qty)
      }
      txIdx++
    }

    // Sum market value for this month
    let value = 0
    for (const [ticker, qty] of Object.entries(holdings)) {
      if (qty <= 0) continue
      const price = priceData[ticker]?.[ms]
      if (price != null) value += qty * price
    }
    rawLine.push({ x: ms, y: value })
  }

  // Trim leading zeros (before first buy) and trailing zeros / missing-price rows
  const firstNonZero = rawLine.findIndex(p => p.y > 0)
  if (firstNonZero < 0) return empty
  let lastNonZero = rawLine.length - 1
  while (lastNonZero > firstNonZero && rawLine[lastNonZero].y < 1) lastNonZero--
  const lineData = rawLine.slice(firstNonZero, lastNonZero + 1)

  if (!lineData.length) return empty

  // ── 3a. VGS benchmark — replay BUY transactions as hypothetical VGS purchases
  const vgsPrices = priceData['VGS.AX']
  if (vgsPrices && Object.keys(vgsPrices).length) {
    const vgsMsSorted = Object.keys(vgsPrices).map(Number).sort((a, b) => a - b)

    const getNearestVgsPrice = (ms) => {
      let closest = vgsMsSorted[0], minDiff = Math.abs(ms - closest)
      for (const v of vgsMsSorted) {
        const diff = Math.abs(ms - v)
        if (diff < minDiff) { minDiff = diff; closest = v }
      }
      return vgsPrices[closest] ?? null
    }

    let vgsUnits = 0
    let bTxIdx = 0
    const buyTx = [...transactions]
      .filter(t => isBuyType(t.movementType))
      .sort((a, b) => toSortableDate(a.date) - toSortableDate(b.date))

    // ── DEBUG: log VGS price range to detect near-zero prices ──────────────
    const vgsPriceValues = Object.values(vgsPrices)
    const vgsPriceMin = Math.min(...vgsPriceValues)
    const vgsPriceMax = Math.max(...vgsPriceValues)
    console.group('[Benchmark DEBUG] VGS.AX price data')
    console.log(`Price range: $${vgsPriceMin.toFixed(2)} – $${vgsPriceMax.toFixed(2)} over ${vgsPriceValues.length} months`)
    const nearZero = vgsPriceValues.filter(p => p < 1)
    if (nearZero.length) console.warn('⚠️  Near-zero prices found:', nearZero)
    else console.log('✅  No near-zero prices')
    console.groupEnd()

    // ── DEBUG: per-BUY transaction log ─────────────────────────────────────
    console.group('[Benchmark DEBUG] BUY transactions → hypothetical VGS units')
    console.log('Note: SELLs are NOT currently reducing VGS units (pure buy-and-hold benchmark)')
    let runningUnits = 0
    for (const t of buyTx) {
      const settlement = Math.abs(Number(t.settlementAmount) || 0)
      const txMs = toSortableDate(t.date)
      const price = getNearestVgsPrice(txMs)
      // Find which month key was actually used
      let closestMs = vgsMsSorted[0], minDiff = Math.abs(txMs - closestMs)
      for (const v of vgsMsSorted) {
        const diff = Math.abs(txMs - v)
        if (diff < minDiff) { minDiff = diff; closestMs = v }
      }
      const daysDiff = Math.round(minDiff / 86400000)
      const units = (price && price > 0 && settlement > 0) ? settlement / price : 0
      runningUnits += units
      const flag = units > 10000 ? '🚨 UNREALISTIC' : units > 1000 ? '⚠️ HIGH' : ''
      console.log(
        `${fmtDate(t.date)}  ${t.code.padEnd(8)}  settlement=$${settlement.toFixed(2).padStart(10)}` +
        `  VGS price=$${price?.toFixed(2).padStart(7) ?? 'NULL'}` +
        `  (nearest month ${daysDiff}d away: ${new Date(closestMs).toLocaleDateString('en-AU')})` +
        `  units=${units.toFixed(4).padStart(12)}  cumUnits=${runningUnits.toFixed(4).padStart(14)}  ${flag}`
      )
    }
    console.groupEnd()

    for (const pt of lineData) {
      // Apply all BUY transactions up to this month-end
      while (bTxIdx < buyTx.length) {
        const t = buyTx[bTxIdx]
        if (toSortableDate(t.date) > pt.x) break
        const settlement = Math.abs(Number(t.settlementAmount) || 0)
        if (settlement > 0) {
          const price = getNearestVgsPrice(toSortableDate(t.date))
          if (price && price > 0) vgsUnits += settlement / price
        }
        bTxIdx++
      }
      // Benchmark value = accumulated VGS units × this month's VGS price
      const monthPrice = vgsPrices[pt.x] ?? getNearestVgsPrice(pt.x)
      pt.bench = monthPrice && monthPrice > 0 ? vgsUnits * monthPrice : null
    }

    // ── DEBUG: benchmark series first/last 10 rows ─────────────────────────
    console.group('[Benchmark DEBUG] Benchmark series (first & last 10 rows)')
    console.log(`Total rows: ${lineData.length}  |  Final VGS units held: ${vgsUnits.toFixed(4)}`)
    const fmtRow = pt =>
      `${fmtDate(new Date(pt.x))}  portfolio=$${(pt.y ?? 0).toFixed(0).padStart(10)}` +
      `  bench=${pt.bench != null ? '$' + pt.bench.toFixed(0).padStart(10) : '        null'}`
    console.log('--- First 10 ---')
    lineData.slice(0, 10).forEach(pt => console.log(fmtRow(pt)))
    if (lineData.length > 10) {
      console.log('--- Last 10 ---')
      lineData.slice(-10).forEach(pt => console.log(fmtRow(pt)))
    }
    console.groupEnd()
  }

  // ── 3b. Scatter markers — placed at nearest prior monthly portfolio value ────
  const relevant = [...transactions]
    .filter(t => {
      const mt = t.movementType
      return isChartBuyOnly(mt) || isChartDRP(mt) || isChartSell(mt) || isChartCashDiv(mt)
    })
    .sort((a, b) => toSortableDate(a.date) - toSortableDate(b.date))

  // For a given tx timestamp, return the most recent lineData y at or before it
  const getValueAt = (txMs) => {
    let val = 0
    for (const pt of lineData) {
      if (pt.x <= txMs) val = pt.y
      else break
    }
    return val
  }

  const settlements = relevant.map(t => Math.abs(Number(t.settlementAmount) || 0))
  const minS   = settlements.length ? Math.min(...settlements) : 0
  const maxS   = settlements.length ? Math.max(...settlements) : 1
  const sRange = maxS - minS || 1
  const normSize = s => 4 + ((s - minS) / sRange) * 16

  const buyPoints = [], drpPoints = [], cashDivPoints = [], sellPoints = []

  for (const t of relevant) {
    const mt         = t.movementType
    const settlement = Math.abs(Number(t.settlementAmount) || 0)
    const x          = toSortableDate(t.date)
    const base = {
      x,
      y:                getValueAt(x),
      dateStr:          fmtDate(t.date),
      code:             t.code,
      movementType:     mt,
      quantity:         Math.abs(Number(t.quantity) || 0),
      transactionPrice: t.transactionPrice,
      settlementAmount: t.settlementAmount,
      size:             normSize(settlement),
    }
    if      (isChartBuyOnly(mt)) buyPoints.push(base)
    else if (isChartDRP(mt))     drpPoints.push(base)
    else if (isChartCashDiv(mt)) cashDivPoints.push(base)
    else                         sellPoints.push(base)
  }

  return { lineData, buyPoints, drpPoints, cashDivPoints, sellPoints }
}

function PortfolioValueChart({ transactions, priceData, pricesLoading, failedTickers }) {
  const [brushDomain, setBrushDomain] = React.useState(null)

  if (pricesLoading) {
    return (
      <div
        className="rounded-2xl border border-gray-800 bg-gray-900 p-6 flex items-center justify-center"
        style={{ minHeight: 200 }}
      >
        <div className="flex flex-col items-center gap-3">
          <div className="animate-spin h-8 w-8 rounded-full border-2 border-t-transparent border-blue-500" />
          <p className="text-sm text-gray-400">Loading historical prices…</p>
        </div>
      </div>
    )
  }

  if (!priceData || !Object.keys(priceData).length) return null

  const { lineData, buyPoints, drpPoints, cashDivPoints, sellPoints } =
    buildPortfolioValueSeries(transactions, priceData)

  if (!lineData.length) return null

  const hasBenchmark = lineData.some(p => p.bench != null)

  const yTickFmt = v => '$' + (v >= 1_000_000
    ? (v / 1_000_000).toFixed(1) + 'm'
    : v >= 1000 ? (v / 1000).toFixed(0) + 'k' : v.toFixed(0))

  const handleBrush = ({ startIndex, endIndex }) => {
    if (startIndex == null || endIndex == null) return
    setBrushDomain([lineData[startIndex].x, lineData[endIndex].x])
  }

  // Filter non-VGS failed tickers for display
  const displayFailed = failedTickers.filter(t => t !== 'VGS.AX')

  return (
    <div className="rounded-2xl border border-gray-800 bg-gray-900 p-6 space-y-4">
      <h3 className="text-sm font-semibold uppercase tracking-widest text-gray-400">
        Portfolio Market Value Over Time
      </h3>

      {displayFailed.length > 0 && (
        <p className="text-xs italic text-gray-500">
          No price data for {displayFailed.join(', ')} — excluded from chart.
        </p>
      )}

      <ResponsiveContainer width="100%" height={420}>
        <ComposedChart margin={{ top: 10, right: 20, left: 10, bottom: 10 }}>
          <CartesianGrid strokeDasharray="3 3" stroke="#1f2937" />
          <XAxis
            dataKey="x"
            type="number"
            domain={brushDomain ?? ['dataMin', 'dataMax']}
            scale="time"
            tickFormatter={v => fmtDate(new Date(v))}
            tick={{ fill: '#6b7280', fontSize: 11 }}
            axisLine={{ stroke: '#374151' }}
            tickLine={false}
            tickCount={8}
          />
          <YAxis
            tickFormatter={yTickFmt}
            tick={{ fill: '#6b7280', fontSize: 11 }}
            axisLine={{ stroke: '#374151' }}
            tickLine={false}
            width={60}
          />
          <Tooltip content={<PortfolioValueTooltip />} />
          <Legend
            formatter={v => <span style={{ color: '#9ca3af', fontSize: 12 }}>{v}</span>}
          />
          <Line
            data={lineData}
            dataKey="y"
            type="monotone"
            stroke="#a78bfa"
            strokeWidth={2}
            dot={false}
            activeDot={{ r: 4, fill: '#a78bfa' }}
            name="Portfolio Value"
            legendType="line"
          />
          {hasBenchmark && (
            <Line
              data={lineData}
              dataKey="bench"
              type="monotone"
              stroke="#fb923c"
              strokeWidth={2}
              strokeDasharray="5 3"
              dot={false}
              activeDot={{ r: 4, fill: '#fb923c' }}
              name="VGS Benchmark"
              legendType="line"
              connectNulls
            />
          )}
          <Scatter data={buyPoints}     dataKey="y" shape={<BuyDiamond />}     name="BUY"           legendType="diamond" />
          <Scatter data={sellPoints}    dataKey="y" shape={<SellDiamond />}    name="SELL"          legendType="diamond" />
          <Scatter data={cashDivPoints} dataKey="y" shape={<CashDivDiamond />} name="Cash Dividend"  legendType="diamond" />
          <Scatter data={drpPoints}     dataKey="y" shape={<DRPDiamond />}     name="DRP"           legendType="diamond" />
          <Brush
            data={lineData}
            dataKey="y"
            startIndex={0}
            endIndex={lineData.length - 1}
            onChange={handleBrush}
            height={28}
            stroke="#374151"
            fill="#111827"
            travellerWidth={6}
          />
        </ComposedChart>
      </ResponsiveContainer>
    </div>
  )
}

// ---------------------------------------------------------------------------
// Stage 9: Returns Breakdown AreaChart
// ---------------------------------------------------------------------------

function ReturnBreakdownTooltip({ active, payload, label }) {
  if (!active || !payload?.length) return null
  const d = payload[0]?.payload
  if (!d) return null
  const total = (d.unrealised ?? 0) + (d.realised ?? 0) + (d.dividends ?? 0) + (d.drp ?? 0)
  const row = (label, value, color) => (
    <p key={label} style={{ color, margin: 0 }}>
      {label}: <strong>{fmtAUD(value)}</strong>
    </p>
  )
  return (
    <div style={{
      background: '#1f2937', border: '1px solid #374151',
      borderRadius: 8, padding: '10px 14px',
      fontSize: 12, color: '#f3f4f6', lineHeight: 1.9,
    }}>
      <p style={{ fontWeight: 700, marginBottom: 4 }}>{fmtDate(new Date(d.x))}</p>
      {row('Unrealised P&L', d.unrealised, '#a78bfa')}
      {row('Realised P&L',   d.realised,   '#4ade80')}
      {row('Dividends',      d.dividends,  '#60a5fa')}
      {row('DRP',            d.drp,        '#2dd4bf')}
      <p style={{ borderTop: '1px solid #374151', marginTop: 6, paddingTop: 6, fontWeight: 700 }}>
        Total: {fmtAUD(total)}
      </p>
    </div>
  )
}

function ReturnBreakdownChart({ returnSeries, pricesLoading }) {
  const [brushDomain, setBrushDomain] = React.useState(null)

  if (pricesLoading) {
    return (
      <div
        className="rounded-2xl border border-gray-800 bg-gray-900 p-6 flex items-center justify-center"
        style={{ minHeight: 200 }}
      >
        <div className="flex flex-col items-center gap-3">
          <div className="animate-spin h-8 w-8 rounded-full border-2 border-t-transparent border-blue-500" />
          <p className="text-sm text-gray-400">Loading historical prices…</p>
        </div>
      </div>
    )
  }

  if (!returnSeries?.length) return null

  const handleBrush = ({ startIndex, endIndex }) => {
    if (startIndex == null || endIndex == null) return
    setBrushDomain([returnSeries[startIndex].x, returnSeries[endIndex].x])
  }

  const yTickFmt = v => '$' + (v >= 1_000_000
    ? (v / 1_000_000).toFixed(1) + 'm'
    : v <= -1_000_000 ? '-' + (Math.abs(v) / 1_000_000).toFixed(1) + 'm'
    : v >= 1000 ? (v / 1000).toFixed(0) + 'k'
    : v <= -1000 ? '-' + (Math.abs(v) / 1000).toFixed(0) + 'k'
    : v.toFixed(0))

  return (
    <div className="rounded-2xl border border-gray-800 bg-gray-900 p-6 space-y-4">
      <h3 className="text-sm font-semibold uppercase tracking-widest text-gray-400">
        Returns Breakdown Over Time
      </h3>
      <ResponsiveContainer width="100%" height={420}>
        <AreaChart data={returnSeries} margin={{ top: 10, right: 20, left: 10, bottom: 10 }}
          stackOffset="sign"
        >
          <defs>
            <linearGradient id="gradUnrealised" x1="0" y1="0" x2="0" y2="1">
              <stop offset="5%"  stopColor="#a78bfa" stopOpacity={0.35} />
              <stop offset="95%" stopColor="#a78bfa" stopOpacity={0.05} />
            </linearGradient>
            <linearGradient id="gradRealised" x1="0" y1="0" x2="0" y2="1">
              <stop offset="5%"  stopColor="#4ade80" stopOpacity={0.35} />
              <stop offset="95%" stopColor="#4ade80" stopOpacity={0.05} />
            </linearGradient>
            <linearGradient id="gradDividends" x1="0" y1="0" x2="0" y2="1">
              <stop offset="5%"  stopColor="#60a5fa" stopOpacity={0.35} />
              <stop offset="95%" stopColor="#60a5fa" stopOpacity={0.05} />
            </linearGradient>
            <linearGradient id="gradDrp" x1="0" y1="0" x2="0" y2="1">
              <stop offset="5%"  stopColor="#2dd4bf" stopOpacity={0.35} />
              <stop offset="95%" stopColor="#2dd4bf" stopOpacity={0.05} />
            </linearGradient>
          </defs>
          <CartesianGrid strokeDasharray="3 3" stroke="#1f2937" />
          <XAxis
            dataKey="x"
            type="number"
            domain={brushDomain ?? ['dataMin', 'dataMax']}
            scale="time"
            tickFormatter={v => fmtDate(new Date(v))}
            tick={{ fill: '#6b7280', fontSize: 11 }}
            axisLine={{ stroke: '#374151' }}
            tickLine={false}
            tickCount={8}
          />
          <YAxis
            tickFormatter={yTickFmt}
            tick={{ fill: '#6b7280', fontSize: 11 }}
            axisLine={{ stroke: '#374151' }}
            tickLine={false}
            width={60}
          />
          <Tooltip content={<ReturnBreakdownTooltip />} />
          <Legend
            formatter={v => <span style={{ color: '#9ca3af', fontSize: 12 }}>{v}</span>}
          />
          <Area dataKey="drp"        type="monotone" stroke="#2dd4bf" fill="url(#gradDrp)"        strokeWidth={1.5} name="DRP"           stackId="stack" />
          <Area dataKey="dividends"  type="monotone" stroke="#60a5fa" fill="url(#gradDividends)"  strokeWidth={1.5} name="Dividends"     stackId="stack" />
          <Area dataKey="realised"   type="monotone" stroke="#4ade80" fill="url(#gradRealised)"   strokeWidth={1.5} name="Realised P&L"  stackId="stack" />
          <Area dataKey="unrealised" type="monotone" stroke="#a78bfa" fill="url(#gradUnrealised)" strokeWidth={1.5} name="Unrealised P&L" stackId="stack" />
          <Brush
            dataKey="x"
            startIndex={0}
            endIndex={returnSeries.length - 1}
            onChange={handleBrush}
            height={28}
            stroke="#374151"
            fill="#111827"
            travellerWidth={6}
          />
        </AreaChart>
      </ResponsiveContainer>
    </div>
  )
}

// ---------------------------------------------------------------------------
// Yahoo Finance historical price fetching
// ---------------------------------------------------------------------------

/**
 * Convert a nabtrade instrument code to a Yahoo Finance ticker.
 * e.g. "BHP" → "BHP.AX", "WBC.ASX" → "WBC.AX", "NAB.AXW" → "NAB.AX"
 */
function nabtradeTicker(code) {
  if (!code) return null
  const base = String(code)
    .replace(/\.(ASX|AXW|AX)$/i, '')
    .trim()
    .toUpperCase()
  return base ? `${base}.AX` : null
}

/**
 * Fetch monthly closing prices for every unique ticker in `transactions`.
 * Returns { [yahooTicker]: { [date_ms]: closePrice } }
 * CORS failures are caught and warned per ticker; the rest still resolve.
 */
async function fetchHistoricalPrices(transactions) {
  // Proxy only works on Netlify; skip gracefully in local dev
  if (window.location.hostname === 'localhost' || window.location.hostname === '127.0.0.1') {
    console.warn('[prices] Skipping price fetch in local dev (Netlify proxy not available).')
    return { prices: {}, failed: [] }
  }

  const tickers = [...new Set([
    ...transactions.map(t => nabtradeTicker(t.code)).filter(Boolean),
    'VGS.AX', // always fetch for benchmark
  ])]

  if (!tickers.length) {
    console.log('[prices] No tickers found.')
    return { prices: {}, failed: [] }
  }

  console.log(`[prices] Fetching ${tickers.length} tickers:`, tickers)

  const settled = await Promise.allSettled(
    tickers.map(async (ticker) => {
      const url = `/.netlify/functions/prices?ticker=${encodeURIComponent(ticker)}`

      let res
      try {
        res = await fetch(url)
      } catch (err) {
        console.warn(`[prices] Network error – ${ticker}:`, err.message)
        return [ticker, null]
      }

      if (!res.ok) {
        console.warn(`[prices] HTTP ${res.status} – ${ticker}`)
        return [ticker, null]
      }

      const json   = await res.json()
      const result = json?.chart?.result?.[0]
      if (!result) {
        console.warn(`[prices] Empty response – ${ticker}`)
        return [ticker, null]
      }

      const timestamps = result.timestamp ?? []
      const closes     = result.indicators?.quote?.[0]?.close ?? []

      const priceMap = {}
      for (let i = 0; i < timestamps.length; i++) {
        if (closes[i] != null) priceMap[timestamps[i] * 1000] = closes[i]
      }

      return [ticker, priceMap]
    }),
  )

  const prices = {}
  const failed = []
  for (const r of settled) {
    if (r.status !== 'fulfilled') {
      // Promise itself rejected — shouldn't normally happen given our try/catch above
      continue
    }
    const [ticker, priceMap] = r.value
    if (!priceMap) {
      failed.push(ticker)
      continue
    }
    prices[ticker] = priceMap
    const ms   = Object.keys(priceMap).map(Number)
    const from = new Date(Math.min(...ms)).toLocaleDateString('en-AU')
    const to   = new Date(Math.max(...ms)).toLocaleDateString('en-AU')
    console.log(`[prices] ${ticker}: ${ms.length} months  ${from} → ${to}`)
  }

  console.log('[prices] Done. Loaded:', Object.keys(prices).join(', ') || '(none)')
  if (failed.length) console.warn('[prices] Failed:', failed.join(', '))
  return { prices, failed }
}

// ---------------------------------------------------------------------------
// App
// ---------------------------------------------------------------------------

export default function App() {
  const [fileName, setFileName]         = useState(null)
  const [isDragging, setIsDragging]     = useState(false)
  const [parsed, setParsed]             = useState(null)
  const [parseError, setParseError]     = useState(null)
  const [pricesLoading, setPricesLoading] = useState(false)
  const [priceData, setPriceData]         = useState(null)
  const [failedTickers, setFailedTickers] = useState([])
  const [returnSeries, setReturnSeries]   = useState([])

  const handleFile = (file) => {
    if (!file || !file.name.endsWith('.xlsx')) return
    setFileName(file.name)
    setParseError(null)
    setPriceData(null)
    setFailedTickers([])
    setReturnSeries([])

    const reader = new FileReader()
    reader.onload = async (e) => {
      try {
        const data = parseWorkbook(e.target.result)
        const metrics = calculateMetrics(data)
        logMetrics(metrics)
        setParsed({ ...data, metrics })
        setPricesLoading(true)
        const { prices, failed } = await fetchHistoricalPrices(data.transactions)
        setPriceData(prices)
        setFailedTickers(failed)
        setReturnSeries(buildReturnComponentSeries(data.transactions, data.dividends, prices))
      } catch (err) {
        console.error('Parse error:', err)
        setParseError(err.message)
      } finally {
        setPricesLoading(false)
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
          <div className="grid grid-cols-2 gap-6">
            <ClosedPositionsTable metrics={parsed.metrics} />
            <OpenPositionsTable   holdings={parsed.holdings} />
          </div>
          <HoldingsTable
            holdings={parsed.holdings}
            metrics={parsed.metrics}
            transactions={parsed.transactions}
          />
          <CapitalDeployedChart transactions={parsed.transactions} />
          <PortfolioValueChart
            transactions={parsed.transactions}
            priceData={priceData}
            pricesLoading={pricesLoading}
            failedTickers={failedTickers}
          />
          <ReturnBreakdownChart
            returnSeries={returnSeries}
            pricesLoading={pricesLoading}
          />
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
