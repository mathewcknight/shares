import React, { useState, useCallback } from 'react'
import {
  ComposedChart, Line, Scatter, XAxis, YAxis, CartesianGrid,
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
// Stage 6: Cumulative Capital Deployed chart
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

  // Trim leading zeros (before portfolio had any value)
  const firstNonZero = rawLine.findIndex(p => p.y > 0)
  const lineData = firstNonZero >= 0 ? rawLine.slice(firstNonZero) : []

  if (!lineData.length) return empty

  // ── 3. Scatter markers — placed at nearest prior monthly portfolio value ────
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
  if (pricesLoading) {
    return (
      <div
        className="rounded-2xl border border-gray-800 bg-gray-900 p-6 flex items-center justify-center"
        style={{ minHeight: 200 }}
      >
        <div className="flex flex-col items-center gap-3">
          <div
            className="animate-spin h-8 w-8 rounded-full border-2 border-t-transparent border-blue-500"
          />
          <p className="text-sm text-gray-400">Loading historical prices…</p>
        </div>
      </div>
    )
  }

  if (!priceData || !Object.keys(priceData).length) return null

  const { lineData, buyPoints, drpPoints, cashDivPoints, sellPoints } =
    buildPortfolioValueSeries(transactions, priceData)

  if (!lineData.length) return null

  const yTickFmt = v => '$' + (v >= 1_000_000
    ? (v / 1_000_000).toFixed(1) + 'm'
    : v >= 1000 ? (v / 1000).toFixed(0) + 'k' : v.toFixed(0))

  return (
    <div className="rounded-2xl border border-gray-800 bg-gray-900 p-6 space-y-4">
      <h3 className="text-sm font-semibold uppercase tracking-widest text-gray-400">
        Portfolio Market Value Over Time
      </h3>

      {failedTickers.length > 0 && (
        <div
          className="flex items-start gap-3 rounded-xl border px-4 py-3"
          style={{ backgroundColor: 'rgba(69,40,10,0.4)', borderColor: '#92400e' }}
        >
          <div className="flex h-5 w-5 flex-shrink-0 items-center justify-center rounded-full bg-amber-600 text-xs text-white mt-0.5">
            !
          </div>
          <div>
            <p className="text-sm font-medium text-amber-300">Price data unavailable for some tickers</p>
            <p className="text-xs text-gray-400 mt-0.5">
              Excluded from chart: <strong>{failedTickers.join(', ')}</strong>
            </p>
          </div>
        </div>
      )}

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
            type="monotone"
            stroke="#a78bfa"
            strokeWidth={2}
            dot={false}
            activeDot={false}
            name="Portfolio Value"
            legendType="line"
          />
          <Scatter data={buyPoints}     dataKey="y" shape={<BuyDiamond />}     name="BUY"           legendType="diamond" />
          <Scatter data={sellPoints}    dataKey="y" shape={<SellDiamond />}    name="SELL"          legendType="diamond" />
          <Scatter data={cashDivPoints} dataKey="y" shape={<CashDivDiamond />} name="Cash Dividend"  legendType="diamond" />
          <Scatter data={drpPoints}     dataKey="y" shape={<DRPDiamond />}     name="DRP"           legendType="diamond" />
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

  const tickers = [...new Set(
    transactions.map(t => nabtradeTicker(t.code)).filter(Boolean),
  )]

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

  const handleFile = (file) => {
    if (!file || !file.name.endsWith('.xlsx')) return
    setFileName(file.name)
    setParseError(null)
    setPriceData(null)
    setFailedTickers([])

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
          <CapitalDeployedChart transactions={parsed.transactions} />
          <PortfolioValueChart
            transactions={parsed.transactions}
            priceData={priceData}
            pricesLoading={pricesLoading}
            failedTickers={failedTickers}
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
