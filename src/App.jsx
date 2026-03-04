import React, { useState, useCallback } from 'react'
import {
  PieChart, Pie, Cell, Tooltip, ResponsiveContainer, Legend,
} from 'recharts'

// ---------------------------------------------------------------------------
// Parsing helpers
// ---------------------------------------------------------------------------

function getCellValue(sheet, row, col) {
  const XLSX = window.XLSX
  const addr = XLSX.utils.encode_cell({ r: row, c: col })
  const cell = sheet[addr]
  return cell ? cell.v : null
}

function parseSummarySheet(sheet) {
  return {
    cashPosition:          getCellValue(sheet, 11, 5),
    domesticHoldingsValue: getCellValue(sheet, 12, 5),
    unsettledTrades:       getCellValue(sheet, 14, 5),
  }
}

function parseDomesticHoldingsSheet(sheet) {
  const rows = []
  let rowIndex = 2
  while (true) {
    const code        = getCellValue(sheet, rowIndex, 2)
    const description = getCellValue(sheet, rowIndex, 1)
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
  const rows = []
  let rowIndex = 2
  while (true) {
    const code = getCellValue(sheet, rowIndex, 2)
    if (code === null || code === undefined || code === '') break
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
  const rows = []
  let rowIndex = 2
  while (true) {
    const description = getCellValue(sheet, rowIndex, 2)
    const account     = getCellValue(sheet, rowIndex, 0)
    if ((account === null || account === '') &&
        (description === null || description === '')) break
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
  return { summary, holdings, transactions, dividends, interest }
}

// ---------------------------------------------------------------------------
// Formatting helpers
// ---------------------------------------------------------------------------

function fmt(value, opts = {}) {
  if (value === null || value === undefined) return '—'
  const { prefix = '$', decimals = 2 } = opts
  const n = Number(value)
  if (isNaN(n)) return String(value)
  return prefix + n.toLocaleString('en-AU', { minimumFractionDigits: decimals, maximumFractionDigits: decimals })
}

function fmtPct(value) {
  if (value === null || value === undefined) return '—'
  const n = Number(value) * 100
  if (isNaN(n)) return '—'
  return n.toFixed(2) + '%'
}

function fmtDate(value) {
  if (!value) return '—'
  if (value instanceof Date) return value.toLocaleDateString('en-AU')
  // Excel serial number
  if (typeof value === 'number') {
    const d = new Date(Math.round((value - 25569) * 86400 * 1000))
    return d.toLocaleDateString('en-AU')
  }
  return String(value)
}

function gainColor(value) {
  if (value === null || value === undefined) return '#9ca3af'
  return Number(value) >= 0 ? '#4ade80' : '#f87171'
}

// ---------------------------------------------------------------------------
// Palette for pie chart
// ---------------------------------------------------------------------------

const PIE_COLORS = [
  '#3b82f6','#8b5cf6','#06b6d4','#10b981','#f59e0b',
  '#ef4444','#ec4899','#14b8a6','#f97316','#6366f1',
  '#84cc16','#a78bfa','#22d3ee','#fb923c','#e879f9',
]

// ---------------------------------------------------------------------------
// Sub-components
// ---------------------------------------------------------------------------

function SummaryCard({ label, value, sub, valueColor }) {
  return (
    <div className="rounded-2xl border border-gray-800 bg-gray-900 p-5">
      <p className="text-xs font-medium uppercase tracking-widest text-gray-500">{label}</p>
      <p className="mt-2 text-2xl font-bold" style={{ color: valueColor || '#f3f4f6' }}>{value}</p>
      {sub && <p className="mt-1 text-xs text-gray-500">{sub}</p>}
    </div>
  )
}

function HoldingsTable({ holdings }) {
  const [sortKey, setSortKey] = useState('marketValue')
  const [sortDir, setSortDir] = useState(-1) // -1 = desc

  const toggleSort = (key) => {
    if (sortKey === key) setSortDir(d => d * -1)
    else { setSortKey(key); setSortDir(-1) }
  }

  const sorted = [...holdings].sort((a, b) => {
    const av = a[sortKey] ?? -Infinity
    const bv = b[sortKey] ?? -Infinity
    return (av < bv ? -1 : av > bv ? 1 : 0) * sortDir
  })

  const cols = [
    { key: 'code',        label: 'Code',        align: 'left'  },
    { key: 'description', label: 'Description', align: 'left'  },
    { key: 'quantity',    label: 'Qty',         align: 'right' },
    { key: 'averagePrice',label: 'Avg Price',   align: 'right' },
    { key: 'marketPrice', label: 'Mkt Price',   align: 'right' },
    { key: 'marketValue', label: 'Mkt Value',   align: 'right' },
    { key: 'gainLoss',    label: 'Gain/Loss',   align: 'right' },
    { key: 'gainLossPct', label: 'G/L %',       align: 'right' },
    { key: 'portfolioPct',label: 'Port %',      align: 'right' },
  ]

  return (
    <div className="overflow-x-auto rounded-2xl border border-gray-800">
      <table className="w-full text-sm">
        <thead>
          <tr className="border-b border-gray-800 bg-gray-900">
            {cols.map(col => (
              <th
                key={col.key}
                onClick={() => toggleSort(col.key)}
                className="cursor-pointer select-none px-4 py-3 font-medium text-gray-400 hover:text-white transition-colors"
                style={{ textAlign: col.align }}
              >
                {col.label}
                {sortKey === col.key && (sortDir === -1 ? ' ↓' : ' ↑')}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {sorted.map((h, i) => (
            <tr
              key={h.code + i}
              className="border-b border-gray-800 transition-colors hover:bg-gray-800/50"
              style={{ backgroundColor: i % 2 === 0 ? '#111827' : '#0f172a' }}
            >
              <td className="px-4 py-3 font-mono font-semibold text-blue-400">{h.code}</td>
              <td className="px-4 py-3 text-gray-300 max-w-xs truncate">{h.description || '—'}</td>
              <td className="px-4 py-3 text-right text-gray-300">{h.quantity?.toLocaleString('en-AU') ?? '—'}</td>
              <td className="px-4 py-3 text-right text-gray-300">{fmt(h.averagePrice)}</td>
              <td className="px-4 py-3 text-right text-gray-300">{fmt(h.marketPrice)}</td>
              <td className="px-4 py-3 text-right font-medium text-white">{fmt(h.marketValue)}</td>
              <td className="px-4 py-3 text-right font-medium" style={{ color: gainColor(h.gainLoss) }}>
                {h.gainLoss != null ? (Number(h.gainLoss) >= 0 ? '+' : '') + fmt(h.gainLoss) : '—'}
              </td>
              <td className="px-4 py-3 text-right" style={{ color: gainColor(h.gainLossPct) }}>
                {h.gainLossPct != null ? (Number(h.gainLossPct) >= 0 ? '+' : '') + fmtPct(h.gainLossPct) : '—'}
              </td>
              <td className="px-4 py-3 text-right text-gray-400">{fmtPct(h.portfolioPct)}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  )
}

function AllocationChart({ holdings }) {
  const data = holdings
    .filter(h => h.marketValue > 0)
    .sort((a, b) => b.marketValue - a.marketValue)
    .slice(0, 15)
    .map(h => ({ name: h.code, value: Number(h.marketValue) || 0 }))

  if (data.length === 0) return null

  return (
    <div className="rounded-2xl border border-gray-800 bg-gray-900 p-6">
      <h3 className="mb-4 text-sm font-semibold uppercase tracking-widest text-gray-400">Portfolio Allocation</h3>
      <ResponsiveContainer width="100%" height={320}>
        <PieChart>
          <Pie
            data={data}
            cx="50%"
            cy="50%"
            innerRadius={70}
            outerRadius={120}
            paddingAngle={2}
            dataKey="value"
          >
            {data.map((_, i) => (
              <Cell key={i} fill={PIE_COLORS[i % PIE_COLORS.length]} />
            ))}
          </Pie>
          <Tooltip
            formatter={(v) => [fmt(v), 'Value']}
            contentStyle={{ background: '#1f2937', border: '1px solid #374151', borderRadius: 8, color: '#f3f4f6' }}
            itemStyle={{ color: '#f3f4f6' }}
          />
          <Legend
            formatter={(value) => <span style={{ color: '#9ca3af', fontSize: 12 }}>{value}</span>}
          />
        </PieChart>
      </ResponsiveContainer>
    </div>
  )
}

function TransactionsTable({ transactions }) {
  if (transactions.length === 0) return <p className="text-sm text-gray-500">No transactions found.</p>
  return (
    <div className="overflow-x-auto rounded-2xl border border-gray-800">
      <table className="w-full text-sm">
        <thead>
          <tr className="border-b border-gray-800 bg-gray-900">
            {['Date','Code','Description','Type','Qty','Price','Value','Brokerage','Settlement'].map(h => (
              <th key={h} className="px-4 py-3 text-left font-medium text-gray-400">{h}</th>
            ))}
          </tr>
        </thead>
        <tbody>
          {transactions.map((t, i) => (
            <tr
              key={i}
              className="border-b border-gray-800 hover:bg-gray-800/50 transition-colors"
              style={{ backgroundColor: i % 2 === 0 ? '#111827' : '#0f172a' }}
            >
              <td className="px-4 py-3 text-gray-400 whitespace-nowrap">{fmtDate(t.date)}</td>
              <td className="px-4 py-3 font-mono font-semibold text-blue-400">{t.code}</td>
              <td className="px-4 py-3 text-gray-300 max-w-xs truncate">{t.description || '—'}</td>
              <td className="px-4 py-3 text-gray-300">{t.movementType || '—'}</td>
              <td className="px-4 py-3 text-right text-gray-300">{t.quantity?.toLocaleString('en-AU') ?? '—'}</td>
              <td className="px-4 py-3 text-right text-gray-300">{fmt(t.transactionPrice)}</td>
              <td className="px-4 py-3 text-right text-gray-300">{fmt(t.value)}</td>
              <td className="px-4 py-3 text-right text-gray-400">{fmt(t.brokerage)}</td>
              <td className="px-4 py-3 text-right text-white">{fmt(t.settlementAmount)}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  )
}

function IncomeTable({ rows, label }) {
  if (rows.length === 0) return <p className="text-sm text-gray-500">No {label.toLowerCase()} found.</p>
  const total = rows.reduce((s, r) => s + (Number(r.value) || 0), 0)
  return (
    <div className="space-y-3">
      <div className="overflow-x-auto rounded-2xl border border-gray-800">
        <table className="w-full text-sm">
          <thead>
            <tr className="border-b border-gray-800 bg-gray-900">
              {['Date','Description','Amount'].map(h => (
                <th key={h} className={`px-4 py-3 font-medium text-gray-400 ${h === 'Amount' ? 'text-right' : 'text-left'}`}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {rows.map((r, i) => (
              <tr
                key={i}
                className="border-b border-gray-800 hover:bg-gray-800/50 transition-colors"
                style={{ backgroundColor: i % 2 === 0 ? '#111827' : '#0f172a' }}
              >
                <td className="px-4 py-3 text-gray-400 whitespace-nowrap">{fmtDate(r.date)}</td>
                <td className="px-4 py-3 text-gray-300">{r.description || '—'}</td>
                <td className="px-4 py-3 text-right font-medium text-green-400">{fmt(r.value)}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <p className="text-right text-sm text-gray-400">
        Total: <span className="font-semibold text-green-400">{fmt(total)}</span>
      </p>
    </div>
  )
}

// ---------------------------------------------------------------------------
// Upload screen
// ---------------------------------------------------------------------------

function UploadScreen({ onFile }) {
  const [isDragging, setIsDragging] = useState(false)

  const handleDrop = useCallback((e) => {
    e.preventDefault()
    setIsDragging(false)
    onFile(e.dataTransfer.files[0])
  }, [onFile])

  const handleDragOver = useCallback((e) => { e.preventDefault(); setIsDragging(true) }, [])
  const handleDragLeave = useCallback(() => setIsDragging(false), [])

  return (
    <main className="flex min-h-[calc(100vh-57px)] items-center justify-center p-8">
      <div className="w-full max-w-lg space-y-6 text-center">
        <div>
          <h2 className="text-3xl font-bold text-white">Upload Your Portfolio</h2>
          <p className="mt-2 text-gray-400">Import your nabtrade Excel export to analyse your holdings</p>
        </div>
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
              <input type="file" accept=".xlsx" className="hidden" onChange={e => onFile(e.target.files[0])} />
            </label>
          </div>
        </div>
        <p className="text-xs text-gray-600">All processing happens locally in your browser. No data is uploaded.</p>
      </div>
    </main>
  )
}

// ---------------------------------------------------------------------------
// Dashboard screen
// ---------------------------------------------------------------------------

const TABS = ['Holdings', 'Allocation', 'Transactions', 'Dividends', 'Interest']

function Dashboard({ parsed, fileName, onReset }) {
  const [activeTab, setActiveTab] = useState('Holdings')
  const { summary, holdings, transactions, dividends, interest } = parsed

  const totalGainLoss = holdings.reduce((s, h) => s + (Number(h.gainLoss) || 0), 0)
  const totalValue = (summary?.domesticHoldingsValue ?? 0) + (summary?.cashPosition ?? 0)

  return (
    <main className="mx-auto max-w-7xl px-6 py-8 space-y-6">
      {/* Summary cards */}
      <div className="grid grid-cols-2 gap-4 sm:grid-cols-4">
        <SummaryCard
          label="Portfolio Value"
          value={fmt(totalValue)}
          sub="Holdings + cash"
        />
        <SummaryCard
          label="Holdings Value"
          value={fmt(summary?.domesticHoldingsValue)}
          sub={`${holdings.length} positions`}
        />
        <SummaryCard
          label="Cash Position"
          value={fmt(summary?.cashPosition)}
          sub={summary?.unsettledTrades != null ? `Unsettled: ${fmt(summary.unsettledTrades)}` : undefined}
        />
        <SummaryCard
          label="Total Gain / Loss"
          value={(totalGainLoss >= 0 ? '+' : '') + fmt(totalGainLoss)}
          valueColor={gainColor(totalGainLoss)}
          sub="Unrealised"
        />
      </div>

      {/* Tabs */}
      <div className="flex items-center gap-1 border-b border-gray-800">
        {TABS.map(tab => (
          <button
            key={tab}
            onClick={() => setActiveTab(tab)}
            className="px-4 py-2.5 text-sm font-medium transition-colors rounded-t-lg"
            style={{
              color: activeTab === tab ? '#fff' : '#6b7280',
              borderBottom: activeTab === tab ? '2px solid #3b82f6' : '2px solid transparent',
            }}
          >
            {tab}
            {tab === 'Transactions' && transactions.length > 0 && (
              <span className="ml-1.5 rounded-full bg-gray-700 px-1.5 py-0.5 text-xs text-gray-300">
                {transactions.length}
              </span>
            )}
            {tab === 'Dividends' && dividends.length > 0 && (
              <span className="ml-1.5 rounded-full bg-gray-700 px-1.5 py-0.5 text-xs text-gray-300">
                {dividends.length}
              </span>
            )}
            {tab === 'Interest' && interest.length > 0 && (
              <span className="ml-1.5 rounded-full bg-gray-700 px-1.5 py-0.5 text-xs text-gray-300">
                {interest.length}
              </span>
            )}
          </button>
        ))}
        <div className="ml-auto flex items-center gap-2 pb-1">
          <span className="text-xs text-gray-500">{fileName}</span>
          <button
            onClick={onReset}
            className="rounded-lg border border-gray-700 px-3 py-1 text-xs text-gray-400 hover:border-gray-500 hover:text-white transition-colors"
          >
            Change file
          </button>
        </div>
      </div>

      {/* Tab content */}
      <div>
        {activeTab === 'Holdings'      && <HoldingsTable holdings={holdings} />}
        {activeTab === 'Allocation'    && <AllocationChart holdings={holdings} />}
        {activeTab === 'Transactions'  && <TransactionsTable transactions={transactions} />}
        {activeTab === 'Dividends'     && <IncomeTable rows={dividends} label="Dividends" />}
        {activeTab === 'Interest'      && <IncomeTable rows={interest}  label="Interest" />}
      </div>
    </main>
  )
}

// ---------------------------------------------------------------------------
// App
// ---------------------------------------------------------------------------

export default function App() {
  const [fileName, setFileName]     = useState(null)
  const [parsed, setParsed]         = useState(null)
  const [parseError, setParseError] = useState(null)

  const handleFile = (file) => {
    if (!file || !file.name.endsWith('.xlsx')) return
    setFileName(file.name)
    setParseError(null)
    const reader = new FileReader()
    reader.onload = (e) => {
      try {
        setParsed(parseWorkbook(e.target.result))
      } catch (err) {
        console.error('Parse error:', err)
        setParseError(err.message)
      }
    }
    reader.readAsArrayBuffer(file)
  }

  const handleReset = () => { setFileName(null); setParsed(null); setParseError(null) }

  return (
    <div className="min-h-screen text-gray-100" style={{ backgroundColor: '#0a0a0f' }}>
      {/* Header */}
      <header className="border-b border-gray-800 bg-gray-900 px-6 py-4">
        <div className="mx-auto max-w-7xl flex items-center gap-3">
          <div className="h-8 w-8 rounded-lg bg-blue-600 flex items-center justify-center text-sm font-bold">P</div>
          <h1 className="text-lg font-semibold tracking-wide text-white">Portfolio Dashboard</h1>
        </div>
      </header>

      {parseError && (
        <div className="mx-auto max-w-7xl px-6 py-4">
          <div
            className="flex items-center gap-3 rounded-xl border px-4 py-3"
            style={{ backgroundColor: 'rgba(69,10,10,0.4)', borderColor: '#7f1d1d' }}
          >
            <div className="flex h-6 w-6 flex-shrink-0 items-center justify-center rounded-full bg-red-600 text-xs text-white">!</div>
            <div>
              <p className="text-sm font-medium text-red-300">Parse failed</p>
              <p className="text-xs text-gray-500">{parseError}</p>
            </div>
            <button onClick={handleReset} className="ml-auto text-xs text-gray-400 hover:text-white">Try again</button>
          </div>
        </div>
      )}

      {!parsed
        ? <UploadScreen onFile={handleFile} />
        : <Dashboard parsed={parsed} fileName={fileName} onReset={handleReset} />
      }
    </div>
  )
}
