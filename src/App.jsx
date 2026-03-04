import React, { useState, useCallback } from 'react'

export default function App() {
  const [fileName, setFileName] = useState(null)
  const [isDragging, setIsDragging] = useState(false)

  const handleFile = (file) => {
    if (file && file.name.endsWith('.xlsx')) {
      setFileName(file.name)
    }
  }

  const handleDrop = useCallback((e) => {
    e.preventDefault()
    setIsDragging(false)
    const file = e.dataTransfer.files[0]
    handleFile(file)
  }, [])

  const handleDragOver = useCallback((e) => {
    e.preventDefault()
    setIsDragging(true)
  }, [])

  const handleDragLeave = useCallback(() => {
    setIsDragging(false)
  }, [])

  const handleFileInput = (e) => {
    const file = e.target.files[0]
    handleFile(file)
  }

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

          {/* Filename display */}
          {fileName && (
            <div className="flex items-center gap-3 rounded-xl border border-green-900 bg-green-950 bg-opacity-40 px-4 py-3 text-left"
              style={{ backgroundColor: 'rgba(5,46,22,0.4)', borderColor: '#14532d' }}>
              <div className="flex h-6 w-6 flex-shrink-0 items-center justify-center rounded-full bg-green-600 text-xs text-white">✓</div>
              <div>
                <p className="text-sm font-medium text-green-300">{fileName}</p>
                <p className="text-xs text-gray-500">Ready to parse</p>
              </div>
            </div>
          )}

          <p className="text-xs text-gray-600">All processing happens locally in your browser. No data is uploaded.</p>
        </div>
      </main>
    </div>
  )
}
