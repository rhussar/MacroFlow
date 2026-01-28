import React, { useState, useEffect } from 'react'
import { Activity, Eye, Edit, Zap, Play, RefreshCw, Palette, CircleDot } from 'lucide-react'
import { Card, CardHeader, CardTitle, CardDescription, CardContent } from './ui/card'
import { Button } from './ui/button'

const HELLO_MACRO = `Sub HelloWorld()
    MsgBox "Hello from MacroFlow!", vbInformation, "MacroFlow"
End Sub`

interface Workbook {
  name: string
  path: string
}

declare global {
  interface Window {
    electron: {
      getSelection: () => Promise<{ success: boolean; address?: string; value?: any; sheetName?: string; message?: string }>
      writeCell: (value: any) => Promise<{ success: boolean; address?: string; message?: string }>
      highlight: (color: string) => Promise<{ success: boolean; color?: string; message?: string }>
    }
    excel: {
      workbook: {
        list: () => Promise<{ success: boolean; workbooks: Workbook[]; message?: string }>
      }
      vba: {
        inject: (args: { moduleName?: string; code: string }) => Promise<{ success: boolean; message: string }>
        run: (args: { macroName: string }) => Promise<{ success: boolean; message: string }>
      }
    }
  }
}

export function ConnectivityDashboard() {
  const [workbooks, setWorkbooks] = useState<Workbook[]>([])
  const [isConnected, setIsConnected] = useState(false)
  const [scannerLoading, setScannerLoading] = useState(false)

  const [selectionResult, setSelectionResult] = useState('')
  const [readerLoading, setReaderLoading] = useState(false)

  const [writeValue, setWriteValue] = useState('')
  const [writeResult, setWriteResult] = useState('')
  const [writerLoading, setWriterLoading] = useState(false)

  const [highlightResult, setHighlightResult] = useState('')

  const [injectResult, setInjectResult] = useState('')
  const [injectorLoading, setInjectorLoading] = useState(false)

  useEffect(() => {
    scanWorkbooks()
    const interval = setInterval(scanWorkbooks, 5000)
    return () => clearInterval(interval)
  }, [])

  const scanWorkbooks = async () => {
    setScannerLoading(true)
    try {
      const result = await window.excel.workbook.list()
      if (result.success) {
        setWorkbooks(result.workbooks)
        setIsConnected(true)
      } else {
        setWorkbooks([])
        setIsConnected(false)
      }
    } catch {
      setWorkbooks([])
      setIsConnected(false)
    } finally {
      setScannerLoading(false)
    }
  }

  const getSelection = async () => {
    setReaderLoading(true)
    setSelectionResult('')
    try {
      const result = await window.electron.getSelection()
      if (result.success) {
        const displayValue = result.value ?? '(empty)'
        setSelectionResult(`${result.sheetName}!${result.address}: '${displayValue}'`)
      } else {
        setSelectionResult(`Error: ${result.message}`)
      }
    } catch (err: any) {
      setSelectionResult(`Error: ${err.message}`)
    } finally {
      setReaderLoading(false)
    }
  }

  const writeCellValue = async () => {
    if (!writeValue.trim()) return
    setWriterLoading(true)
    setWriteResult('')
    try {
      const result = await window.electron.writeCell(writeValue)
      if (result.success) {
        setWriteResult(`Wrote to ${result.address}`)
        setWriteValue('')
      } else {
        setWriteResult(`Error: ${result.message}`)
      }
    } catch (err: any) {
      setWriteResult(`Error: ${err.message}`)
    } finally {
      setWriterLoading(false)
    }
  }

  const highlightCells = async (color: string) => {
    setHighlightResult('')
    try {
      const result = await window.electron.highlight(color)
      if (result.success) {
        setHighlightResult(`Applied: ${result.color}`)
      } else {
        setHighlightResult(`Error: ${result.message}`)
      }
    } catch (err: any) {
      setHighlightResult(`Error: ${err.message}`)
    }
  }

  const injectMacro = async () => {
    setInjectorLoading(true)
    setInjectResult('')
    try {
      const result = await window.excel.vba.inject({
        moduleName: 'MacroFlowDemo',
        code: HELLO_MACRO
      })
      if (result.success) {
        setInjectResult('Macro injected!')
      } else {
        setInjectResult(`Error: ${result.message}`)
      }
    } catch (err: any) {
      setInjectResult(`Error: ${err.message}`)
    } finally {
      setInjectorLoading(false)
    }
  }

  const runMacro = async () => {
    setInjectorLoading(true)
    try {
      const result = await window.excel.vba.run({ macroName: 'HelloWorld' })
      if (result.success) {
        setInjectResult('Macro executed!')
      } else {
        setInjectResult(`Error: ${result.message}`)
      }
    } catch (err: any) {
      setInjectResult(`Error: ${err.message}`)
    } finally {
      setInjectorLoading(false)
    }
  }

  return (
    <div className="min-h-screen bg-zinc-950 p-6">
      <div className="max-w-6xl mx-auto space-y-6">
        {/* Header */}
        <div className="flex items-center justify-between mb-2">
          <div className="flex items-center gap-3">
            <Activity className="w-7 h-7 text-white" />
            <h1 className="text-2xl font-bold text-white">MacroFlow</h1>
          </div>
        </div>

        {/* Connection Status Card */}
        <Card>
          <CardHeader className="pb-3">
            <div className="flex items-center justify-between">
              <div className="flex items-center gap-3">
                <CircleDot className={`w-5 h-5 ${isConnected ? 'text-emerald-500' : 'text-red-500'}`} />
                <CardTitle className="text-lg">Connection Status</CardTitle>
              </div>
              <Button
                variant="default"
                size="sm"
                onClick={scanWorkbooks}
                disabled={scannerLoading}
              >
                <RefreshCw className={`w-4 h-4 ${scannerLoading ? 'animate-spin' : ''}`} />
                Scan
              </Button>
            </div>
          </CardHeader>
          <CardContent>
            <div className="flex items-center gap-3">
              <span className={`w-3 h-3 rounded-full ${isConnected ? 'bg-emerald-500 shadow-lg shadow-emerald-500/50' : 'bg-red-500 shadow-lg shadow-red-500/50'}`} />
              <div>
                <p className={`font-medium ${isConnected ? 'text-emerald-400' : 'text-red-400'}`}>
                  {isConnected ? 'Connected to Excel' : 'Excel Not Found'}
                </p>
                <CardDescription className="mt-0.5">
                  {workbooks.length > 0
                    ? `${workbooks.length} workbook${workbooks.length > 1 ? 's' : ''} open: ${workbooks.map(w => w.name).join(', ')}`
                    : 'No workbooks open'}
                </CardDescription>
              </div>
            </div>
          </CardContent>
        </Card>

        {/* Main Action Cards Grid */}
        <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
          {/* Data Interaction Card */}
          <Card>
            <CardHeader>
              <CardTitle className="flex items-center gap-2">
                <Eye className="w-5 h-5" />
                Data Interaction
              </CardTitle>
              <CardDescription>Read from and write to Excel cells</CardDescription>
            </CardHeader>
            <CardContent className="space-y-4">
              {/* Read Section */}
              <div className="space-y-2">
                <label className="text-xs font-medium text-zinc-200 uppercase tracking-wide">Read Selection</label>
                <Button
                  variant="outline"
                  className="w-full"
                  onClick={getSelection}
                  disabled={readerLoading}
                >
                  <Eye className="w-4 h-4" />
                  {readerLoading ? 'Reading...' : 'Get Selection'}
                </Button>
                {selectionResult && (
                  <div className={`text-xs font-mono p-2.5 rounded-md ${
                    selectionResult.startsWith('Error') 
                      ? 'bg-red-900/80 text-red-200 border-2 border-red-600' 
                      : 'bg-emerald-900/80 text-emerald-200 border-2 border-emerald-600'
                  }`}>
                    {selectionResult}
                  </div>
                )}
              </div>

              <div className="border-t border-zinc-800 my-4" />

              {/* Write Section */}
              <div className="space-y-2">
                <label className="text-xs font-medium text-zinc-200 uppercase tracking-wide">Write to Cell</label>
                <div className="flex gap-2">
                  <input
                    type="text"
                    value={writeValue}
                    onChange={(e) => setWriteValue(e.target.value)}
                    onKeyDown={(e) => e.key === 'Enter' && writeCellValue()}
                    placeholder="Enter value..."
                    className="flex-1 h-10 bg-zinc-900 border-2 border-zinc-600 text-white placeholder-zinc-400 rounded-md px-3 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
                  />
                  <Button
                    variant="outline"
                    size="icon"
                    onClick={writeCellValue}
                    disabled={writerLoading || !writeValue.trim()}
                  >
                    <Edit className="w-4 h-4" />
                  </Button>
                </div>
                {writeResult && (
                  <div className={`text-xs font-mono p-2.5 rounded-md ${
                    writeResult.startsWith('Error') 
                      ? 'bg-red-950/50 text-red-400 border border-red-900/50' 
                      : 'bg-emerald-950/50 text-emerald-400 border border-emerald-900/50'
                  }`}>
                    {writeResult}
                  </div>
                )}
              </div>
            </CardContent>
          </Card>

          {/* Cell Highlighter Card */}
          <Card>
            <CardHeader>
              <CardTitle className="flex items-center gap-2">
                <Palette className="w-5 h-5" />
                Cell Highlighter
              </CardTitle>
              <CardDescription>Highlight selected cells with colors</CardDescription>
            </CardHeader>
            <CardContent>
              <div className="flex justify-center gap-3 mb-4">
                <button
                  onClick={() => highlightCells('Yellow')}
                  className="w-16 h-16 rounded-full bg-yellow-500 hover:ring-2 hover:ring-yellow-400 hover:scale-110 transition-all shadow-lg shadow-yellow-500/30"
                  title="Yellow"
                >
                  <span className="sr-only">Yellow</span>
                </button>
                <button
                  onClick={() => highlightCells('Green')}
                  className="w-16 h-16 rounded-full bg-green-500 hover:ring-2 hover:ring-green-400 hover:scale-110 transition-all shadow-lg shadow-green-500/30"
                  title="Green"
                >
                  <span className="sr-only">Green</span>
                </button>
                <button
                  onClick={() => highlightCells('Red')}
                  className="w-16 h-16 rounded-full bg-red-500 hover:ring-2 hover:ring-red-400 hover:scale-110 transition-all shadow-lg shadow-red-500/30"
                  title="Red"
                >
                  <span className="sr-only">Red</span>
                </button>
                <button
                  onClick={() => highlightCells('None')}
                  className="w-16 h-16 rounded-full bg-zinc-700 border-2 border-zinc-500 hover:ring-2 hover:ring-zinc-400 hover:scale-110 transition-all flex items-center justify-center text-xs text-zinc-200 font-medium"
                  title="Clear"
                >
                  CLR
                </button>
              </div>

              {highlightResult && (
                <div className={`text-xs font-mono p-2.5 rounded-md text-center ${
                  highlightResult.startsWith('Error') 
                    ? 'bg-red-900/80 text-red-200 border-2 border-red-600' 
                    : 'bg-emerald-900/80 text-emerald-200 border-2 border-emerald-600'
                }`}>
                  {highlightResult}
                </div>
              )}
            </CardContent>
          </Card>
        </div>

        {/* Macro Engine Card */}
        <Card>
          <CardHeader>
            <CardTitle className="flex items-center gap-2">
              <Zap className="w-5 h-5" />
              Macro Engine
            </CardTitle>
            <CardDescription>Inject and execute VBA macros in the active workbook</CardDescription>
          </CardHeader>
          <CardContent className="space-y-4">
            <div className="flex gap-3">
              <Button
                variant="default"
                className="flex-1"
                onClick={injectMacro}
                disabled={injectorLoading}
              >
                <Zap className="w-4 h-4" />
                Inject Macro
              </Button>
              <Button
                variant="outline"
                className="flex-1"
                onClick={runMacro}
                disabled={injectorLoading}
              >
                <Play className="w-4 h-4" />
                Run Macro
              </Button>
            </div>

            {injectResult && (
              <div className={`text-xs font-mono p-2.5 rounded-md ${
                injectResult.startsWith('Error') 
                  ? 'bg-red-950/50 text-red-400 border border-red-900/50' 
                  : 'bg-emerald-950/50 text-emerald-400 border border-emerald-900/50'
              }`}>
                {injectResult}
              </div>
            )}
          </CardContent>
        </Card>

      </div>
    </div>
  )
}
