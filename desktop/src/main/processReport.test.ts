import { describe, expect, mock, test } from 'bun:test'

mock.module('electron', () => ({
  app: { getAppMetrics: () => [] },
  BrowserWindow: { fromWebContents: () => null },
  webContents: { getAllWebContents: () => [] },
}))

const { documentRendererLabel, formatProcessReport } = await import('./processReport.js')

const rows = [
  { pid: 4812, label: 'Renderer · Document window', memoryMB: 807.5, peakMB: 902.1, cpu: 0 },
  { pid: 9120, label: 'GPU · compositing + canvas raster', memoryMB: 195.8, peakMB: 210, cpu: 0.4 },
  { pid: 1004, label: 'Main · Electron browser process', memoryMB: 36.5, peakMB: 40, cpu: 0 },
]

describe('formatProcessReport', () => {
  test('lists every pid with its memory', () => {
    const report = formatProcessReport(rows)
    for (const row of rows) {
      expect(report).toContain(String(row.pid))
      expect(report).toContain(row.label)
    }
    expect(report).toContain('807.5 MB')
  })

  test('totals the working sets', () => {
    expect(formatProcessReport(rows)).toContain('1039.8 MB')
  })

  test('aligns every row to the header width', () => {
    const lines = formatProcessReport(rows).split('\n')
    const width = lines[0]?.length ?? 0
    for (const line of lines.slice(1, -1)) expect(line.length).toBe(width)
  })

  test('renders a header and no rows for an empty snapshot', () => {
    const report = formatProcessReport([])
    expect(report).toContain('PROCESS')
    expect(report).toContain('0.0 MB')
  })
})

describe('documentRendererLabel', () => {
  test('names the open file and its size', () => {
    expect(documentRendererLabel({ kind: 'docx', name: 'report.docx', bytes: 19_293_798 })).toBe(
      'Renderer · DOCX · report.docx (18.4 MB)',
    )
  })

  test('says so when the window holds nothing', () => {
    expect(documentRendererLabel(null)).toBe('Renderer · Document window (no document)')
  })
})

describe('formatProcessReport detail', () => {
  const detail = {
    diagnostics: {
      document: { kind: 'docx' as const, name: 'report.docx', bytes: 19_293_798 },
      open: { read: 82, transfer: 41, mount: 1840, interactive: 2210 },
      memory: [
        { label: 'JS heap', bytes: 189_000_000 },
        { label: 'wasm · resident engine (worker)', bytes: 642_000_000 },
      ],
      heapStages: [
        { label: 'seed', liveBytes: 299_000_000, peakBytes: 1_493_000_000, atMs: 0 },
        { label: 'layout', liveBytes: 1_507_000_000, peakBytes: 2_628_000_000, atMs: 24_150 },
      ],
    },
    sampleAgeMs: 2_400,
  }

  test('leaves the table untouched when no detail is supplied', () => {
    expect(formatProcessReport(rows, undefined)).toBe(formatProcessReport(rows))
  })

  test('breaks the renderer down largest first and totals what it accounted for', () => {
    const report = formatProcessReport(rows, detail)
    const heap = report.indexOf('wasm · resident engine (worker)')
    const js = report.indexOf('JS heap')
    expect(heap).toBeGreaterThan(-1)
    expect(heap).toBeLessThan(js)
    expect(report).toContain('612.3 MB')
    expect(report).toContain('accounted')
    expect(report).toContain('792.5 MB')
  })

  test('dates the sample so a stale reading is not read as current', () => {
    expect(formatProcessReport(rows, detail)).toContain('sampled 2s ago')
    expect(formatProcessReport(rows, { ...detail, sampleAgeMs: null })).toContain('never sampled')
    expect(formatProcessReport(rows, { ...detail, sampleAgeMs: 120_000 })).toContain(
      'sampled 2m ago',
    )
  })

  test('reports every open phase and their total', () => {
    const report = formatProcessReport(rows, detail)
    expect(report).toContain('Document open · report.docx')
    expect(report).toContain('read 82ms · transfer 41ms · mount 1840ms · interactive 2210ms')
    expect(report).toContain('total 4173ms')
  })

  test('omits both sections when the renderer reported nothing', () => {
    const report = formatProcessReport(rows, { diagnostics: null, sampleAgeMs: null })
    expect(report).toBe(formatProcessReport(rows))
  })

  test('omits the open section for a renderer holding no document', () => {
    const report = formatProcessReport(rows, {
      diagnostics: {
        document: null,
        open: null,
        memory: [{ label: 'JS heap', bytes: 1_048_576 }],
        heapStages: [],
      },
      sampleAgeMs: 100,
    })
    expect(report).toContain('JS heap')
    expect(report).not.toContain('Document open')
  })
})
