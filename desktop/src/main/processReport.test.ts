import { describe, expect, mock, test } from 'bun:test'

mock.module('electron', () => ({
  app: { getAppMetrics: () => [] },
  BrowserWindow: { fromWebContents: () => null },
  webContents: { getAllWebContents: () => [] },
}))

const { formatProcessReport } = await import('./processReport.js')

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
