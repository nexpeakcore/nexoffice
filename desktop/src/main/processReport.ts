import { app, BrowserWindow, webContents, type ProcessMetric, type WebContents } from 'electron'

export interface ProcessRow {
  pid: number
  label: string
  memoryMB: number
  peakMB: number
  cpu: number
}

/**
 * Windows Task Manager shows every Chromium child as the same `NexOffice.exe`,
 * so the only way to tell them apart is by PID. This builds that mapping.
 */
const explicitLabels = new WeakMap<WebContents, string>()

/** Name a renderer so the report can say which window a PID belongs to. */
export function labelWebContents(contents: WebContents, label: string): void {
  explicitLabels.set(contents, label)
}

const UTILITY_LABELS: Record<string, string> = {
  'network.mojom.NetworkService': 'Utility · Network',
  'storage.mojom.StorageService': 'Utility · Storage',
  'audio.mojom.AudioService': 'Utility · Audio',
  'video_capture.mojom.VideoCaptureService': 'Utility · Video capture',
  'tracing.mojom.TracingService': 'Utility · Tracing',
  'data_decoder.mojom.DataDecoderService': 'Utility · Data decoder',
  'proxy_resolver.mojom.ProxyResolverFactory': 'Utility · Proxy resolver',
}

function describeWebContents(contents: WebContents): string {
  const explicit = explicitLabels.get(contents)
  if (explicit) return explicit
  const type = contents.getType()
  if (type === 'webview' || type === 'browserView') return `Renderer · ${type}`
  const owner = BrowserWindow.fromWebContents(contents)
  const title = owner && !owner.isDestroyed() ? owner.getTitle() : ''
  return title ? `Renderer · ${title}` : 'Renderer'
}

function rendererLabelsByPid(): Map<number, string[]> {
  const byPid = new Map<number, string[]>()
  for (const contents of webContents.getAllWebContents()) {
    if (contents.isDestroyed()) continue
    let pid = 0
    try {
      pid = contents.getOSProcessId()
    } catch {
      continue
    }
    if (!pid) continue
    const existing = byPid.get(pid)
    if (existing) existing.push(describeWebContents(contents))
    else byPid.set(pid, [describeWebContents(contents)])
  }
  return byPid
}

function labelFor(metric: ProcessMetric, renderers: Map<number, string[]>): string {
  switch (metric.type) {
    case 'Browser':
      return 'Main · Electron browser process'
    case 'GPU':
      return 'GPU · compositing + canvas raster'
    case 'Tab': {
      const named = renderers.get(metric.pid)
      return named ? named.join(' + ') : 'Renderer'
    }
    case 'Utility':
      return metric.serviceName
        ? (UTILITY_LABELS[metric.serviceName] ?? `Utility · ${metric.serviceName}`)
        : 'Utility'
    default:
      return metric.name ? `${metric.type} · ${metric.name}` : metric.type
  }
}

/** Per-process memory, newest snapshot, largest first. `workingSetSize` is KB. */
export function collectProcessReport(): ProcessRow[] {
  const renderers = rendererLabelsByPid()
  return app
    .getAppMetrics()
    .map((metric) => ({
      pid: metric.pid,
      label: labelFor(metric, renderers),
      memoryMB: metric.memory.workingSetSize / 1024,
      peakMB: metric.memory.peakWorkingSetSize / 1024,
      cpu: metric.cpu.percentCPUUsage,
    }))
    .sort((a, b) => b.memoryMB - a.memoryMB)
}

function pad(value: string, width: number): string {
  return value.length >= width ? value : value + ' '.repeat(width - value.length)
}

function padStart(value: string, width: number): string {
  return value.length >= width ? value : ' '.repeat(width - value.length) + value
}

export function formatProcessReport(rows: readonly ProcessRow[]): string {
  const labelWidth = Math.max(5, ...rows.map((row) => row.label.length))
  const header = `${pad('PROCESS', labelWidth)}  ${padStart('PID', 7)}  ${padStart('RAM', 9)}  ${padStart('PEAK', 9)}  ${padStart('CPU', 6)}`
  const lines = rows.map(
    (row) =>
      `${pad(row.label, labelWidth)}  ${padStart(String(row.pid), 7)}  ${padStart(`${row.memoryMB.toFixed(1)} MB`, 9)}  ${padStart(`${row.peakMB.toFixed(1)} MB`, 9)}  ${padStart(`${row.cpu.toFixed(1)}%`, 6)}`
  )
  const total = rows.reduce((sum, row) => sum + row.memoryMB, 0)
  return [
    header,
    '-'.repeat(header.length),
    ...lines,
    '-'.repeat(header.length),
    `${pad('TOTAL', labelWidth)}  ${padStart('', 7)}  ${padStart(`${total.toFixed(1)} MB`, 9)}`,
  ].join('\n')
}
