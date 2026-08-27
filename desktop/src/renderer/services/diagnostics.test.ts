import { beforeEach, describe, expect, test } from 'bun:test'
import { registerMemoryReader, unregisterMemoryReader } from '@betteroffice/docx/diagnostics'
import type { RendererDiagnostics } from '../../shared/ipc.js'
import {
  clearDocumentProfile,
  noteDocumentOpened,
  publishDiagnostics,
  serveDiagnosticsSamples,
} from './diagnostics.js'

const published: RendererDiagnostics[] = []
let sampleHandler: (() => void) | null = null

globalThis.window = {
  nexoffice: {
    reportDiagnostics: (diagnostics: RendererDiagnostics) => published.push(diagnostics),
    onDiagnosticsSample: (handler: () => void) => {
      sampleHandler = handler
      return () => {
        sampleHandler = null
      }
    },
  },
} as unknown as Window & typeof globalThis

const latest = (): RendererDiagnostics => published[published.length - 1]!

const opened = {
  path: '/docs/report.docx',
  name: 'report.docx',
  kind: 'docx' as const,
  data: new Uint8Array(19_293_798),
  readMs: 82,
  sentAt: Date.now(),
}

beforeEach(() => {
  published.length = 0
  clearDocumentProfile()
})

describe('noteDocumentOpened', () => {
  test('names the document and charges the phases the main process measured', () => {
    noteDocumentOpened(opened, 'docx')
    const { document, open } = latest()
    expect(document).toEqual({ kind: 'docx', name: 'report.docx', bytes: 19_293_798 })
    expect(open?.read).toBe(82)
    expect(open?.transfer).toBeGreaterThanOrEqual(0)
    expect(open?.mount).toBeUndefined()
  })

  test('leaves phases unmeasured rather than zeroed when main sent no timings', () => {
    const { readMs, sentAt, ...bare } = opened
    noteDocumentOpened(bare, 'docx')
    expect(latest().open).toEqual({})
  })
})

describe('clearDocumentProfile', () => {
  test('drops the document and its timings', () => {
    noteDocumentOpened(opened, 'docx')
    clearDocumentProfile()
    expect(latest().document).toBeNull()
    expect(latest().open).toBeNull()
  })
})

describe('memory breakdown', () => {
  test('carries every registered reader, largest first', () => {
    registerMemoryReader('test · big', () => 900)
    registerMemoryReader('test · small', () => 100)
    publishDiagnostics()
    const labels = latest().memory.map((row) => row.label)
    expect(labels.indexOf('test · big')).toBeLessThan(labels.indexOf('test · small'))
    unregisterMemoryReader('test · big')
    unregisterMemoryReader('test · small')
  })

  test('leaves out a reader holding nothing, so empty caches do not read as rows', () => {
    registerMemoryReader('test · empty', () => 0)
    publishDiagnostics()
    expect(latest().memory.some((row) => row.label === 'test · empty')).toBe(false)
    unregisterMemoryReader('test · empty')
  })

  test('survives a reader that throws instead of dropping the whole sample', () => {
    registerMemoryReader('test · broken', () => {
      throw new Error('gone')
    })
    registerMemoryReader('test · fine', () => 42)
    publishDiagnostics()
    expect(latest().memory.some((row) => row.label === 'test · fine')).toBe(true)
    unregisterMemoryReader('test · broken')
    unregisterMemoryReader('test · fine')
  })
})

describe('serveDiagnosticsSamples', () => {
  test('answers the main process on request, and stops once unsubscribed', () => {
    const stop = serveDiagnosticsSamples()
    const before = published.length
    sampleHandler?.()
    expect(published.length).toBe(before + 1)
    stop()
    expect(sampleHandler).toBeNull()
  })
})
