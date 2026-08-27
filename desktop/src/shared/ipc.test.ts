import { describe, expect, test } from 'bun:test'
import {
  ALL_EDIT_CAPABILITIES,
  kindFromPath,
  readEditCapabilities,
  readRendererDiagnostics,
  sameEditCapabilities,
  type EditCapabilities,
} from './ipc.js'

const NONE: EditCapabilities = {
  cut: false,
  copy: false,
  paste: false,
  delete: false,
  selectAll: false,
}

describe('sameEditCapabilities', () => {
  test('matches identical capability sets', () => {
    expect(sameEditCapabilities(ALL_EDIT_CAPABILITIES, { ...ALL_EDIT_CAPABILITIES })).toBe(true)
    expect(sameEditCapabilities(NONE, { ...NONE })).toBe(true)
  })

  test('sees a single flipped verb, whichever one it is', () => {
    for (const key of ['cut', 'copy', 'paste', 'delete', 'selectAll'] as const) {
      expect(sameEditCapabilities(NONE, { ...NONE, [key]: true })).toBe(false)
    }
  })
})

describe('readEditCapabilities', () => {
  test('accepts a complete payload', () => {
    expect(readEditCapabilities({ ...ALL_EDIT_CAPABILITIES })).toEqual(ALL_EDIT_CAPABILITIES)
  })

  test('drops the extra properties a sender may add', () => {
    expect(readEditCapabilities({ ...NONE, shapes: true })).toEqual(NONE)
  })

  test('rejects payloads that are not a capability set', () => {
    expect(readEditCapabilities(null)).toBeNull()
    expect(readEditCapabilities(undefined)).toBeNull()
    expect(readEditCapabilities('cut')).toBeNull()
    expect(readEditCapabilities({})).toBeNull()
    expect(readEditCapabilities({ ...NONE, paste: 'yes' })).toBeNull()
    const { selectAll: _dropped, ...missing } = NONE
    expect(readEditCapabilities(missing)).toBeNull()
  })
})

describe('kindFromPath', () => {
  test('reads the supported extensions, case-insensitively', () => {
    expect(kindFromPath('/tmp/deck.pptx')).toBe('pptx')
    expect(kindFromPath('/tmp/Report.DOCX')).toBe('docx')
    expect(kindFromPath('C:\\books\\budget.xlsx')).toBe('xlsx')
  })

  test('refuses anything else', () => {
    expect(kindFromPath('/tmp/notes.txt')).toBeNull()
    expect(kindFromPath('/tmp/deck.pptx.zip')).toBeNull()
    expect(kindFromPath('/tmp/no-extension')).toBeNull()
  })
})

describe('readRendererDiagnostics', () => {
  const sample = {
    document: { kind: 'docx' as const, name: 'report.docx', bytes: 19_300_000 },
    open: { read: 82, transfer: 41, mount: 1840, interactive: 2210 },
    memory: [
      { label: 'wasm · resident engine (worker)', bytes: 642_000_000 },
      { label: 'JS heap', bytes: 189_000_000 },
    ],
  }

  test('keeps a well-formed payload whole', () => {
    expect(readRendererDiagnostics(sample)).toEqual(sample)
  })

  test('rejects a non-object and an unknown document kind', () => {
    expect(readRendererDiagnostics(null)).toBeNull()
    expect(readRendererDiagnostics('docx')).toBeNull()
    expect(readRendererDiagnostics({ ...sample, document: { ...sample.document, kind: 'rtf' } })).toBeNull()
  })

  test('accepts a renderer holding no document', () => {
    const empty = readRendererDiagnostics({ document: null, open: null, memory: [] })
    expect(empty).toEqual({ document: null, open: null, memory: [] })
  })

  test('drops memory rows that carry no readable byte count', () => {
    const read = readRendererDiagnostics({
      ...sample,
      memory: [
        { label: 'good', bytes: 10 },
        { label: 'zero', bytes: 0 },
        { label: 'negative', bytes: -5 },
        { label: 'nan', bytes: Number.NaN },
        { label: 'missing' },
        'not a row',
      ],
    })
    expect(read?.memory).toEqual([{ label: 'good', bytes: 10 }])
  })

  test('drops phases that were never measured rather than zeroing them', () => {
    const read = readRendererDiagnostics({ ...sample, open: { read: 12, mount: 'slow' } })
    expect(read?.open).toEqual({ read: 12 })
  })
})
