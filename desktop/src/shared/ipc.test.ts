import { describe, expect, test } from 'bun:test'
import {
  ALL_EDIT_CAPABILITIES,
  kindFromPath,
  readEditCapabilities,
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
