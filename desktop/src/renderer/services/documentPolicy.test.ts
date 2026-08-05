import { describe, expect, test } from 'bun:test'
import { createTranslator } from '../../i18n/index.js'
import { ALL_EDIT_CAPABILITIES } from '../../shared/ipc.js'
import {
  canSave,
  editCapabilities,
  exportSuffixes,
  exportedStatusKey,
  hasUnsavableEdits,
} from './documentPolicy.js'

describe('canSave', () => {
  test('allows the kinds whose engines serialize', () => {
    expect(canSave('docx')).toBe(true)
    expect(canSave('xlsx')).toBe(true)
  })

  test('refuses presentations', () => {
    expect(canSave('pptx')).toBe(false)
  })
})

describe('hasUnsavableEdits', () => {
  test('is false for a deck that has not been touched since it opened', () => {
    expect(hasUnsavableEdits('pptx', false)).toBe(false)
  })

  test('is true for a deck edited since it opened', () => {
    expect(hasUnsavableEdits('pptx', true)).toBe(true)
  })

  test('is false for savable kinds however much they changed', () => {
    expect(hasUnsavableEdits('docx', true)).toBe(false)
    expect(hasUnsavableEdits('xlsx', true)).toBe(false)
  })

  test('is false when nothing is open', () => {
    expect(hasUnsavableEdits(null, true)).toBe(false)
    expect(hasUnsavableEdits(null, false)).toBe(false)
  })
})

describe('editCapabilities', () => {
  const caret = { hasTextSelection: true, hasTextRange: false, canSelectAll: true }
  const range = { hasTextSelection: true, hasTextRange: true, canSelectAll: true }
  const shapeOnly = { hasTextSelection: false, hasTextRange: false, canSelectAll: true }

  test('leaves every verb enabled for kinds that always accept them', () => {
    expect(editCapabilities('docx', null)).toEqual(ALL_EDIT_CAPABILITIES)
    expect(editCapabilities('xlsx', null)).toEqual(ALL_EDIT_CAPABILITIES)
    expect(editCapabilities(null, null)).toEqual(ALL_EDIT_CAPABILITIES)
  })

  test('enables the clipboard verbs for a text range in a deck', () => {
    expect(editCapabilities('pptx', range)).toEqual({
      cut: true,
      copy: true,
      paste: true,
      delete: true,
      selectAll: true,
    })
  })

  test('lets a bare caret paste but not cut, copy or delete', () => {
    expect(editCapabilities('pptx', caret)).toEqual({
      cut: false,
      copy: false,
      paste: true,
      delete: false,
      selectAll: true,
    })
  })

  test('keeps the clipboard disabled for a shape selection with no text', () => {
    expect(editCapabilities('pptx', shapeOnly)).toEqual({
      cut: false,
      copy: false,
      paste: false,
      delete: false,
      selectAll: true,
    })
  })

  test('disables everything for a deck that has reported nothing yet', () => {
    expect(editCapabilities('pptx', null)).toEqual({
      cut: false,
      copy: false,
      paste: false,
      delete: false,
      selectAll: false,
    })
  })

  test('follows the editor when select all has nothing to select', () => {
    expect(
      editCapabilities('pptx', { ...shapeOnly, canSelectAll: false }).selectAll,
    ).toBe(false)
  })
})

describe('exportedStatusKey', () => {
  test('reports an unknown page count without a number', () => {
    expect(exportedStatusKey(null)).toBe('status.exported')
    expect(exportedStatusKey(undefined)).toBe('status.exported')
  })

  test('uses the singular key for one page', () => {
    expect(exportedStatusKey(1)).toBe('status.exportedPagesOne')
  })

  test('uses the plural key for any other count', () => {
    expect(exportedStatusKey(0)).toBe('status.exportedPagesMany')
    expect(exportedStatusKey(2)).toBe('status.exportedPagesMany')
    expect(exportedStatusKey(100)).toBe('status.exportedPagesMany')
  })
})

describe('exportSuffixes', () => {
  test('is empty for a complete, current export', () => {
    expect(exportSuffixes({ truncated: false, skipped: 0, asOpened: false })).toEqual([])
  })

  test('reports truncation', () => {
    expect(exportSuffixes({ truncated: true, skipped: 0, asOpened: false })).toEqual([
      { key: 'status.truncatedSuffix' },
    ])
  })

  test('reports that the export is the document as opened', () => {
    expect(exportSuffixes({ truncated: false, skipped: 0, asOpened: true })).toEqual([
      { key: 'status.asOpenedSuffix' },
    ])
  })

  test('reports a single skipped slide without a count', () => {
    expect(exportSuffixes({ truncated: false, skipped: 1, asOpened: false })).toEqual([
      { key: 'status.skippedSuffixOne' },
    ])
  })

  test('counts several skipped slides', () => {
    expect(exportSuffixes({ truncated: false, skipped: 3, asOpened: false })).toEqual([
      { key: 'status.skippedSuffixMany', vars: { slides: 3 } },
    ])
  })

  test('never reports skipped slides as truncation', () => {
    expect(exportSuffixes({ truncated: false, skipped: 2, asOpened: false })).not.toContainEqual({
      key: 'status.truncatedSuffix',
    })
  })

  test('keeps every notice rather than dropping one', () => {
    expect(exportSuffixes({ truncated: true, skipped: 2, asOpened: true })).toEqual([
      { key: 'status.truncatedSuffix' },
      { key: 'status.skippedSuffixMany', vars: { slides: 2 } },
      { key: 'status.asOpenedSuffix' },
    ])
  })
})

// The i18n check only follows translator calls with a literal key, so the keys
// these helpers hand to the translator are covered here instead.
describe('translated keys', () => {
  const t = createTranslator('en')

  test('every status key these helpers produce resolves to a string', () => {
    const keys = [
      exportedStatusKey(null),
      exportedStatusKey(1),
      exportedStatusKey(7),
      ...exportSuffixes({ truncated: true, skipped: 1, asOpened: true }).map(
        (suffix) => suffix.key,
      ),
      ...exportSuffixes({ truncated: false, skipped: 4, asOpened: false }).map(
        (suffix) => suffix.key,
      ),
    ]
    for (const key of keys) expect(t(key)).not.toBe(key)
  })
})
