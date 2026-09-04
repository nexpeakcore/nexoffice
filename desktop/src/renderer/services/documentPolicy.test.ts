import { describe, expect, test } from 'bun:test'
import type { PptxSaveFaultCode } from '@betteroffice/pptx'
import { createTranslator, SUPPORTED_LOCALES } from '../../i18n/index.js'
import { ALL_EDIT_CAPABILITIES } from '../../shared/ipc.js'
import {
  editCapabilities,
  exportSuffixes,
  exportedStatusKey,
  saveOutcomeStatus,
  saveRefusal,
  seedStandsIn,
  unsavedStep,
} from './documentPolicy.js'

const REASON = 'slide 1 shape "Title 1" added or removed a paragraph'

// What the boundary actually throws: an `Error` carrying the writer's code.
// Building it here rather than importing a class keeps the tests honest about
// what the desktop reads — the properties, never the wording.
function thrown(code: PptxSaveFaultCode, undoingHelps: boolean, reason = REASON): Error {
  const error = new Error(`the writer's own sentence: ${reason}`)
  return Object.assign(error, { code, reason, undoingHelps })
}

describe('saveRefusal', () => {
  test('reads the writer’s account off a change an undo would clear', () => {
    expect(saveRefusal(thrown('unprojectable', true))).toBe(REASON)
  })

  // These all stop a save, and none of them is the user's change. Reading one
  // as a refusal offers to abandon edits that nothing the user did cost — a
  // budget wants a smaller save, a broken write wants another attempt, and a
  // replica with no source bytes has no way out at all.
  test('never reads a stopped save the user did not cause as a refusal', () => {
    expect(saveRefusal(thrown('limit', false))).toBeNull()
    expect(saveRefusal(thrown('write-failed', false))).toBeNull()
    expect(saveRefusal(thrown('verification-failed', false))).toBeNull()
    expect(saveRefusal(thrown('unsavable', false))).toBeNull()
  })

  // Every one of these throws from the same call site a refusal does, so
  // classifying by where the call sat would read them all as refusals.
  test('refuses to read a failed save as a refusal', () => {
    expect(saveRefusal(new Error('presentation handle is disposed'))).toBeNull()
    expect(saveRefusal(new WebAssembly.RuntimeError('unreachable'))).toBeNull()
    expect(saveRefusal(new TypeError('doc.saveBytes is not a function'))).toBeNull()
    expect(saveRefusal('some string the boundary threw')).toBeNull()
    expect(saveRefusal(undefined)).toBeNull()
  })

  // The wording is not the signal. An error that quotes the writer word for
  // word, or claims a code the writer does not issue, still carries nothing.
  test('does not promote an error that merely reads like a refusal', () => {
    expect(
      saveRefusal(
        new Error('this deck holds a change the PPTX writer cannot save yet: ' + REASON),
      ),
    ).toBeNull()
    expect(
      saveRefusal(Object.assign(new Error('x'), { code: 'made-up', undoingHelps: true })),
    ).toBeNull()
  })
})

describe('saveOutcomeStatus', () => {
  test('names the file a save wrote', () => {
    expect(saveOutcomeStatus({ status: 'saved', path: '/decks/Deck.pptx' })).toEqual({
      key: 'status.saved',
      vars: { path: '/decks/Deck.pptx' },
    })
  })

  test('says nothing more about a canceled dialog', () => {
    expect(saveOutcomeStatus({ status: 'canceled' })).toEqual({ key: 'status.saveCanceled' })
  })

  test('passes a refusal through whole rather than summarizing it', () => {
    expect(saveOutcomeStatus({ status: 'refused', message: REASON })).toEqual({
      key: 'status.saveRefused',
      vars: { message: REASON },
    })
  })

  test('keeps a refusal apart from a failed write', () => {
    const refused = saveOutcomeStatus({ status: 'refused', message: 'x' })
    const failed = saveOutcomeStatus({ status: 'failed', message: 'x' })
    expect(failed).toEqual({ key: 'status.saveFailed', vars: { message: 'x' } })
    expect(refused.key).not.toBe(failed.key)
  })
})

describe('unsavedStep', () => {
  test('lets the close continue once the save wrote', () => {
    expect(unsavedStep({ status: 'saved', path: '/decks/Deck.pptx' })).toEqual({ step: 'saved' })
  })

  test('stops on a canceled save dialog rather than closing the document', () => {
    expect(unsavedStep({ status: 'canceled' })).toEqual({ step: 'stop' })
  })

  test('stops on a failed write, which retrying may still fix', () => {
    expect(unsavedStep({ status: 'failed', message: 'disk full' })).toEqual({ step: 'stop' })
  })

  test('offers the escape, carrying the reason, only for a refusal', () => {
    expect(unsavedStep({ status: 'refused', message: REASON })).toEqual({
      step: 'escape',
      message: REASON,
    })
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
describe('seedStandsIn', () => {
  test('the opened bytes stand for a document nobody has edited', () => {
    expect(seedStandsIn(0, 0)).toBe(true)
    expect(seedStandsIn(9, 9)).toBe(true)
  })

  test('they never stand for one with an edit the editor did not hand back', () => {
    // Writing them here is the silent loss: the save reports success and the
    // file on disk is the document as it was before the first keystroke.
    expect(seedStandsIn(10, 9)).toBe(false)
  })
})

describe('translated keys', () => {
  const t = createTranslator('en')

  test('every status key these helpers produce resolves to a string', () => {
    const keys = [
      exportedStatusKey(null),
      exportedStatusKey(1),
      exportedStatusKey(7),
      saveOutcomeStatus({ status: 'saved', path: '/decks/Deck.pptx' }).key,
      saveOutcomeStatus({ status: 'canceled' }).key,
      saveOutcomeStatus({ status: 'refused', message: REASON }).key,
      saveOutcomeStatus({ status: 'failed', message: 'disk full' }).key,
      ...exportSuffixes({ truncated: true, skipped: 1, asOpened: true }).map(
        (suffix) => suffix.key,
      ),
      ...exportSuffixes({ truncated: false, skipped: 4, asOpened: false }).map(
        (suffix) => suffix.key,
      ),
    ]
    for (const key of keys) expect(t(key)).not.toBe(key)
  })

  test('every locale can say that a save produced nothing', () => {
    for (const locale of SUPPORTED_LOCALES) {
      const translate = createTranslator(locale)
      expect(translate('status.saveProducedNothing')).not.toBe('status.saveProducedNothing')
    }
  })
})
