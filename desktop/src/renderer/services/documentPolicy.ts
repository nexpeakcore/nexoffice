import { saveFault } from '@betteroffice/pptx'
import {
  ALL_EDIT_CAPABILITIES,
  type DocumentKind,
  type EditCapabilities,
} from '../../shared/ipc.js'

// The part of the pptx editor's selection state the Edit menu depends on,
// declared structurally so this stays a plain policy module.
export interface TextSelectionState {
  hasTextSelection: boolean
  hasTextRange: boolean
  canSelectAll: boolean
}

// A shape selection is not a text selection: the pptx editor has no shape
// clipboard, so cut, copy, paste and delete no-op over a selected shape and
// the menu must say so. Every other editor takes the verbs unconditionally,
// as it always has.
export function editCapabilities(
  kind: DocumentKind | null,
  selection: TextSelectionState | null,
): EditCapabilities {
  if (kind !== 'pptx') return ALL_EDIT_CAPABILITIES
  return {
    cut: selection?.hasTextRange ?? false,
    copy: selection?.hasTextRange ?? false,
    paste: selection?.hasTextSelection ?? false,
    delete: selection?.hasTextRange ?? false,
    selectAll: selection?.canSelectAll ?? false,
  }
}

export interface StatusSuffix {
  key: string
  vars?: Record<string, string | number>
}

export interface StatusMessage {
  key: string
  vars?: Record<string, string | number>
  suffixes?: StatusSuffix[]
}

// How one attempt to write the open document ended.
//
// A refusal is not a failure. The PresentationML writer expresses text edits
// inside a single run and refuses everything else by name — a paragraph break,
// a formatting patch, a moved shape, an added slide — so its message is the
// only account of what cannot be written, and it reaches the user whole. A
// failure is the write itself going wrong (a full disk, a revoked path) and
// says nothing about the edit.
export type SaveOutcome =
  | { status: 'saved'; path: string }
  | { status: 'canceled' }
  | { status: 'refused'; message: string }
  | { status: 'failed'; message: string }

// What the presentation writer said it could not express, or null when the
// throw meant something else.
//
// The writer classifies its own failures and the boundary carries the code, so
// this asks `undoingHelps` rather than reading the message: that is true for
// exactly the fault an undo clears, and it is the only one worth offering an
// escape for. A blown write budget, a broken zip, a deck that read back wrong
// and a replica that never had source bytes all say `false` — offering to
// abandon edits over any of them throws away work that the change the user
// made did not cost.
//
// The reason, not the whole message, is what comes back: the message repeats
// the writer's own preamble, which the status line already supplies.
export function saveRefusal(error: unknown): string | null {
  const fault = saveFault(error)
  return fault?.undoingHelps === true ? fault.reason : null
}

/**
 * Whether the bytes a document opened with still stand for its content.
 *
 * They do until something is edited. After that the edit lives nowhere but the
 * editor, so an editor that cannot produce bytes has failed to write — saving
 * the seed there would report success for the document as it was before the
 * first keystroke, and leave the file on disk missing every change.
 */
export function seedStandsIn(editGeneration: number, savedGeneration: number): boolean {
  return editGeneration === savedGeneration
}

export function saveOutcomeStatus(outcome: SaveOutcome): StatusMessage {
  switch (outcome.status) {
    case 'saved':
      return { key: 'status.saved', vars: { path: outcome.path } }
    case 'canceled':
      return { key: 'status.saveCanceled' }
    case 'refused':
      return { key: 'status.saveRefused', vars: { message: outcome.message } }
    case 'failed':
      return { key: 'status.saveFailed', vars: { message: outcome.message } }
  }
}

// What the unsaved-changes loop does next with the save it asked for.
//
// Only a refusal opens the escape prompt: nothing was written, the change is
// still in the deck, and asking for the same save again would refuse the same
// way — so the user is offered the exit that does not need a successful save,
// carrying the reason with it. A canceled dialog or a failed write is the
// user's own next move, and neither closes the document behind their back.
export type UnsavedStep =
  | { step: 'saved' }
  | { step: 'stop' }
  | { step: 'escape'; message: string }

export function unsavedStep(outcome: SaveOutcome): UnsavedStep {
  if (outcome.status === 'saved') return { step: 'saved' }
  if (outcome.status === 'refused') return { step: 'escape', message: outcome.message }
  return { step: 'stop' }
}

export function exportedStatusKey(pages: number | null | undefined): string {
  if (pages == null) return 'status.exported'
  return pages === 1 ? 'status.exportedPagesOne' : 'status.exportedPagesMany'
}

// Stopping at the page cap and failing to render a page are different things
// that go wrong, so an export that hit one never borrows the other's wording,
// and an export that hit both says both.
export function exportSuffixes(result: {
  truncated: boolean
  skipped: number
  asOpened: boolean
}): StatusSuffix[] {
  const suffixes: StatusSuffix[] = []
  if (result.truncated) suffixes.push({ key: 'status.truncatedSuffix' })
  if (result.skipped === 1) suffixes.push({ key: 'status.skippedSuffixOne' })
  else if (result.skipped > 1) {
    suffixes.push({ key: 'status.skippedSuffixMany', vars: { slides: result.skipped } })
  }
  if (result.asOpened) suffixes.push({ key: 'status.asOpenedSuffix' })
  return suffixes
}
