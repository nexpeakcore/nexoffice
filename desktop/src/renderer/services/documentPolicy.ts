import {
  ALL_EDIT_CAPABILITIES,
  type DocumentKind,
  type EditCapabilities,
} from '../../shared/ipc.js'

// The pptx engine keeps edits in its CRDT and never projects them back into
// PresentationML, so there is no serializer to write an edited deck with.
// Saving the bytes it opened with would clear a dirty flag over a file that
// still holds the original content, so a presentation is never dirty and never
// saved.
export function canSave(kind: DocumentKind): boolean {
  return kind !== 'pptx'
}

// The missing serializer also decides what an export renders: with nothing to
// write, both save and export can only reach for the bytes the document opened
// with, so once such a document is edited the PDF no longer matches the screen.
// This is the one condition behind the footer notice and the export prompt —
// it is deliberately not the dirty flag, which would trap the window in
// `ensureSaved` for a document that can never be saved.
export function hasUnsavableEdits(
  kind: DocumentKind | null,
  changedSinceOpen: boolean,
): boolean {
  return kind !== null && !canSave(kind) && changedSinceOpen
}

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

export function exportedStatusKey(pages: number | null | undefined): string {
  if (pages == null) return 'status.exported'
  return pages === 1 ? 'status.exportedPagesOne' : 'status.exportedPagesMany'
}

export function exportSuffixKeys(result: { truncated: boolean; asOpened: boolean }): string[] {
  const keys: string[] = []
  if (result.truncated) keys.push('status.truncatedSuffix')
  if (result.asOpened) keys.push('status.asOpenedSuffix')
  return keys
}
