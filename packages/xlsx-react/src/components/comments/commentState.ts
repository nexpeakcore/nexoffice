/**
 * Pure cell-comment logic, kept DOM-free so it is unit-testable. The editor
 * reads a sheet's comments through the wasm boundary and keeps them in state;
 * these helpers answer per-cell lookups and resolve the author a new or edited
 * comment should carry.
 */

import type { SheetComment } from '@betteroffice/xlsx';

export const DEFAULT_COMMENT_AUTHOR = 'User';

export function commentAt(
  comments: readonly SheetComment[],
  row: number,
  col: number
): SheetComment | null {
  return comments.find((comment) => comment.row === row && comment.col === col) ?? null;
}

/**
 * The author a saved comment carries: an existing comment keeps its author,
 * a new one takes the host-provided user name, and both fall back to
 * {@link DEFAULT_COMMENT_AUTHOR} when blank.
 */
export function resolveCommentAuthor(
  existing: Pick<SheetComment, 'author'> | null,
  userName: string | undefined
): string {
  const existingAuthor = existing?.author.trim() ?? '';
  if (existingAuthor !== '') return existingAuthor;
  const name = userName?.trim() ?? '';
  return name === '' ? DEFAULT_COMMENT_AUTHOR : name;
}

/** Whether a draft is savable: comments never store blank text. */
export function isSavableCommentText(text: string): boolean {
  return text.trim() !== '';
}
