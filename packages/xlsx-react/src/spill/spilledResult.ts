import type { CellEdit, CellRange, EditRefusal } from '@betteroffice/xlsx';
import type { TranslationKey } from '@betteroffice/xlsx-i18n';

/**
 * The cells a focused cell's spilled result fills, in selection coordinates,
 * or `null` when the cell is not part of one.
 *
 * Excel outlines the whole result whenever any of it is selected, which is
 * the only way to tell a spilled value from one somebody typed.
 */
export function spilledRegion(cell: CellEdit | null | undefined): CellRange | null {
  const spill = cell?.spill;
  if (!spill) return null;
  return {
    top: spill.range.start.row,
    left: spill.range.start.col,
    bottom: spill.range.end.row,
    right: spill.range.end.col,
  };
}

/** What the formula bar shows, and whether it belongs to another cell. */
export interface FormulaBarView {
  value: string;
  /** The cell whose formula is on show, when it is not the focused one. */
  borrowedFrom: string | null;
}

/**
 * A cell filled by a spilled result has no formula of its own — it holds a
 * value the formula in the anchor wrote. Excel shows that formula, greyed,
 * rather than the value, so the reader can see where the number came from.
 * A draft the reader has started always wins.
 */
export function formulaBarView(
  cell: CellEdit | null | undefined,
  draft: string | null
): FormulaBarView {
  if (draft !== null) return { value: draft, borrowedFrom: null };
  const spill = cell?.spill;
  if (!spill || !cell || cell.a1 === spill.anchor) {
    return { value: cell?.input ?? '', borrowedFrom: null };
  }
  return { value: spill.input, borrowedFrom: spill.anchor };
}

/** A message key and its variables, for the host to translate. */
export interface Notice {
  key: TranslationKey;
  vars: Record<string, string>;
}

/** What to tell the reader about a refused edit. */
export function refusalNotice(refusal: EditRefusal | undefined): Notice | null {
  if (!refusal) return null;
  if (refusal.kind === 'spilledCell') {
    return {
      key: 'spill.cannotEdit',
      vars: { cell: refusal.at ?? '', anchor: refusal.anchor },
    };
  }
  return {
    key: 'spill.cannotTear',
    vars: { anchor: refusal.anchor, range: spanA1(refusal) },
  };
}

function spanA1(refusal: EditRefusal): string {
  const range = refusal.range;
  if (!range) return refusal.anchor;
  return `${columnName(range.start.col)}${range.start.row + 1}:${columnName(range.end.col)}${range.end.row + 1}`;
}

function columnName(col: number): string {
  let name = '';
  for (let rest = col; rest >= 0; rest = Math.floor(rest / 26) - 1) {
    name = String.fromCharCode(65 + (rest % 26)) + name;
  }
  return name;
}
