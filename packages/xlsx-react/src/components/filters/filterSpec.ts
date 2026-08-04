/**
 * Pure auto-filter spec logic, kept DOM-free so it is unit-testable. The engine
 * matches criteria against the raw stored cell text (xlsx-ops `filter_text`):
 * numbers as their canonical string, booleans as TRUE/FALSE, text verbatim —
 * never the number-formatted display text. Current cores expose that exact
 * text as `CellEdit.filterText`, used verbatim when present. Against older
 * cores that omit it, `input` mirrors the raw text except for two cases
 * handled here: guarded literals carry a leading apostrophe (stripped), and
 * formula cells carry the formula source instead of the cached value
 * (excluded — nothing else derivable matches the engine).
 */

import type { CellRange } from '@betteroffice/xlsx';

export interface FilterRange {
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
}

export interface FilterColumnCriteria {
  col: number;
  values: string[] | null;
  showBlanks: boolean;
  /**
   * Source xml of criteria the engine keeps but cannot evaluate — custom
   * comparisons, top ten, colour, date groups. Such a column hides nothing and
   * must travel back untouched so editing a neighbouring column does not erase
   * it from the file; replacing this column's own criteria drops it.
   */
  unsupported?: string;
}

export interface FilterSpec {
  range: FilterRange;
  columns: FilterColumnCriteria[];
}

export interface FilterCellText {
  input: string;
  isFormula: boolean;
  filterText?: string;
}

export interface CollectedFilterValues {
  values: string[];
  hasBlanks: boolean;
  truncated: boolean;
}

export const MAX_FILTER_VALUES = 1000;
export const MAX_FILTER_COLUMNS = 256;

export function columnLabel(col: number): string {
  let out = '';
  for (let n = col; n >= 0; n = Math.floor(n / 26) - 1) {
    out = String.fromCharCode(65 + (n % 26)) + out;
  }
  return out;
}

export function rawFilterText(cell: FilterCellText): string | null {
  if (cell.filterText !== undefined) return cell.filterText;
  if (cell.isFormula) return null;
  return cell.input.startsWith("'") ? cell.input.slice(1) : cell.input;
}

/**
 * Distinct raw texts of a column's body cells, capped at `cap` so a huge column
 * cannot build an unbounded popover. `selected` (the column's applied criteria
 * values) is always unioned in, on top of and regardless of the cap: an active
 * selection that fell outside the cap would otherwise show as checked-but-absent
 * and be silently dropped on Apply, narrowing or emptying a valid filter. The
 * list therefore holds at most `cap` collected values plus every active
 * selection, and `truncated` stays true whenever any collected value was left
 * out — callers must treat a truncated list as incomplete (see
 * `resolveCriteria`).
 */
export function collectFilterValues(
  cells: readonly FilterCellText[],
  cap = MAX_FILTER_VALUES,
  selected: readonly string[] | null = null
): CollectedFilterValues {
  const seen = new Set<string>((selected ?? []).filter((value) => value !== ''));
  let collected = 0;
  let hasBlanks = false;
  let truncated = false;
  for (const cell of cells) {
    const text = rawFilterText(cell);
    if (text === null) continue;
    if (text === '') {
      hasBlanks = true;
      continue;
    }
    if (seen.has(text)) continue;
    if (collected >= cap) {
      truncated = true;
      continue;
    }
    seen.add(text);
    collected += 1;
  }
  const values = [...seen].sort((a, b) =>
    a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' })
  );
  return { values, hasBlanks, truncated };
}

export function filterColumnIndices(range: FilterRange): number[] {
  const end = Math.min(range.endCol, range.startCol + MAX_FILTER_COLUMNS - 1);
  const cols: number[] = [];
  for (let col = range.startCol; col <= end; col++) cols.push(col);
  return cols;
}

export function emptyFilterSpec(range: FilterRange): FilterSpec {
  return {
    range,
    columns: filterColumnIndices(range).map((col) => ({
      col,
      values: null,
      showBlanks: true,
    })),
  };
}

export function columnCriteria(spec: FilterSpec, col: number): FilterColumnCriteria {
  return spec.columns.find((c) => c.col === col) ?? { col, values: null, showBlanks: true };
}

export function hasCriteria(criteria: FilterColumnCriteria): boolean {
  return criteria.values !== null || criteria.unsupported !== undefined;
}

/**
 * Replace one column's criteria. Preserved xml the engine cannot evaluate
 * survives an unconstrained result, because a column that constrains nothing
 * is what such criteria already look like in the popover — only an explicit
 * `clearPreserved` (the Clear button) or a real allow-list drops it.
 */
export function withColumnCriteria(
  spec: FilterSpec,
  col: number,
  values: string[] | null,
  showBlanks: boolean,
  clearPreserved = false
): FilterSpec {
  const preserved = clearPreserved ? undefined : columnCriteria(spec, col).unsupported;
  const next: FilterColumnCriteria =
    values === null
      ? { col, values: null, showBlanks: true, ...(preserved ? { unsupported: preserved } : {}) }
      : { col, values: [...values], showBlanks };
  const columns = spec.columns.some((c) => c.col === col)
    ? spec.columns.map((c) => (c.col === col ? next : c))
    : [...spec.columns, next];
  return { range: spec.range, columns };
}

/**
 * Collapse a popover's checkbox state into wire criteria: everything checked
 * (values and blanks) means the column is unconstrained (`values: null`).
 * `complete` is false when the list was capped (`CollectedFilterValues.truncated`)
 * — an all-checked truncated list says nothing about the values that were left
 * out, so it stays an explicit allow-list instead of widening to unconstrained.
 * `wasUnconstrained` reopens that door for the column that was already
 * unconstrained: leaving every box checked changed nothing, so narrowing it to
 * the values that happened to fit the cap would hide the rest for no reason.
 */
export function resolveCriteria(
  allValues: readonly string[],
  checked: ReadonlySet<string>,
  showBlanks: boolean,
  complete = true,
  wasUnconstrained = false
): { values: string[] | null; showBlanks: boolean } {
  const allChecked = allValues.every((v) => checked.has(v));
  if (allChecked && showBlanks && (complete || wasUnconstrained)) {
    return { values: null, showBlanks: true };
  }
  return { values: allValues.filter((v) => checked.has(v)), showBlanks };
}

export interface RegionProbes {
  rowHasContent(row: number, left: number, right: number): boolean;
  colHasContent(col: number, top: number, bottom: number): boolean;
}

export interface RegionBounds {
  maxRow: number;
  maxCol: number;
}

const MAX_REGION_STEPS = 10_000;

/**
 * Grow a selection into the contiguous non-empty block around it (the data
 * region, including its header row). Each step annexes one adjacent row or
 * column that has content; growth stops at empty boundaries, `bounds`, or the
 * step guard. A selection with no non-empty neighbors comes back unchanged.
 */
export function expandDataRegion(
  selection: CellRange,
  probes: RegionProbes,
  bounds: RegionBounds
): CellRange {
  let bottom = Math.min(selection.bottom, bounds.maxRow);
  let right = Math.min(selection.right, bounds.maxCol);
  let top = Math.min(selection.top, bottom);
  let left = Math.min(selection.left, right);
  let steps = MAX_REGION_STEPS;
  while (steps-- > 0) {
    if (top > 0 && probes.rowHasContent(top - 1, left, right)) {
      top -= 1;
      continue;
    }
    if (bottom < bounds.maxRow && probes.rowHasContent(bottom + 1, left, right)) {
      bottom += 1;
      continue;
    }
    if (left > 0 && probes.colHasContent(left - 1, top, bottom)) {
      left -= 1;
      continue;
    }
    if (right < bounds.maxCol && probes.colHasContent(right + 1, top, bottom)) {
      right += 1;
      continue;
    }
    break;
  }
  return { top, left, bottom, right };
}
