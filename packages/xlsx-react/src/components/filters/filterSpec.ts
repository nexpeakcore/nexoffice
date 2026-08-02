/**
 * Pure auto-filter spec logic, kept DOM-free so it is unit-testable. The engine
 * matches criteria against the raw stored cell text (xlsx-ops `filter_text`):
 * numbers as their canonical string, booleans as TRUE/FALSE, text verbatim —
 * never the number-formatted display text. `CellEdit.input` mirrors that raw
 * text except for two cases handled here: guarded literals carry a leading
 * apostrophe (stripped), and formula cells carry the formula source instead of
 * the cached value (excluded — the raw computed value never crosses the wasm
 * boundary).
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
}

export interface FilterSpec {
  range: FilterRange;
  columns: FilterColumnCriteria[];
}

export interface FilterCellText {
  input: string;
  isFormula: boolean;
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
  if (cell.isFormula) return null;
  return cell.input.startsWith("'") ? cell.input.slice(1) : cell.input;
}

export function collectFilterValues(
  cells: readonly FilterCellText[],
  cap = MAX_FILTER_VALUES
): CollectedFilterValues {
  const seen = new Set<string>();
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
    if (seen.size >= cap) {
      truncated = true;
      continue;
    }
    seen.add(text);
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
  return criteria.values !== null;
}

export function withColumnCriteria(
  spec: FilterSpec,
  col: number,
  values: string[] | null,
  showBlanks: boolean
): FilterSpec {
  const next: FilterColumnCriteria =
    values === null
      ? { col, values: null, showBlanks: true }
      : { col, values: [...values], showBlanks };
  const columns = spec.columns.some((c) => c.col === col)
    ? spec.columns.map((c) => (c.col === col ? next : c))
    : [...spec.columns, next];
  return { range: spec.range, columns };
}

/**
 * Collapse a popover's checkbox state into wire criteria: everything checked
 * (values and blanks) means the column is unconstrained (`values: null`).
 */
export function resolveCriteria(
  allValues: readonly string[],
  checked: ReadonlySet<string>,
  showBlanks: boolean
): { values: string[] | null; showBlanks: boolean } {
  const allChecked = allValues.every((v) => checked.has(v));
  if (allChecked && showBlanks) return { values: null, showBlanks: true };
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
