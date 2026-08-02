import { describe, expect, it } from 'bun:test';
import {
  collectFilterValues,
  columnCriteria,
  columnLabel,
  emptyFilterSpec,
  expandDataRegion,
  filterColumnIndices,
  hasCriteria,
  rawFilterText,
  resolveCriteria,
  withColumnCriteria,
} from './filterSpec';
import type { RegionProbes } from './filterSpec';

describe('columnLabel', () => {
  it('maps indices to spreadsheet letters', () => {
    expect(columnLabel(0)).toBe('A');
    expect(columnLabel(25)).toBe('Z');
    expect(columnLabel(26)).toBe('AA');
    expect(columnLabel(27)).toBe('AB');
    expect(columnLabel(701)).toBe('ZZ');
    expect(columnLabel(702)).toBe('AAA');
  });
});

describe('rawFilterText', () => {
  it('returns plain text and canonical numbers verbatim', () => {
    expect(rawFilterText({ input: 'Line item 3', isFormula: false })).toBe('Line item 3');
    expect(rawFilterText({ input: '1.5', isFormula: false })).toBe('1.5');
    expect(rawFilterText({ input: 'TRUE', isFormula: false })).toBe('TRUE');
    expect(rawFilterText({ input: '', isFormula: false })).toBe('');
  });

  it('strips the guard apostrophe from guarded literals', () => {
    expect(rawFilterText({ input: "'123", isFormula: false })).toBe('123');
    expect(rawFilterText({ input: "'=not a formula", isFormula: false })).toBe('=not a formula');
  });

  it('cannot derive text for formula cells', () => {
    expect(rawFilterText({ input: '=SUM(A1:A2)', isFormula: true })).toBeNull();
  });
});

describe('collectFilterValues', () => {
  it('dedupes, flags blanks, skips formulas, and sorts numerically', () => {
    const collected = collectFilterValues([
      { input: '10', isFormula: false },
      { input: '2', isFormula: false },
      { input: '10', isFormula: false },
      { input: '', isFormula: false },
      { input: '=A1', isFormula: true },
      { input: "'2", isFormula: false },
      { input: 'apple', isFormula: false },
    ]);
    expect(collected.values).toEqual(['2', '10', 'apple']);
    expect(collected.hasBlanks).toBe(true);
    expect(collected.truncated).toBe(false);
  });

  it('caps distinct values and reports truncation', () => {
    const cells = Array.from({ length: 5 }, (_, i) => ({
      input: `v${i}`,
      isFormula: false,
    }));
    const collected = collectFilterValues(cells, 3);
    expect(collected.values.length).toBe(3);
    expect(collected.truncated).toBe(true);
  });
});

describe('filter spec construction', () => {
  const range = { startRow: 1, startCol: 2, endRow: 9, endCol: 4 };

  it('builds an unconstrained spec covering every column', () => {
    const spec = emptyFilterSpec(range);
    expect(spec.columns).toEqual([
      { col: 2, values: null, showBlanks: true },
      { col: 3, values: null, showBlanks: true },
      { col: 4, values: null, showBlanks: true },
    ]);
    expect(spec.columns.every((c) => !hasCriteria(c))).toBe(true);
  });

  it('caps the column count', () => {
    const wide = filterColumnIndices({ startRow: 0, startCol: 0, endRow: 0, endCol: 9999 });
    expect(wide.length).toBe(256);
    expect(wide[0]).toBe(0);
    expect(wide[wide.length - 1]).toBe(255);
  });

  it('updates one column and leaves the others untouched', () => {
    const spec = emptyFilterSpec(range);
    const next = withColumnCriteria(spec, 3, ['a', 'b'], false);
    expect(columnCriteria(next, 3)).toEqual({ col: 3, values: ['a', 'b'], showBlanks: false });
    expect(columnCriteria(next, 2)).toEqual({ col: 2, values: null, showBlanks: true });
    expect(columnCriteria(spec, 3).values).toBeNull();
  });

  it('normalizes an unconstrained update to values null with blanks shown', () => {
    const spec = withColumnCriteria(emptyFilterSpec(range), 2, ['x'], false);
    const cleared = withColumnCriteria(spec, 2, null, false);
    expect(columnCriteria(cleared, 2)).toEqual({ col: 2, values: null, showBlanks: true });
  });

  it('appends criteria for a column missing from the spec', () => {
    const spec = { range, columns: [] };
    const next = withColumnCriteria(spec, 4, ['x'], true);
    expect(next.columns).toEqual([{ col: 4, values: ['x'], showBlanks: true }]);
  });
});

describe('resolveCriteria', () => {
  const values = ['a', 'b', 'c'];

  it('collapses everything-checked to unconstrained', () => {
    expect(resolveCriteria(values, new Set(values), true)).toEqual({
      values: null,
      showBlanks: true,
    });
  });

  it('keeps a checked subset in list order', () => {
    expect(resolveCriteria(values, new Set(['c', 'a']), true)).toEqual({
      values: ['a', 'c'],
      showBlanks: true,
    });
  });

  it('keeps the full value list when only blanks are excluded', () => {
    expect(resolveCriteria(values, new Set(values), false)).toEqual({
      values: ['a', 'b', 'c'],
      showBlanks: false,
    });
  });
});

describe('expandDataRegion', () => {
  // a 4x3 block at rows 2-5, cols 1-3 with a header on row 2.
  const block = { top: 2, left: 1, bottom: 5, right: 3 };
  const probes: RegionProbes = {
    rowHasContent: (row, left, right) =>
      row >= block.top && row <= block.bottom && left <= block.right && right >= block.left,
    colHasContent: (col, top, bottom) =>
      col >= block.left && col <= block.right && top <= block.bottom && bottom >= block.top,
  };
  const bounds = { maxRow: 100, maxCol: 100 };

  it('grows a single-cell selection to the whole block', () => {
    const region = expandDataRegion({ top: 3, left: 2, bottom: 3, right: 2 }, probes, bounds);
    expect(region).toEqual(block);
  });

  it('keeps a selection with no non-empty neighbors unchanged', () => {
    const region = expandDataRegion({ top: 20, left: 20, bottom: 21, right: 21 }, probes, bounds);
    expect(region).toEqual({ top: 20, left: 20, bottom: 21, right: 21 });
  });

  it('stops at the given bounds', () => {
    const region = expandDataRegion(
      { top: 3, left: 2, bottom: 3, right: 2 },
      probes,
      { maxRow: 4, maxCol: 2 }
    );
    expect(region).toEqual({ top: 2, left: 1, bottom: 4, right: 2 });
  });
});
