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
import type { FilterSpec, RegionProbes } from './filterSpec';

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

  it('cannot derive text for formula cells without an engine filterText', () => {
    expect(rawFilterText({ input: '=SUM(A1:A2)', isFormula: true })).toBeNull();
  });

  it('uses the engine filterText verbatim when the core provides it', () => {
    expect(rawFilterText({ input: '=SUM(A1:A2)', isFormula: true, filterText: '15' })).toBe('15');
    expect(rawFilterText({ input: '=A1&"x"', isFormula: true, filterText: '10x' })).toBe('10x');
    expect(rawFilterText({ input: '=1/0', isFormula: true, filterText: '#DIV/0!' })).toBe(
      '#DIV/0!'
    );
    expect(rawFilterText({ input: '=IF(A1,"","x")', isFormula: true, filterText: '' })).toBe('');
    expect(rawFilterText({ input: "'123", isFormula: false, filterText: '123' })).toBe('123');
  });
});

describe('collectFilterValues', () => {
  it('dedupes, flags blanks, skips opaque formulas, and sorts numerically', () => {
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

  it('includes formula cells by their engine filterText, deduped against literals', () => {
    const collected = collectFilterValues([
      { input: '15', isFormula: false, filterText: '15' },
      { input: '=SUM(A1:A2)', isFormula: true, filterText: '15' },
      { input: '=B1*2', isFormula: true, filterText: '30' },
      { input: '=IF(A1>0,"","x")', isFormula: true, filterText: '' },
    ]);
    expect(collected.values).toEqual(['15', '30']);
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

  it('keeps active selections that fall beyond the cap', () => {
    const cells = Array.from({ length: 5 }, (_, i) => ({
      input: `v${i}`,
      isFormula: false,
    }));
    const collected = collectFilterValues(cells, 2, ['v4']);
    expect(collected.values).toEqual(['v0', 'v1', 'v4']);
    expect(collected.truncated).toBe(true);
  });

  it('does not let selections eat the collection budget', () => {
    const cells = Array.from({ length: 5 }, (_, i) => ({
      input: `v${i}`,
      isFormula: false,
    }));
    const collected = collectFilterValues(cells, 3, ['v0', 'v3']);
    expect(collected.values).toEqual(['v0', 'v1', 'v2', 'v3', 'v4']);
    expect(collected.truncated).toBe(false);
  });

  it('surfaces selections that no longer exist in the column', () => {
    const collected = collectFilterValues([{ input: 'b', isFormula: false }], 10, ['a', 'z']);
    expect(collected.values).toEqual(['a', 'b', 'z']);
    expect(collected.truncated).toBe(false);
  });

  it('ignores a blank selection entry, which blanks state already carries', () => {
    const collected = collectFilterValues([{ input: '', isFormula: false }], 10, ['']);
    expect(collected.values).toEqual([]);
    expect(collected.hasBlanks).toBe(true);
  });

  it('keeps every selection visible when the column has no body rows left', () => {
    const collected = collectFilterValues([], 10, ['b', 'a']);
    expect(collected.values).toEqual(['a', 'b']);
    expect(collected.hasBlanks).toBe(false);
    expect(collected.truncated).toBe(false);
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

  it('never widens an all-checked truncated list to unconstrained', () => {
    expect(resolveCriteria(values, new Set(values), true, false)).toEqual({
      values: ['a', 'b', 'c'],
      showBlanks: true,
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

describe('unsupported criteria', () => {
  const spec: FilterSpec = {
    range: { startRow: 0, startCol: 0, endRow: 9, endCol: 2 },
    columns: [
      { col: 0, values: ['a'], showBlanks: false },
      { col: 1, values: null, showBlanks: true, unsupported: '<top10 val="5"/>' },
    ],
  };

  it('keeps another column preserved xml when one column is edited', () => {
    const next = withColumnCriteria(spec, 0, ['b'], false);
    expect(next.columns.find((c) => c.col === 1)?.unsupported).toBe('<top10 val="5"/>');
  });

  it('drops the preserved xml when that column is edited', () => {
    const next = withColumnCriteria(spec, 1, ['x'], false);
    expect(next.columns.find((c) => c.col === 1)?.unsupported).toBeUndefined();
  });

  it('reports a preserved column as constrained', () => {
    expect(hasCriteria(columnCriteria(spec, 1))).toBe(true);
    expect(hasCriteria(columnCriteria(spec, 2))).toBe(false);
  });
});
