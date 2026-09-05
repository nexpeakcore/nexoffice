import { describe, expect, it } from 'bun:test';
import type { CellEdit, EditRefusal } from '@betteroffice/xlsx';

import { formulaBarView, refusalNotice, spilledRegion } from './spilledResult';

const anchor: CellEdit = {
  a1: 'F6',
  input: '=SEQUENCE(3)',
  isFormula: true,
  spill: {
    anchor: 'F6',
    range: { start: { row: 5, col: 5 }, end: { row: 7, col: 5 } },
    input: '=SEQUENCE(3)',
  },
};

const child: CellEdit = { ...anchor, a1: 'F7', input: '2', isFormula: false };

const plain: CellEdit = { a1: 'A1', input: '7', isFormula: false };

describe('spilledRegion', () => {
  it('outlines the whole result from any cell in it', () => {
    const region = { top: 5, left: 5, bottom: 7, right: 5 };
    expect(spilledRegion(anchor)).toEqual(region);
    expect(spilledRegion(child)).toEqual(region);
  });

  it('has nothing to outline for an ordinary cell', () => {
    expect(spilledRegion(plain)).toBeNull();
    expect(spilledRegion(null)).toBeNull();
  });
});

describe('formulaBarView', () => {
  it('shows the anchor formula for a cell the result filled', () => {
    expect(formulaBarView(child, null)).toEqual({
      value: '=SEQUENCE(3)',
      borrowedFrom: 'F6',
    });
  });

  it('does not call the anchor a borrower of its own formula', () => {
    expect(formulaBarView(anchor, null)).toEqual({
      value: '=SEQUENCE(3)',
      borrowedFrom: null,
    });
  });

  it('leaves an ordinary cell alone', () => {
    expect(formulaBarView(plain, null)).toEqual({ value: '7', borrowedFrom: null });
    expect(formulaBarView(null, null)).toEqual({ value: '', borrowedFrom: null });
  });

  it('lets a draft the reader started win over both', () => {
    expect(formulaBarView(child, '=1+')).toEqual({ value: '=1+', borrowedFrom: null });
    expect(formulaBarView(child, '')).toEqual({ value: '', borrowedFrom: null });
  });
});

describe('refusalNotice', () => {
  it('names the cell and the formula that owns it', () => {
    const refusal: EditRefusal = {
      kind: 'spilledCell',
      at: 'F7',
      anchor: 'F6',
      range: null,
    };
    expect(refusalNotice(refusal)).toEqual({
      key: 'spill.cannotEdit',
      vars: { cell: 'F7', anchor: 'F6' },
    });
  });

  it('spells the torn region in A1, past the 26th column', () => {
    const refusal: EditRefusal = {
      kind: 'spillTorn',
      anchor: 'AA1',
      range: { start: { row: 0, col: 26 }, end: { row: 2, col: 27 } },
    };
    expect(refusalNotice(refusal)).toEqual({
      key: 'spill.cannotTear',
      vars: { anchor: 'AA1', range: 'AA1:AB3' },
    });
  });

  it('says nothing when nothing was refused', () => {
    expect(refusalNotice(undefined)).toBeNull();
  });
});
