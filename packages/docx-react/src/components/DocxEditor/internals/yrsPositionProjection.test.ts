import { describe, expect, test } from 'bun:test';
import {
  createYrsInputPositionMap,
  displayPositionToYrsLoc,
  type YrsSession,
  type YrsStoryOutlineSegment,
} from '@betteroffice/docx/yrs';
import { YrsPositionProjection } from './yrsPositionProjection';

const segments: Record<string, YrsStoryOutlineSegment[]> = {
  body: [
    { kind: 'text', len: 6 },
    { kind: 'pilcrow', paraId: 'body:p0' },
    {
      kind: 'embed',
      embedKind: 'table',
      payload: {
        grid: [100],
        rows: [{ cells: [{ story: 'body:t0:r0c0', tcPr: {} }] }],
      },
      attributes: {},
    },
    { kind: 'text', len: 5 },
    { kind: 'pilcrow', paraId: 'body:p1' },
  ],
  'body:t0:r0c0': [
    { kind: 'text', len: 4 },
    { kind: 'pilcrow', paraId: 'cell:p0' },
  ],
};

const session = {
  storyOutline: (story: string) => segments[story] ?? [],
} as unknown as YrsSession;

describe('YrsPositionProjection', () => {
  test('maps post-table positions back to the root story input map', () => {
    const projection = new YrsPositionProjection(session, 'body');
    const tableEmbedUnit = 1;
    const map = createYrsInputPositionMap('body', [
      { paraId: 'body:p0', length: 'before'.length },
      { paraId: 'body:p1', length: tableEmbedUnit + 'after'.length },
    ]);
    const target = projection.targetAt(23);
    const loc = displayPositionToYrsLoc(map, target.displayPosition);

    expect(target).toEqual({ story: 'body', displayPosition: 12 });
    expect(loc).toEqual({
      story: 'body',
      paraId: 'body:p1',
      offset: 3,
    });
    expect(projection.positionForLoc(loc!)).toBe(23);
  });

  test('keeps table cell positions scoped to the cell input map', () => {
    const projection = new YrsPositionProjection(session, 'body');
    const map = createYrsInputPositionMap('body:t0:r0c0', [
      { paraId: 'cell:p0', length: 4 },
    ]);
    const target = projection.targetAt(12);

    expect(target).toMatchObject({
      story: 'body:t0:r0c0',
      displayPosition: 1,
    });
    const loc = displayPositionToYrsLoc(map, target.displayPosition);
    expect(loc).toEqual({
      story: 'body:t0:r0c0',
      paraId: 'cell:p0',
      offset: 0,
    });
    expect(projection.positionForLoc(loc!)).toBe(12);
  });
});
