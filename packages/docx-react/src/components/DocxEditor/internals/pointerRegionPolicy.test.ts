import { describe, expect, test } from 'bun:test';
import { mousedownRegionAction } from './pointerRegionPolicy';

describe('mousedownRegionAction', () => {
  // The regression this pins: outside edit mode, a header mousedown must be
  // ignored for every click count. The old code made an exception for
  // double-clicks, and that exception is the whole bug — the caret path ran
  // with no body position under the point, defaulted to the end of the
  // document, and the scroll that followed moved body content under the
  // cursor before the click handler could recognise the header.
  test('ignores header and footer mousedowns outside edit mode', () => {
    expect(mousedownRegionAction('header', null)).toBe('ignore');
    expect(mousedownRegionAction('footer', null)).toBe('ignore');
    expect(mousedownRegionAction('header', undefined)).toBe('ignore');
  });

  test('lets the body take its own mousedowns', () => {
    expect(mousedownRegionAction('body', null)).toBe('body');
    expect(mousedownRegionAction(null, null)).toBe('body');
  });

  test('keeps editing when the mousedown lands in the region being edited', () => {
    expect(mousedownRegionAction('header', 'header')).toBe('body');
    expect(mousedownRegionAction('footer', 'footer')).toBe('body');
  });

  test('leaves edit mode when the mousedown lands anywhere else', () => {
    expect(mousedownRegionAction('body', 'header')).toBe('exit-hf');
    expect(mousedownRegionAction('footer', 'header')).toBe('exit-hf');
    expect(mousedownRegionAction(null, 'footer')).toBe('exit-hf');
  });
});
