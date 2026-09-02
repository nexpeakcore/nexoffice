import { describe, expect, test } from 'bun:test';

import {
  COMMITTED_COMPOSITION_TIMEOUT_MS,
  CommittedCompositionHold,
} from './committedComposition';

const WORD = { text: 'xin', left: 100, top: 200, height: 20 };

describe('CommittedCompositionHold', () => {
  test('holds the commit until the frame carrying it is presented', () => {
    const hold = new CommittedCompositionHold();
    hold.hold(WORD, 7, 0);
    expect(hold.visible()).toEqual(WORD);

    expect(hold.onFramePresented(6)).toBe(false);
    expect(hold.visible()).toEqual(WORD);

    expect(hold.onFramePresented(7)).toBe(true);
    expect(hold.visible()).toBeNull();
  });

  test('a later frame also releases it, so a skipped epoch cannot strand the text', () => {
    const hold = new CommittedCompositionHold();
    hold.hold(WORD, 7, 0);
    expect(hold.onFramePresented(9)).toBe(true);
    expect(hold.visible()).toBeNull();
  });

  test('a frame with no epoch never releases it', () => {
    const hold = new CommittedCompositionHold();
    hold.hold(WORD, 7, 0);
    expect(hold.onFramePresented(null)).toBe(false);
    expect(hold.visible()).toEqual(WORD);
  });

  test('a commit that awaits no frame is released only by time', () => {
    const hold = new CommittedCompositionHold();
    hold.hold(WORD, null, 0);
    expect(hold.onFramePresented(42)).toBe(false);
    expect(hold.visible()).toEqual(WORD);

    expect(hold.tick(COMMITTED_COMPOSITION_TIMEOUT_MS - 1)).toBe(false);
    expect(hold.tick(COMMITTED_COMPOSITION_TIMEOUT_MS)).toBe(true);
    expect(hold.visible()).toBeNull();
  });

  test('a frame that never comes stops painting over a page without the text', () => {
    const hold = new CommittedCompositionHold();
    hold.hold(WORD, 7, 1000);
    expect(hold.tick(1000 + COMMITTED_COMPOSITION_TIMEOUT_MS - 1)).toBe(false);
    expect(hold.tick(1000 + COMMITTED_COMPOSITION_TIMEOUT_MS)).toBe(true);
    expect(hold.visible()).toBeNull();
  });

  test('an empty commit holds nothing and clears whatever was held', () => {
    const hold = new CommittedCompositionHold();
    hold.hold(WORD, 7, 0);
    hold.hold({ ...WORD, text: '' }, 8, 1);
    expect(hold.visible()).toBeNull();
    expect(hold.onFramePresented(8)).toBe(false);
  });

  test('a fresh commit replaces the one still on screen', () => {
    const hold = new CommittedCompositionHold();
    hold.hold(WORD, 7, 0);
    const next = { text: 'chào', left: 140, top: 200, height: 20 };
    hold.hold(next, 9, 50);
    expect(hold.visible()).toEqual(next);
    expect(hold.onFramePresented(7)).toBe(false);
    expect(hold.onFramePresented(9)).toBe(true);
  });
});
