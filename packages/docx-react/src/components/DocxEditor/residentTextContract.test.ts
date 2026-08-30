import { describe, expect, test } from 'bun:test';
import { isBatchableInputText, isResidentInputText } from './residentTextContract';

describe('isResidentInputText', () => {
  test('carries the scripts the ASCII-only rule used to exclude', () => {
    // Each of these is pinned against a full rebuild in the engine by
    // `resident_region_fast_path_matches_a_full_pass_on_non_ascii_text`.
    expect(isResidentInputText('\u0111\u01b0\u1ee3c \u0111o')).toBe(true);
    expect(isResidentInputText('\u5bbd\u5ea6\u51b3\u5b9a')).toBe(true);
    expect(isResidentInputText('\u0645\u0631\u062d\u0628\u0627')).toBe(true);
    expect(isResidentInputText('\u{1f642}\u{1f680}')).toBe(true);
    expect(isResidentInputText('e\u0301')).toBe(true);
  });

  test('still carries plain ASCII', () => {
    expect(isResidentInputText('hello world')).toBe(true);
    expect(isResidentInputText(' ')).toBe(true);
  });

  test('refuses text that is not text', () => {
    expect(isResidentInputText('')).toBe(false);
    expect(isResidentInputText('a\nb')).toBe(false);
    expect(isResidentInputText('a\rb')).toBe(false);
    expect(isResidentInputText('a\tb')).toBe(false);
    expect(isResidentInputText('a\u2028b')).toBe(false);
    expect(isResidentInputText('a\u2029b')).toBe(false);
  });

  test('does not coalesce what the fast path newly admits', () => {
    // Batching widens the blast radius of the drop path in `insertText`: one
    // lost operation loses every character coalesced into it. Vietnamese takes
    // the fast path but is applied one operation at a time until that path
    // cannot lose text.
    expect(isBatchableInputText('hello')).toBe(true);
    expect(isBatchableInputText('\u0111\u01b0\u1ee3c')).toBe(false);
    expect(isBatchableInputText('\u5bbd\u5ea6')).toBe(false);
    expect(isBatchableInputText('\u{1f642}')).toBe(false);
    expect(isBatchableInputText('a\nb')).toBe(false);
  });
});
