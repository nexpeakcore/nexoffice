import { describe, expect, test } from 'bun:test';
import { isResidentInputText } from './residentTextContract';

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
});
