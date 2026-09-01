import { describe, expect, test } from 'bun:test';
import { isSupersededSessionError, RustDisplayListSourceError } from './rustDisplayList';
import { YrsSessionDestroyedError } from '../../yrs';

describe('isSupersededSessionError', () => {
  test('sees through the display-list wrapper to the destroyed session', () => {
    // What the renderer actually catches: the build wraps whatever the engine
    // threw, and before this the wrapper flattened the cause into a string,
    // so the caller could not tell a swapped document from a broken one.
    const wrapped = new RustDisplayListSourceError(
      'build',
      new YrsSessionDestroyedError('build_display_list_frame')
    );
    expect(isSupersededSessionError(wrapped)).toBe(true);
  });

  test('matches by name across package copies', () => {
    const foreign = new Error('yrs session was destroyed');
    foreign.name = 'YrsSessionDestroyedError';
    expect(isSupersededSessionError(new RustDisplayListSourceError('build', foreign))).toBe(true);
  });

  test('does not swallow a real failure', () => {
    expect(isSupersededSessionError(new RustDisplayListSourceError('build', new Error('boom')))).toBe(
      false
    );
    expect(isSupersededSessionError(new Error('null pointer passed to rust'))).toBe(false);
    expect(isSupersededSessionError(undefined)).toBe(false);
  });

  test('terminates on a cause cycle', () => {
    const a = new Error('a');
    const b = new Error('b');
    (a as { cause?: unknown }).cause = b;
    (b as { cause?: unknown }).cause = a;
    expect(isSupersededSessionError(a)).toBe(false);
  });
});
