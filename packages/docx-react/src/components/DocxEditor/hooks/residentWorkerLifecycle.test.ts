import { describe, expect, test } from 'bun:test';
import { workerServesEngine } from './residentWorkerLifecycle';

describe('workerServesEngine', () => {
  const engineA = { id: 'a' };
  const engineB = { id: 'b' };

  test('serves the engine it was bootstrapped for', () => {
    expect(workerServesEngine({ engine: engineA }, engineA)).toBe(true);
  });

  test('refuses a worker left over from another document', () => {
    // Open a document under the worker cutoff, then one over it: the second
    // takes the main-thread path, and before this the first document's worker
    // stayed in place and kept answering input.
    expect(workerServesEngine({ engine: engineA }, engineB)).toBe(false);
  });

  test('refuses when either side is absent', () => {
    expect(workerServesEngine(null, engineA)).toBe(false);
    expect(workerServesEngine(undefined, engineA)).toBe(false);
    expect(workerServesEngine({ engine: engineA }, null)).toBe(false);
    expect(workerServesEngine({ engine: engineA }, undefined)).toBe(false);
  });

  test('compares by identity, not by shape', () => {
    expect(workerServesEngine({ engine: { id: 'a' } }, { id: 'a' })).toBe(false);
  });
});
