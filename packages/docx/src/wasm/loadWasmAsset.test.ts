import { describe, expect, it } from 'bun:test';
import { mkdtempSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { pathToFileURL } from 'node:url';

import { createWasmModuleState } from './loadWasmAsset';

function makeState(overrides?: { assetUrl?: () => URL; initAsyncNeverResolves?: boolean }) {
  const calls = { sync: 0, async: 0 };
  let resolveAsync: (() => void) | undefined;
  let rejectAsync: ((error: Error) => void) | undefined;
  const state = createWasmModuleState({
    label: 'test',
    preloadName: 'preloadTestWasm',
    assetUrl: overrides?.assetUrl ?? (() => new URL('https://example.invalid/never.wasm')),
    initAsync: () => {
      calls.async += 1;
      return new Promise<void>((resolve, reject) => {
        resolveAsync = resolve;
        rejectAsync = reject;
        if (!overrides?.initAsyncNeverResolves) resolve();
      });
    },
    initSync: () => {
      calls.sync += 1;
    },
  });
  return {
    state,
    calls,
    resolveAsync: () => resolveAsync?.(),
    rejectAsync: (error: Error) => rejectAsync?.(error),
  };
}

describe('createWasmModuleState', () => {
  it('inits synchronously when preload receives bytes', async () => {
    const { state, calls } = makeState();
    await state.preload(new Uint8Array([0, 97, 115, 109]));
    expect(calls).toEqual({ sync: 1, async: 0 });
    state.ensure();
    expect(calls).toEqual({ sync: 1, async: 0 });
  });

  // Blink refuses a synchronous instantiation of a >8MB module on the main
  // thread, and the module this path exists to share is 11.5MB. A shared
  // module must therefore reach the async init, never initSync.
  it('inits asynchronously when preload receives a compiled module', async () => {
    const { state, calls } = makeState();
    const module = new WebAssembly.Module(new Uint8Array([0, 97, 115, 109, 1, 0, 0, 0]));
    await state.preload(module);
    expect(calls).toEqual({ sync: 0, async: 1 });
    state.ensure();
    expect(calls).toEqual({ sync: 0, async: 1 });
  });

  it('inits synchronously from a local file asset', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'wasm-state-'));
    const asset = join(dir, 'a.wasm');
    writeFileSync(asset, new Uint8Array([1, 2, 3]));
    const { state, calls } = makeState({ assetUrl: () => pathToFileURL(asset) });
    await state.preload();
    expect(calls).toEqual({ sync: 1, async: 0 });
  });

  it('refuses a sync init while an async preload is in flight', async () => {
    const { state, calls, resolveAsync } = makeState({ initAsyncNeverResolves: true });
    const preloading = state.preload('https://example.invalid/module.wasm');
    expect(calls).toEqual({ sync: 0, async: 1 });
    expect(() => state.ensure()).toThrow(/in flight/);
    expect(calls.sync).toBe(0);
    resolveAsync();
    await preloading;
    state.ensure();
    expect(calls).toEqual({ sync: 0, async: 1 });
  });

  it('allows a retry after an async preload failure', async () => {
    const { state, calls, rejectAsync } = makeState({ initAsyncNeverResolves: true });
    const first = state.preload('https://example.invalid/module.wasm');
    rejectAsync(new Error('network down'));
    await expect(first).rejects.toThrow('network down');
    const { promise } = { promise: state.preload(new Uint8Array([7])) };
    await promise;
    expect(calls.sync).toBe(1);
  });
});
