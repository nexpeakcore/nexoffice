/**
 * Loader for the docx-edit wasm (the Rust yrs editing core + resident layout
 * engine). External-asset pattern — see ./loadWasmAsset.ts for the init
 * contract and URL-geometry invariant.
 *
 * Reached only via the `src/yrs/` facade's dynamic `import()`, so the glue
 * (and the wasm fetch) stay out of every non-editor bundle path. The facade's
 * async `createYrsSession` awaits {@link preloadEditWasm} before constructing
 * a session, which also covers the resident engine worker (it bootstraps
 * through the same facade inside the worker context).
 */

import wasmInit, {
  initSync,
  EditSession,
  heap_live_bytes,
  heap_peak_bytes,
  heap_reset_peak,
  heap_stats_available,
} from './generated/edit/docx_edit.js';
import { registerHeapProbe } from '../diagnostics';
import { compileWasmAsset, createWasmModuleState, type WasmAsyncInput } from './loadWasmAsset';

const assetUrl = (): URL => new URL('./generated/edit/docx_edit_bg.wasm', import.meta.url);

const state = createWasmModuleState({
  label: 'docx-edit',
  preloadName: 'preloadEditWasm',
  assetUrl,
  initAsync: wasmInit,
  initSync,
});

let compiled: Promise<WebAssembly.Module> | undefined;

/**
 * The compiled editing-core module, compiled at most once per agent. The
 * resident engine worker takes this over `postMessage` instead of fetching
 * and compiling the binary a second time; see `residentEngineWorkerClient`.
 */
export function compileEditWasmModule(): Promise<WebAssembly.Module> {
  compiled ??= compileWasmAsset(assetUrl());
  return compiled;
}

/** Load + instantiate the editing-core wasm (browser path). Idempotent. */
export function preloadEditWasm(input?: WasmAsyncInput): Promise<void> {
  if (input !== undefined || state.initialized()) return state.preload(input);
  // Route the default path through the shared module so a later worker can be
  // handed the very same compilation. Instantiation stays async — this module
  // is far past the 8MB ceiling Blink puts on synchronous instantiation on the
  // main thread (see `syncInput`).
  return compileEditWasmModule().then(
    (module) => state.preload(module),
    () => state.preload()
  );
}

/**
 * Constructs a raw wasm `EditSession` replica after ensuring the module is
 * initialized. The `src/yrs/` facade wraps this in the typed `YrsSession`
 * surface — nothing else should call it.
 */
export function createEditSession(clientId: number): EditSession {
  state.ensure();
  publishHeapProbe();
  return new EditSession(clientId);
}

let heapProbePublished = false;

/**
 * Hands the diagnostics layer this module's allocator counters. They report
 * what the document occupies, which the module's linear-memory size cannot:
 * that figure includes the allocator's free lists and never shrinks.
 */
function publishHeapProbe(): void {
  if (heapProbePublished) return;
  heapProbePublished = true;
  try {
    if (!heap_stats_available()) return;
    registerHeapProbe({
      liveBytes: () => heap_live_bytes(),
      peakBytes: () => heap_peak_bytes(),
      resetPeak: () => heap_reset_peak(),
    });
  } catch {
    // A build without the counters still opens documents.
  }
}

export type { EditSession };
