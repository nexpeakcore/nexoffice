/**
 * Shared init-state machinery for the four generated wasm modules (opc /
 * layout / edit / parse). The binaries are external assets (gitignored,
 * rebuilt by scripts/build-docx-wasm.ts, shipped in dist/generated/) —
 * never embedded base64.
 *
 * Two ways a module becomes ready:
 *  - `preload()` (async): browsers and workers — passes the asset URL to the
 *    wasm-bindgen glue, which fetches and instantiates it.
 *  - `ensure()` (sync): Node/Bun/SSR — reads the asset from disk via
 *    `process.getBuiltinModule` (no static `node:fs` import, so browser
 *    bundlers never see a node builtin) and feeds `initSync`. In a browser
 *    without a prior `preload()` this throws with a call-to-action.
 *
 * IMPORTANT — URL geometry: every `new URL('./generated/…', import.meta.url)`
 * literal must live in a module physically inside `src/wasm/`, and each loader
 * is its own root-named tsup entry. That keeps the relative path valid in BOTH
 * layouts: `src/wasm/*` next to `src/wasm/generated/` in source mode, and
 * root-level chunks next to `dist/generated/` in package builds (the xlsx
 * package established the pattern).
 */

export type WasmSyncInput = BufferSource | WebAssembly.Module;
export type WasmAsyncInput = RequestInfo | URL | Response | BufferSource | WebAssembly.Module;

interface NodeFsLike {
  readFileSync(path: string): Uint8Array;
}
interface NodeUrlLike {
  fileURLToPath(url: URL): string;
}

function builtinModule<T>(name: string): T | undefined {
  const proc = (
    globalThis as {
      process?: { getBuiltinModule?: (id: string) => unknown };
    }
  ).process;
  if (typeof proc?.getBuiltinModule !== 'function') return undefined;
  try {
    return proc.getBuiltinModule(name) as T;
  } catch {
    return undefined;
  }
}

/** Read a `file:` wasm asset synchronously from disk; undefined off-Node. */
export function readWasmSync(url: URL): Uint8Array | undefined {
  if (url.protocol !== 'file:') return undefined;
  const fs = builtinModule<NodeFsLike>('node:fs');
  const nodeUrl = builtinModule<NodeUrlLike>('node:url');
  if (!fs || !nodeUrl) return undefined;
  try {
    return fs.readFileSync(nodeUrl.fileURLToPath(url));
  } catch {
    return undefined;
  }
}

/**
 * Compile a wasm asset into a `WebAssembly.Module` that several agents can
 * instantiate. A `Module` is structured-cloneable to a dedicated worker, so
 * compiling once here and handing the result to both the main thread and the
 * worker replaces two independent compilations of the same binary — for
 * docx-edit that is ~11MB of bytecode, and its machine code, twice over.
 */
export async function compileWasmAsset(url: URL): Promise<WebAssembly.Module> {
  const bytes = readWasmSync(url);
  if (bytes) return WebAssembly.compile(bytes);
  if (typeof WebAssembly.compileStreaming === 'function') {
    try {
      return await WebAssembly.compileStreaming(fetch(url));
    } catch {
      // A response without `application/wasm` cannot stream; buffer instead.
    }
  }
  const response = await fetch(url);
  return WebAssembly.compile(await response.arrayBuffer());
}

export interface WasmModuleState {
  /** Async init from the packaged asset URL (or an explicit override). */
  preload(input?: WasmAsyncInput): Promise<void>;
  /** Whether a module is already instantiated, so callers can skip setup. */
  initialized(): boolean;
  /** Sync guard used by every call site; disk-inits on Node/Bun, throws in a browser before `preload()`. */
  ensure(): void;
}

function syncInput(input: WasmAsyncInput | undefined): WasmSyncInput | undefined {
  if (input === undefined) return undefined;
  if (input instanceof WebAssembly.Module) return input;
  if (input instanceof ArrayBuffer || ArrayBuffer.isView(input)) return input;
  return undefined;
}

export function createWasmModuleState(options: {
  label: string;
  preloadName: string;
  assetUrl: () => URL;
  initAsync: (input: { module_or_path: WasmAsyncInput | Promise<WasmAsyncInput> }) => Promise<unknown>;
  initSync: (input: { module: WasmSyncInput }) => unknown;
}): WasmModuleState {
  let initialized = false;
  let pending: Promise<void> | undefined;

  return {
    initialized(): boolean {
      return initialized;
    },
    preload(input?: WasmAsyncInput): Promise<void> {
      if (initialized) return Promise.resolve();
      if (pending) return pending;
      // Init synchronously whenever the bytes are already at hand (explicit
      // BufferSource/Module, or the asset on local disk). A pending async
      // init is a hazard: the glue's re-init guard is checked before its
      // fetch, so a sync init landing inside that window is later clobbered
      // by a second, fresh instance while sessions keep pointers into the
      // first. No async window, no race.
      const bytes = syncInput(input) ?? (input === undefined ? readWasmSync(options.assetUrl()) : undefined);
      if (bytes) {
        try {
          options.initSync({ module: bytes });
        } catch (error) {
          return Promise.reject(error instanceof Error ? error : new Error(String(error)));
        }
        initialized = true;
        return Promise.resolve();
      }
      pending = options
        .initAsync({ module_or_path: input ?? options.assetUrl() })
        .then(
          () => {
            initialized = true;
          },
          (error: unknown) => {
            pending = undefined;
            throw error instanceof Error ? error : new Error(String(error));
          }
        );
      return pending;
    },
    ensure(): void {
      if (initialized) return;
      if (pending) {
        throw new Error(
          `${options.label} wasm preload is still in flight; await ${options.preloadName}() before ` +
            'synchronous use (a sync init now would be clobbered when the async init lands)'
        );
      }
      const bytes = readWasmSync(options.assetUrl());
      if (bytes) {
        options.initSync({ module: bytes });
        initialized = true;
        return;
      }
      throw new Error(
        `${options.label} wasm is not initialized; await ${options.preloadName}() before first use ` +
          '(browser builds load the external wasm asset asynchronously)'
      );
    },
  };
}
