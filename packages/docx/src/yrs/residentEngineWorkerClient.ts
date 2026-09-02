import type {
  YrsEngineApplyProfile,
  YrsResidentCaretSnapshot,
  YrsResidentWorkerOpen,
  YrsResidentWorkerSnapshot,
  YrsSelection,
} from './index';
import { registerMemoryReader } from '../diagnostics';
import type { ResidentCaretPaintStyle } from './residentCaret';
import type {
  ResidentEngineWorkerMessage,
  ResidentEngineWorkerRequest,
  ResidentEngineWorkerRequestWithoutId,
  ResidentEngineWorkerResponse,
  ResidentEngineWorkerStage,
  YrsOpenProfile,
} from './residentEngineWorkerProtocol';
import { residentWorkerSilenceBudgetMs } from './residentWorkerDeadline';

export interface ResidentEngineWorkerFrame {
  frame: Uint8Array;
  updates: Uint8Array[];
  engineMs: number;
  workerTotalMs: number;
  /** How long the request waited inside the worker before it was handled. */
  workerQueuedMs: number;
  engineProfile?: YrsEngineApplyProfile;
  /** Present on the reply to `open` and `bootstrap`. */
  openProfile?: YrsOpenProfile;
  caret: YrsResidentCaretSnapshot;
  selection: YrsSelection | null;
  /** The presented frame carries the worker-painted caret line. */
  caretPainted: boolean;
  replayMs: number;
  replayedPages: number;
  layoutRevision: number;
  /** Wasm heap held in the worker thread, folded into the renderer process
   * by Electron's metrics and only separable from here. */
  heapBytes: number;
}

export interface ResidentEngineOffscreenPage {
  pageId: string;
  canvas: OffscreenCanvas;
}

export interface ResidentEngineWorkerApplyResult extends ResidentEngineWorkerFrame {
  applied: true;
}

type PendingRequest = {
  resolve(response: ResidentEngineWorkerResponse & { ok: true }): void;
  reject(error: Error): void;
  timeout: ReturnType<typeof setTimeout> | null;
  /** Restarts this request's silence budget after a sign of life. */
  renew(stage: ResidentEngineWorkerStage): void;
};

/** Dedicated-worker owner for resident input, pagination, and FrameDelta output. */
export class ResidentEngineWorkerClient {
  private readonly worker: Worker;
  private readonly pending = new Map<number, PendingRequest>();
  /** Resolves once the shared wasm module has been offered to the worker;
   * every request posts behind it so `initWasm` is always the first message. */
  private readonly wasmReady: Promise<void>;
  private nextId = 1;
  private destroyed = false;
  private ready = false;
  private revision = 0;
  private remoteVector: Uint8Array | null = null;
  private appliedFontsRevision: number | null = null;

  constructor() {
    this.worker = new Worker(new URL('./residentEngineWorker.mjs', import.meta.url), {
      type: 'module',
      name: 'openooxml-resident-engine',
    });
    this.worker.onmessage = (event: MessageEvent<ResidentEngineWorkerMessage>) => {
      const message = event.data;
      if ('progress' in message) {
        // Not a result: the worker is telling us it is still running, which is
        // the only way to know that from outside a blocking wasm call.
        this.pending.get(message.id)?.renew(message.progress);
        return;
      }
      const response = message;
      if (response.ok && response.stateVector) {
        this.remoteVector = new Uint8Array(response.stateVector);
      }
      const pending = this.pending.get(response.id);
      if (!pending) return;
      this.pending.delete(response.id);
      if (pending.timeout) clearTimeout(pending.timeout);
      if (response.ok) pending.resolve(response);
      else pending.reject(residentWorkerError(response.error, response.residentUnavailable));
    };
    this.worker.onerror = (event) => {
      this.failAll(new Error(`Resident engine worker failed: ${event.message}`));
      this.ready = false;
    };
    this.worker.onmessageerror = () => {
      this.failAll(new Error('Resident engine worker returned an unreadable message'));
      this.ready = false;
    };
    this.wasmReady = this.shareCompiledModule();
  }

  /**
   * Shares this agent's compiled editing-core module with the worker so the
   * two threads do not compile ~11MB each. Best-effort: on failure the worker
   * loads the asset itself.
   */
  private async shareCompiledModule(): Promise<void> {
    try {
      const { compileEditWasmModule } = await import('./wasm/index');
      const module = await compileEditWasmModule();
      if (this.destroyed) return;
      const message: ResidentEngineWorkerRequest = {
        id: this.nextId++,
        type: 'initWasm',
        module,
      };
      this.worker.postMessage(message);
    } catch {
      // Optimization only — the worker's own preload path still works.
    }
  }

  isReady(): boolean {
    return this.ready;
  }

  layoutRevision(): number {
    return this.revision;
  }

  /** The worker replica's last reported yrs state vector (null before any). */
  remoteStateVector(): Uint8Array | null {
    return this.remoteVector;
  }

  /** The fonts revision this worker last applied (null before bootstrap). */
  syncedFontsRevision(): number | null {
    return this.appliedFontsRevision;
  }

  /**
   * Hands the worker a seeded-but-unlaid-out document and lets it paginate.
   * Unlike `bootstrap`, nothing on the main thread has lowered, measured or
   * paginated it first, so this is the whole cost once rather than twice.
   */
  async open(
    open: YrsResidentWorkerOpen,
    layoutInput: string,
    extras: string
  ): Promise<ResidentEngineWorkerFrame> {
    const fontsRevision = open.fontsRevision;
    const response = await this.request(
      { type: 'open', open, layoutInput, extras },
      [open.state.buffer, ...open.fonts.map((font) => font.buffer)],
      true
    );
    const result = frameResult(response);
    this.recordSync(response, fontsRevision);
    this.ready = true;
    this.revision = result.layoutRevision;
    return result;
  }

  async bootstrap(
    snapshot: YrsResidentWorkerSnapshot,
    extras: string
  ): Promise<ResidentEngineWorkerFrame> {
    const fontsRevision = snapshot.fontsRevision;
    const response = await this.request(
      {
        type: 'bootstrap',
        snapshot,
        extras,
        expectedFrameEpoch: 0,
      },
      snapshotTransfers(snapshot),
      true
    );
    const result = frameResult(response);
    this.recordSync(response, fontsRevision);
    this.ready = true;
    this.revision = result.layoutRevision;
    return result;
  }

  /**
   * Re-paginates a document the worker already owns. No state travels: the
   * worker's replica is kept current by `invalidate`, so only the region
   * request goes over.
   */
  async relayout(
    layoutInput: string,
    extras: string,
    expectedFrameEpoch: number,
    paintCaret = false
  ): Promise<ResidentEngineWorkerFrame> {
    const response = await this.request({
      type: 'relayout',
      layoutInput,
      extras,
      expectedFrameEpoch,
      paintCaret,
    });
    const result = frameResult(response);
    this.ready = true;
    this.revision = result.layoutRevision;
    return result;
  }

  async sync(
    snapshot: YrsResidentWorkerSnapshot,
    extras: string,
    expectedFrameEpoch: number,
    paintCaret = false
  ): Promise<ResidentEngineWorkerFrame> {
    const fontsRevision = snapshot.fontsRevision;
    const response = await this.request(
      { type: 'sync', snapshot, extras, expectedFrameEpoch, paintCaret },
      snapshotTransfers(snapshot)
    );
    const result = frameResult(response);
    this.recordSync(response, fontsRevision);
    this.ready = true;
    this.revision = result.layoutRevision;
    return result;
  }

  async buildFrame(
    extras: string,
    expectedFrameEpoch: number,
    paintCaret = false
  ): Promise<ResidentEngineWorkerFrame> {
    const result = frameResult(
      await this.request({ type: 'buildFrame', extras, expectedFrameEpoch, paintCaret })
    );
    return result;
  }

  /** Requests posted and not yet answered — what a new one waits behind. */
  inFlight(): number {
    return this.pending.size;
  }

  async applyInput(
    text: string,
    selection: YrsSelection,
    expectedFrameEpoch: number,
    profile = false,
    paintCaret = false
  ): Promise<ResidentEngineWorkerApplyResult | { applied: false }> {
    if (!this.ready) return { applied: false };
    try {
      const result = frameResult(
        await this.request({
          type: 'applyInput',
          text,
          selection,
          expectedFrameEpoch,
          profile,
          paintCaret,
        })
      );
      return { applied: true, ...result };
    } catch (error) {
      if (error instanceof ResidentWorkerUnavailableError) return { applied: false };
      throw error;
    }
  }

  async applyDelete(
    direction: 'backward' | 'forward',
    selection: YrsSelection,
    expectedFrameEpoch: number,
    profile = false,
    paintCaret = false
  ): Promise<ResidentEngineWorkerApplyResult | { applied: false }> {
    if (!this.ready) return { applied: false };
    try {
      const result = frameResult(
        await this.request({
          type: 'applyDelete',
          direction,
          selection,
          expectedFrameEpoch,
          profile,
          paintCaret,
        })
      );
      return { applied: true, ...result };
    } catch (error) {
      if (error instanceof ResidentWorkerUnavailableError) return { applied: false };
      throw error;
    }
  }

  /** Drop the worker-painted caret line by re-presenting the caret page's
   * retained raster. Fire-and-forget and idempotent. */
  eraseCaret(): void {
    if (this.destroyed) return;
    const id = this.nextId++;
    this.post({ id, type: 'eraseCaret' });
  }

  /**
   * The pages the host is looking at, so the worker can send the rest as
   * geometry alone. Fire-and-forget: the next frame carries the change, and a
   * window that arrives late costs one frame of stale content, never
   * correctness.
   */
  setPageWindow(start: number, count: number): void {
    if (this.destroyed) return;
    const id = this.nextId++;
    this.post({ id, type: 'setPageWindow', start, count });
  }

  invalidate(update: Uint8Array, selection: YrsSelection | null): void {
    if (this.destroyed) return;
    this.ready = false;
    const owned = update.slice();
    const id = this.nextId++;
    this.post({ id, type: 'applyUpdate', update: owned, selection }, [owned.buffer]);
  }

  async attachCanvases(
    pages: ResidentEngineOffscreenPage[],
    activePageIds: string[],
    devicePixelRatio: number,
    zoom: number,
    caretStyle: ResidentCaretPaintStyle
  ): Promise<void> {
    const canvases = pages.map((page) => page.canvas);
    await this.request(
      { type: 'attachCanvases', pages, activePageIds, devicePixelRatio, zoom, caretStyle },
      canvases
    );
  }

  destroy(): void {
    if (this.destroyed) return;
    this.destroyed = true;
    const id = this.nextId++;
    const message: ResidentEngineWorkerRequest = { id, type: 'destroy' };
    this.worker.postMessage(message);
    this.worker.terminate();
    forgetWorkerHeap();
    this.failAll(new Error('Resident engine worker was destroyed'));
    this.ready = false;
  }

  /** Queues every message behind the shared-module handshake, preserving order. */
  private post(message: ResidentEngineWorkerRequest, transfer: Transferable[] = []): void {
    void this.wasmReady.then(() => {
      if (this.destroyed) return;
      this.worker.postMessage(message, transfer);
    });
  }

  /**
   * `deadlined` requests are given up on when the worker goes quiet for longer
   * than its current stage allows — not when they take too long. A worker that
   * reports it is laying out has proved it is alive, and killing it there would
   * throw away the work and hand the whole document back to a main thread that
   * is no faster.
   */
  private request(
    request: ResidentEngineWorkerRequestWithoutId,
    transfer: Transferable[] = [],
    deadlined = false
  ): Promise<ResidentEngineWorkerResponse & { ok: true }> {
    if (this.destroyed) return Promise.reject(new Error('Resident engine worker was destroyed'));
    const id = this.nextId++;
    return new Promise((resolve, reject) => {
      let stage: ResidentEngineWorkerStage | null = null;
      const abandon = (budgetMs: number) => {
        const pending = this.pending.get(id);
        if (!pending) return;
        this.pending.delete(id);
        const since = stage ? `after reaching ${stage}` : 'before it started';
        const error = new Error(
          `Resident engine worker went quiet on ${request.type} ${since}, for over ${budgetMs}ms`
        );
        pending.reject(error);
        this.failAll(error);
        this.ready = false;
        this.destroyed = true;
        this.worker.terminate();
        forgetWorkerHeap();
      };
      const arm = (): ReturnType<typeof setTimeout> | null => {
        if (!deadlined) return null;
        const budgetMs = residentWorkerSilenceBudgetMs(stage);
        return setTimeout(() => abandon(budgetMs), budgetMs);
      };
      const renew = (next: ResidentEngineWorkerStage): void => {
        const pending = this.pending.get(id);
        if (!pending) return;
        if (pending.timeout) clearTimeout(pending.timeout);
        stage = next;
        pending.timeout = arm();
      };
      this.pending.set(id, { resolve, reject, timeout: arm(), renew });
      this.post({ ...request, id } as ResidentEngineWorkerRequest, transfer);
    });
  }

  /** Record a successfully applied bootstrap/sync payload's fonts revision.
   * The state vector is tracked centrally in `onmessage`. */
  private recordSync(
    _response: ResidentEngineWorkerResponse & { ok: true },
    fontsRevision: number
  ): void {
    this.appliedFontsRevision = fontsRevision;
  }

  private failAll(error: Error): void {
    for (const pending of this.pending.values()) {
      if (pending.timeout) clearTimeout(pending.timeout);
      pending.reject(error);
    }
    this.pending.clear();
  }
}

class ResidentWorkerUnavailableError extends Error {}

function residentWorkerError(message: string, unavailable = false): Error {
  return unavailable ? new ResidentWorkerUnavailableError(message) : new Error(message);
}

function snapshotTransfers(snapshot: YrsResidentWorkerSnapshot): Transferable[] {
  return [snapshot.state.buffer, ...snapshot.fonts.map((font) => font.buffer)];
}

// The worker reports its heap with every frame it returns, so the reader below
// is a cache of the last report rather than a live probe: nothing on the main
// thread can reach across into the worker's wasm memory to measure it.
let workerHeapBytes = 0;
registerMemoryReader('wasm · resident engine (worker)', () => workerHeapBytes);

/**
 * A terminated worker holds nothing. Without this the last frame's figure
 * outlives the worker, so a main-thread fallback or a closed document keeps
 * reporting memory that has already been released.
 */
function forgetWorkerHeap(): void {
  workerHeapBytes = 0;
}

function frameResult(
  response: ResidentEngineWorkerResponse & { ok: true }
): ResidentEngineWorkerFrame {
  if (response.heapBytes !== undefined) workerHeapBytes = response.heapBytes;
  if (!response.frame) throw new Error('Resident engine worker response omitted its FrameDelta');
  if (!response.caret) throw new Error('Resident engine worker response omitted its caret snapshot');
  if (response.selection === undefined) {
    throw new Error('Resident engine worker response omitted its selection');
  }
  return {
    frame: new Uint8Array(response.frame),
    updates: (response.updates ?? []).map((update) => new Uint8Array(update)),
    engineMs: response.engineMs ?? 0,
    workerTotalMs: response.workerTotalMs ?? 0,
    workerQueuedMs: response.workerQueuedMs ?? 0,
    engineProfile: response.engineProfile,
    openProfile: response.openProfile,
    caret: response.caret,
    selection: response.selection,
    caretPainted: response.caretPainted ?? false,
    replayMs: response.replayMs ?? 0,
    replayedPages: response.replayedPages ?? 0,
    layoutRevision: response.layoutRevision ?? 0,
    heapBytes: workerHeapBytes,
  };
}

export function canUseResidentEngineWorker(): boolean {
  return typeof Worker !== 'undefined';
}
