import type {
  YrsEngineApplyProfile,
  YrsResidentCaretSnapshot,
  YrsResidentWorkerOpen,
  YrsResidentWorkerSnapshot,
  YrsSelection,
} from './index';
import type { ResidentCaretPaintStyle } from './residentCaret';

export type ResidentEngineWorkerRequest =
  | {
      /** Hands over the compiled editing-core module; posted before any other request. */
      id: number;
      type: 'initWasm';
      module: WebAssembly.Module;
    }
  | {
      /**
       * Opens the document in the worker, which then owns its layout: it takes
       * the seeded replica, then lowers, measures, paginates and builds the
       * display list once, for good.
       *
       * `bootstrap` below does that same work a second time, after a main
       * thread that has already done all of it — the duplicate is what makes a
       * long document cost two of everything.
       *
       * The replica arrives as state rather than as DOCX bytes on purpose:
       * seeding independently would give the worker its own block keys, and
       * the display positions the main thread maps selections through would
       * then address a different document.
       */
      id: number;
      type: 'open';
      open: YrsResidentWorkerOpen;
      /** The region layout request, as `layoutDocumentWithRegionsVoid` takes it. */
      layoutInput: string;
      extras: string;
    }
  | {
      id: number;
      type: 'bootstrap';
      snapshot: YrsResidentWorkerSnapshot;
      extras: string;
      expectedFrameEpoch: number;
    }
  | {
      /**
       * Re-paginates what the worker already holds. Its replica is kept
       * current by `applyUpdate`, so a document it opened never needs state
       * shipped to it again — only the instruction to lay it out, and the
       * region request in case the page setup changed.
       */
      id: number;
      type: 'relayout';
      layoutInput: string;
      extras: string;
      expectedFrameEpoch: number;
      paintCaret: boolean;
    }
  | {
      id: number;
      type: 'sync';
      snapshot: YrsResidentWorkerSnapshot;
      extras: string;
      expectedFrameEpoch: number;
      paintCaret: boolean;
    }
  | {
      id: number;
      type: 'buildFrame';
      extras: string;
      expectedFrameEpoch: number;
      paintCaret: boolean;
    }
  | {
      id: number;
      type: 'applyInput';
      text: string;
      selection: YrsSelection;
      expectedFrameEpoch: number;
      profile: boolean;
      paintCaret: boolean;
    }
  | {
      id: number;
      type: 'applyDelete';
      direction: 'backward' | 'forward';
      selection: YrsSelection;
      expectedFrameEpoch: number;
      profile: boolean;
      paintCaret: boolean;
    }
  | {
      id: number;
      type: 'applyUpdate';
      update: Uint8Array;
      selection: YrsSelection | null;
    }
  | {
      id: number;
      type: 'attachCanvases';
      pages: Array<{ pageId: string; canvas: OffscreenCanvas }>;
      activePageIds: string[];
      devicePixelRatio: number;
      zoom: number;
      caretStyle: ResidentCaretPaintStyle;
    }
  | {
      /**
       * The pages the host is looking at. Pages outside travel as geometry
       * without content, so a long document's weight stops being something
       * both threads carry in full. `count` below zero clears the window.
       *
       * Fire-and-forget: the next frame carries the change.
       */
      id: number;
      type: 'setPageWindow';
      start: number;
      count: number;
    }
  | { id: number; type: 'eraseCaret' }
  | { id: number; type: 'destroy' };

export type ResidentEngineWorkerRequestWithoutId = ResidentEngineWorkerRequest extends infer Request
  ? Request extends { id: number }
    ? Omit<Request, 'id'>
    : never
  : never;

/**
 * How far a long request has got. Opening a document is one message but many
 * steps, and the host cannot otherwise tell a worker that is busy from one that
 * is broken: a worker inside a wasm call answers nothing at all.
 */
export type ResidentEngineWorkerStage =
  /** The request reached the worker thread — it exists and is running. */
  | 'received'
  /** The wasm module instantiated and a session exists. */
  | 'sessionReady'
  /** The document state is loaded and its fonts are registered. */
  | 'stateLoaded'
  /** Lowering, measuring and pagination have begun — one long wasm call
   * during which the worker cannot answer anything, however healthy. */
  | 'layingOut'
  /** That call returned; the display list frame is being built. */
  | 'laidOut';

/**
 * Where opening a document went, measured in the worker that does the work.
 *
 * A keystroke has had a phase breakdown for a while, which is the only reason
 * its cost could be attributed and cut. Opening had none: the host could see
 * that a document took seconds and not which part of it did.
 */
export interface YrsOpenProfile {
  /** Instantiating the editing core and creating a session. */
  sessionMs: number;
  /** Decoding the document state and registering its fonts. */
  loadMs: number;
  /** Lowering, measuring and paginating the whole document. */
  layoutMs: number;
  /** Building the first display list and encoding its frame. */
  frameMs: number;
  /**
   * The pass covered only the leading blocks the open asked for. The document
   * is not fully paginated until the host asks again without a restriction.
   */
  partial: boolean;
}

/** A sign of life, not a result: the request it names is still running. */
export interface ResidentEngineWorkerProgress {
  id: number;
  progress: ResidentEngineWorkerStage;
}

export type ResidentEngineWorkerResponse =
  | {
      id: number;
      ok: true;
      frame?: ArrayBuffer;
      updates?: ArrayBuffer[];
      engineMs?: number;
      workerTotalMs?: number;
      /** How long the request waited in the worker before it was handled. */
      workerQueuedMs?: number;
      /** The worker clock when the request arrived, for host-side comparison. */
      workerArrivedAt?: number;
      engineProfile?: YrsEngineApplyProfile;
      /** Present on the reply to `open` and `bootstrap`. */
      openProfile?: YrsOpenProfile;
      caret?: YrsResidentCaretSnapshot;
      selection?: YrsSelection | null;
      /** The presented frame carries the worker-painted caret line. */
      caretPainted?: boolean;
      replayMs?: number;
      replayedPages?: number;
      layoutRevision?: number;
      /** Wasm linear memory held in the worker thread, which the renderer's
       * own process metrics cannot separate out. */
      heapBytes?: number;
      /** The worker replica's yrs state vector after this operation, so the
       * next sync can ship a diff instead of the whole document state. */
      stateVector?: ArrayBuffer;
    }
  | {
      id: number;
      ok: false;
      error: string;
      residentUnavailable?: boolean;
    };

/** Everything the worker posts back: a result, or word that it is still going. */
export type ResidentEngineWorkerMessage =
  | ResidentEngineWorkerResponse
  | ResidentEngineWorkerProgress;
