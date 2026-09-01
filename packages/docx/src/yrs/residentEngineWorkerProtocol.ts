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
  | { id: number; type: 'eraseCaret' }
  | { id: number; type: 'destroy' };

export type ResidentEngineWorkerRequestWithoutId = ResidentEngineWorkerRequest extends infer Request
  ? Request extends { id: number }
    ? Omit<Request, 'id'>
    : never
  : never;

export type ResidentEngineWorkerResponse =
  | {
      id: number;
      ok: true;
      frame?: ArrayBuffer;
      updates?: ArrayBuffer[];
      engineMs?: number;
      workerTotalMs?: number;
      engineProfile?: YrsEngineApplyProfile;
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
