import type {
  YrsEngineApplyProfile,
  YrsResidentCaretSnapshot,
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
       * Opens the document in the worker, which then owns it: it seeds the
       * replica, lowers, measures, paginates and builds the display list once.
       *
       * `bootstrap` below does the same work a second time, from a main-thread
       * replica that has already done all of it — that duplicate is what makes
       * a long document cost two of everything and what the bootstrap budget
       * runs out of.
       */
      id: number;
      type: 'open';
      bytes: Uint8Array;
      fonts: Uint8Array[];
      fontsRevision: number;
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
