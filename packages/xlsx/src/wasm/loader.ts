/**
 * Wasm loader for the xlsx-wasm core.
 *
 * Call {@link initWasm} once before opening workbooks. The default browser path
 * streams the packaged wasm asset; non-fetch environments can pass bytes or a
 * precompiled module. Callers never see the JSON-string boundary.
 */

import initWasmModule, { XlsxDocument } from './generated/xlsx_wasm.js';
import type { InitInput } from './generated/xlsx_wasm.js';
import type { CollaborationReplica, CollaborationUpdateOrigin } from '../collaboration/types';
import type { DisplayList } from '../display-list/types';

/** last 0-based column index in the xlsx grid (column XFD). */
const XLSX_MAX_COL = 16_383;

/**
 * A scrolled window into a sheet. `x`/`y` are content-pixel offsets into the
 * non-frozen body. Mirrors the Rust `Viewport`.
 */
export interface Viewport {
  x: number;
  y: number;
  width: number;
  height: number;
}

/**
 * Chrome-facing sheet metadata: stable IDs, tab names, active index, and the
 * scrollable content extent of the active sheet. Mirrors the Rust `SheetInfo`.
 */
/** what `addChart` needs: chart family, grid anchor, and A1 data ranges. */
export interface AddChartArgs {
  sheet: number;
  chartType: 'column' | 'bar' | 'pie' | 'line' | 'doughnut';
  title?: string;
  /** A1 rectangle the chart floats over, e.g. "D2:K16". */
  anchor: string;
  /** category-labels range shared by every series. */
  categories?: string;
  series: Array<{ name?: string; values: string }>;
}

export interface SheetInfo {
  sheetIds: string[];
  sheetNames: string[];
  activeSheet: number;
  contentWidth: number;
  contentHeight: number;
  frozenRows: number;
  frozenCols: number;
  initialScrollX: number;
  initialScrollY: number;
}

export interface CellPosition {
  x: number;
  y: number;
}

/**
 * Result of any mutating call: whether it changed the workbook and the
 * (possibly grown) sheet metadata. Mirrors the Rust `EditResult`.
 *
 * `changed` lists the a1 addresses (on the active sheet) of *other* cells whose
 * displayed value moved as a fallout of the edit — the dependents the recalc
 * pass recomputed. It excludes the directly edited cell(s) and is omitted by
 * cores built before the recalc engine, so treat it as additive.
 */
export interface EditResult {
  applied: boolean;
  sheetInfo: SheetInfo;
  changed?: string[];
  limitedCells?: string[];
}

export interface CalculationStatus {
  limitedCells: string[];
}

/**
 * An auto filter over `range`. Each column entry keeps only the listed cell
 * texts visible (`values: null` means no value criteria for that column);
 * `showBlanks` keeps blank cells visible alongside `values`.
 */
export interface AutoFilterSpec {
  range: { startRow: number; startCol: number; endRow: number; endCol: number };
  columns: {
    col: number;
    values: string[] | null;
    showBlanks: boolean;
    /**
     * source xml of criteria the engine keeps but cannot evaluate (custom
     * comparisons, top ten, colour, date groups). Such a column constrains
     * nothing. Send it back untouched to keep it in the file; replacing a
     * column's criteria means dropping it.
     */
    unsupported?: string;
  }[];
}

/** A classic cell note: plain text plus its author. Mirrors the Rust `Comment`. */
export interface CellComment {
  author: string;
  text: string;
}

/** A cell note with its coordinates, as returned by {@link WorkbookHandle.sheetComments}. */
export interface SheetComment extends CellComment {
  row: number;
  col: number;
}

export type NumberFormat =
  | 'automatic'
  | 'plainText'
  | 'number'
  | 'percent'
  | 'scientific'
  | 'currency'
  | 'date'
  | 'time'
  | 'custom';

export type BorderPreset =
  | 'all'
  | 'inner'
  | 'horizontal'
  | 'vertical'
  | 'outer'
  | 'left'
  | 'top'
  | 'right'
  | 'bottom'
  | 'none';

export type BorderStyle = 'solid' | 'dashed' | 'dotted' | 'double';
export type HorizontalAlignment = 'left' | 'center' | 'right';
export type VerticalAlignment = 'top' | 'middle' | 'bottom';
export type TextWrapping = 'overflow' | 'wrap' | 'clip';
export type StyleProperty =
  | 'bold'
  | 'italic'
  | 'strikethrough'
  | 'fontFamily'
  | 'fontSize'
  | 'textColor'
  | 'fillColor'
  | 'borders'
  | 'horizontalAlignment'
  | 'verticalAlignment'
  | 'textWrapping';

export interface BorderPatch {
  preset?: BorderPreset;
  style?: BorderStyle;
  color?: string;
}

export interface RangeStylePatch {
  bold?: boolean;
  italic?: boolean;
  strikethrough?: boolean;
  fontFamily?: string;
  fontSize?: number;
  textColor?: string;
  fillColor?: string;
  border?: BorderPatch;
  horizontalAlignment?: HorizontalAlignment;
  verticalAlignment?: VerticalAlignment;
  textWrapping?: TextWrapping;
  clear?: StyleProperty[];
}

export type NumberFormatMutation =
  | Exclude<NumberFormat, 'custom'>
  | 'increaseDecimal'
  | 'decreaseDecimal'
  | { type: 'custom'; pattern: string };

export interface SelectionFormatting {
  numberFormat?: NumberFormat;
  numberFormatPattern?: string;
  fontFamily?: string;
  fontSize?: number;
  bold?: boolean;
  italic?: boolean;
  strikethrough?: boolean;
  textColor?: string;
  fillColor?: string;
  borderPreset?: BorderPreset;
  borderStyle?: BorderStyle;
  borderColor?: string;
  horizontalAlignment?: HorizontalAlignment;
  verticalAlignment?: VerticalAlignment;
  textWrapping?: TextWrapping;
}

export interface CapturedFormat {
  rows: number;
  columns: number;
  formats: unknown[];
}

export interface CellPoint {
  row: number;
  col: number;
}

export interface MergedRange {
  start: CellPoint;
  end: CellPoint;
}

export interface HistoryState {
  canUndo: boolean;
  canRedo: boolean;
  undoDepth: number;
  redoDepth: number;
}

export type WorkbookUpdateOrigin = CollaborationUpdateOrigin;
export type WorkbookUpdateListener = (update: Uint8Array, origin: WorkbookUpdateOrigin) => void;

export interface OpenWorkbookOptions {
  /** Open a Yrs-backed replica that can accept peer updates. */
  collaborative?: boolean;
  /** Peer-unique positive safe integer. Generated securely when omitted. */
  clientId?: number;
}

/**
 * The editable view of one cell: A1 address, the exact string the user would
 * edit (formulas as `=...`, guarded literals with a leading `'`), and whether
 * that string is a formula. Mirrors the Rust `CellEdit`.
 *
 * `filterText` is the text the engine's auto filter matches criteria against:
 * the cached bare value (for formula cells, the calculated result) with no
 * number formats. Omitted by cores built before it, so treat it as additive.
 */
export interface CellEdit {
  a1: string;
  input: string;
  isFormula: boolean;
  filterText?: string;
}

/** One cell of a batch edit: target coordinates plus the raw user input. */
export interface CellInputEdit {
  row: number;
  col: number;
  input: string;
}

/**
 * One edit inside a proposal. Same shape as {@link CellInputEdit} plus the
 * sheet index — a proposal's cells carry their own sheet on the wire, so it
 * can span sheets rather than being pinned to the active one.
 */
export interface ProposalEdit extends CellInputEdit {
  sheet: number;
  numberFormat?: NumberFormatMutation;
}

/**
 * One cell of a pending proposal: where it lands and the *formatted display
 * texts* before/after applying it. `oldText`/`newText` are what the grid shows,
 * not the raw input — the ghost decoration paints `newText` directly. Mirrors
 * the Rust proposal cell.
 */
export interface ProposalCell {
  sheet: number;
  row: number;
  col: number;
  a1: string;
  oldText: string;
  newText: string;
}

/**
 * A pending, un-applied set of cell changes an agent has staged for human
 * review. `note` is the agent's rationale (may be null); `cells` are the
 * per-cell before/after previews. Applying it is one undo step.
 */
export interface Proposal {
  id: string;
  agentId: string;
  note: string | null;
  cells: ProposalCell[];
}

/**
 * Thrown by {@link WorkbookHandle.acceptProposal} when the workbook changed
 * under a proposal since it was staged (an edit touched one of its base cells)
 * and `force` was not set. `cells` are the a1 addresses that moved, so the UI
 * can name them and offer a force-apply.
 */
export class StaleProposalError extends Error {
  readonly cells: string[];
  constructor(cells: string[]) {
    super(`stale: ${cells.join(', ')}`);
    this.name = 'StaleProposalError';
    this.cells = cells;
  }
}

// the wasm signals a stale accept with a string starting `"stale: "` followed
// by a comma-separated a1 list; parse it back into the typed error.
const STALE_PREFIX = 'stale: ';

function staleErrorFrom(message: string): StaleProposalError {
  const cells = message
    .slice(STALE_PREFIX.length)
    .split(',')
    .map((s) => s.trim())
    .filter((s) => s.length > 0);
  return new StaleProposalError(cells);
}

/**
 * A typed handle over an open workbook, hiding the wasm object and its JSON
 * boundary. Call {@link WorkbookHandle.dispose} to free the wasm memory.
 */
export interface WorkbookHandle extends CollaborationReplica {
  readonly clientId: number;
  /** Available in both modes; encodes this handle's current Yrs state vector. */
  encodeStateVector(): Uint8Array;
  /** Available in both modes; pass a peer vector to encode only the missing state. */
  encodeStateAsUpdate(remoteStateVector?: Uint8Array): Uint8Array;
  /** Apply a peer update. Standalone handles throw the Rust `NotCollaborative` error. */
  applyUpdate(update: Uint8Array): EditResult;
  /** Observe owned update bytes from local commits and accepted remote updates. */
  onUpdate(listener: WorkbookUpdateListener): () => void;
  sheetInfo(): SheetInfo;
  /** A1 extent of a sheet's populated cells; null when the sheet is empty (or the core predates the export). */
  usedRange(sheet: number): string | null;
  calculationStatus(): CalculationStatus;
  displayList(viewport: Viewport): DisplayList;
  setActiveSheet(index: number): void;
  /**
   * apply one user input to a cell (parses number/bool/formula/text) and
   * recalc its transitive dependents. one undo step; moved dependents come back
   * in `EditResult.changed`.
   */
  editCell(sheet: number, row: number, col: number, input: string): EditResult;
  /** apply a batch of inputs (paste path) as one undo step; dependents recalc. */
  editCells(sheet: number, edits: CellInputEdit[]): EditResult;
  /**
   * author a chart floating over `anchor`. Series and category ranges are A1
   * references (optionally sheet-qualified); data tracks the cells live.
   * Refused on sheets that already have drawings from the source file.
   */
  addChart(args: AddChartArgs): SheetInfo;
  /** re-anchor a chart created this session over a new A1 rectangle. */
  setChartAnchor(sheet: number, index: number, anchor: string): SheetInfo;
  /** remove a chart created this session by its index on the sheet. */
  removeChart(sheet: number, index: number): SheetInfo;
  /** raw op-list escape hatch for structural ops (insert/delete rows, merges…). */
  applyOps(ops: unknown[]): EditResult;
  /**
   * freeze the first `row` rows and `col` columns of a sheet — `row`/`col` are
   * the index of the first scrollable row/column, i.e. the count frozen.
   * `null` (or 0) leaves that axis unfrozen; both `null` clears the freeze.
   * one undo step; persists through save.
   */
  setFreezePane(sheet: number, row: number | null, col: number | null): EditResult;
  /**
   * sort rows `startRow..=endRow` (inclusive, whole rows) by the values in
   * absolute column `keyCol`; ascending=true for A→Z. one undo step.
   */
  sortRange(
    sheet: number,
    startRow: number,
    endRow: number,
    keyCol: number,
    ascending: boolean
  ): EditResult;
  /** set or replace a sheet's auto filter; `null` clears it and unhides its rows. */
  setAutoFilter(sheet: number, filter: AutoFilterSpec | null): EditResult;
  /** the sheet's current auto filter, or `null` when it has none. */
  sheetAutoFilter(sheet: number): AutoFilterSpec | null;
  /** every comment on a sheet in row-major order; empty when it has none. */
  sheetComments(sheet: number): SheetComment[];
  /**
   * set or replace the comment at a cell; `null` deletes it. one undo step;
   * persists through save.
   */
  setComment(sheet: number, row: number, col: number, comment: CellComment | null): EditResult;
  undo(): EditResult;
  redo(): EditResult;
  /** the editable view of one cell (formula bar / in-cell editor prefill). */
  cell(sheet: number, row: number, col: number): CellEdit;
  cellPosition(sheet: number, row: number, col: number): CellPosition;
  /** row-major editable views for a range, e.g. "A1:C3" (clipboard copy). */
  rangeCells(sheet: number, range: string): CellEdit[][];
  patchRangeStyle(sheet: number, range: string, patch: RangeStylePatch): EditResult;
  setNumberFormat(sheet: number, range: string, format: NumberFormatMutation): EditResult;
  selectionFormatting(sheet: number, range: string): SelectionFormatting;
  captureFormat(sheet: number, range: string): CapturedFormat;
  applyFormat(sheet: number, range: string, format: CapturedFormat): EditResult;
  mergedRanges(sheet: number, range: string): MergedRange[];
  historyState(): HistoryState;
  /**
   * render the current sheet viewport to png bytes via the native raster
   * backend — the same display list the canvas paints, rasterized server-side.
   * throws if the embedded wasm was built without png export (the `raster`
   * cargo feature); guard with {@link isPngExportAvailable}.
   */
  renderPng(viewport: Viewport): Uint8Array;
  /**
   * render an a1 range (default: the used range) at an optional scale to png.
   * dimensions are capped wasm-side; same availability guard as renderPng.
   */
  renderRangePng(opts?: { range?: string; scale?: number }): Uint8Array;
  /** serialize the workbook back to .xlsx bytes. */
  save(): Uint8Array;
  /**
   * stage an agent's proposed edits for human review without applying them.
   * returns the {@link Proposal} with per-cell formatted before/after previews.
   * throws if the embedded wasm predates proposals; guard with
   * {@link WorkbookHandle.isProposalsAvailable}.
   */
  propose(agentId: string, note: string | null, edits: ProposalEdit[]): Proposal;
  /** every pending proposal, oldest first. empty when none or unsupported. */
  listProposals(): Proposal[];
  /**
   * apply a pending proposal as one undo step and recalc dependents. throws
   * {@link StaleProposalError} when the base changed since it was staged and
   * `force` is not set; pass `{ force: true }` to apply anyway.
   */
  acceptProposal(id: string, opts?: { force?: boolean }): EditResult & { proposalId: string };
  /** drop a pending proposal without applying it; false when the id was unknown. */
  rejectProposal(id: string): boolean;
  /** whether the embedded wasm core was built with the proposals api. */
  isProposalsAvailable(): boolean;
  dispose(): void;
}

let initialized = false;
let initialization: Promise<void> | undefined;

export type WasmInitInput = InitInput | Promise<InitInput>;

/** Initialize the workbook engine. Concurrent calls share the same attempt. */
export function initWasm(
  input: WasmInitInput = new URL('./generated/xlsx_wasm_bg.wasm', import.meta.url)
): Promise<void> {
  if (initialized) return Promise.resolve();
  if (initialization) return initialization;
  initialization = initWasmModule({ module_or_path: input }).then(
    () => {
      initialized = true;
    },
    (error: unknown) => {
      initialization = undefined;
      throw toError(error);
    }
  );
  return initialization;
}

function requireInitialized(): void {
  if (!initialized) throw new Error('xlsx wasm is not initialized; call initWasm() first');
}

// wasm rejects throw strings; normalize them (and anything else) to Error.
function toError(e: unknown): Error {
  if (e instanceof Error) return e;
  return new Error(typeof e === 'string' ? e : String(e));
}

/**
 * Whether this environment exposes the WebAssembly runtime required by the core.
 */
export function isWasmAvailable(): boolean {
  return typeof WebAssembly === 'object';
}

/**
 * Whether the embedded wasm core was built with png export (the `raster` cargo
 * feature). The wasm-bindgen method only exists on the class when compiled in,
 * so chrome can disable an export control instead of calling and catching.
 */
export function isPngExportAvailable(): boolean {
  return typeof (XlsxDocument.prototype as { renderPng?: unknown }).renderPng === 'function';
}

/**
 * Whether the embedded wasm core exposes the proposals api (propose / accept /
 * reject / list). The methods only exist on the class when compiled in, so the
 * UI can hide proposal chrome and degrade gracefully against an older module.
 */
export function isProposalsAvailable(): boolean {
  return typeof (XlsxDocument.prototype as { proposeJson?: unknown }).proposeJson === 'function';
}

/**
 * Open a workbook from raw `.xlsx` bytes, initializing the core if needed.
 * Throws an `Error` if the bytes are not a readable workbook.
 */
export function openWorkbook(
  bytes: Uint8Array,
  options: OpenWorkbookOptions = {}
): WorkbookHandle {
  requireInitialized();
  const collaborativeClientId = resolveCollaborativeClientId(options);
  let doc: XlsxDocument;
  try {
    doc =
      collaborativeClientId === undefined
        ? XlsxDocument.open(bytes)
        : XlsxDocument.openCollaborative(bytes, collaborativeClientId);
  } catch (e) {
    throw toError(e);
  }

  const listeners = new Map<number, WorkbookUpdateListener>();
  const pendingUpdates: Array<{ update: Uint8Array; origin: WorkbookUpdateOrigin }> = [];
  let nextListenerId = 0;
  let disposed = false;
  let observerInstalled = false;
  let wasmCallDepth = 0;
  let flushingUpdates = false;

  function assertAlive(): void {
    if (disposed) throw new Error('workbook handle is disposed');
  }

  function flushUpdates(): void {
    if (disposed || flushingUpdates || wasmCallDepth !== 0) return;
    flushingUpdates = true;
    try {
      while (!disposed && pendingUpdates.length > 0) {
        const event = pendingUpdates.shift();
        if (!event) break;
        for (const [id, listener] of [...listeners]) {
          if (disposed) return;
          if (listeners.get(id) !== listener) continue;
          try {
            listener(event.update.slice(), event.origin);
          } catch {}
        }
      }
    } finally {
      flushingUpdates = false;
      if (disposed) pendingUpdates.length = 0;
    }
  }

  function drainWasmUpdates(): void {
    if (!observerInstalled || disposed) return;
    while (true) {
      const encoded = doc.drainUpdateEvent();
      if (encoded.byteLength === 0) return;
      const origin = encoded[0];
      if (origin !== 0 && origin !== 1) {
        throw new Error(`xlsx wasm returned unknown update origin ${origin}`);
      }
      pendingUpdates.push({
        update: encoded.slice(1),
        origin: origin === 0 ? 'local' : 'remote',
      });
    }
  }

  function wasmCall<T>(operation: () => T, drainUpdates = false): T {
    assertAlive();
    wasmCallDepth += 1;
    try {
      let result: T | undefined;
      let failure: unknown;
      let failed = false;
      try {
        result = operation();
      } catch (error) {
        failure = error;
        failed = true;
      }
      if (drainUpdates) {
        try {
          drainWasmUpdates();
        } catch (error) {
          if (!failed) {
            failure = error;
            failed = true;
          }
        }
      }
      if (failed) throw toError(failure);
      return result as T;
    } finally {
      wasmCallDepth -= 1;
      if (wasmCallDepth === 0) flushUpdates();
    }
  }

  function mutatingWasmCall<T>(operation: () => T): T {
    return wasmCall(operation, true);
  }

  function parseJson<T>(operation: () => string, drainUpdates = false): T {
    return wasmCall(() => JSON.parse(operation()) as T, drainUpdates);
  }

  function ensureUpdateObserver(): void {
    if (observerInstalled) return;
    wasmCall(() => doc.startUpdateObservation());
    observerInstalled = true;
  }

  function clearUnusedUpdateObserver(): void {
    if (!observerInstalled || listeners.size > 0 || disposed) return;
    pendingUpdates.length = 0;
    wasmCall(() => doc.clearUpdateObservation());
    observerInstalled = false;
  }

  const handle: WorkbookHandle = {
    get clientId(): number {
      return wasmCall(() => doc.clientId);
    },
    encodeStateVector(): Uint8Array {
      return wasmCall(() => doc.encodeStateVector().slice());
    },
    encodeStateAsUpdate(remoteStateVector?: Uint8Array): Uint8Array {
      return wasmCall(() =>
        remoteStateVector === undefined
          ? doc.encodeStateAsUpdate().slice()
          : doc.encodeDiff(remoteStateVector.slice()).slice()
      );
    },
    applyUpdate(update: Uint8Array): EditResult {
      return parseJson(() => doc.applyUpdateJson(update.slice()), true);
    },
    onUpdate(listener: WorkbookUpdateListener): () => void {
      assertAlive();
      if (typeof listener !== 'function') throw new TypeError('update listener must be a function');
      const id = nextListenerId++;
      listeners.set(id, listener);
      try {
        ensureUpdateObserver();
      } catch (error) {
        listeners.delete(id);
        throw error;
      }
      let subscribed = true;
      return () => {
        if (!subscribed) return;
        subscribed = false;
        listeners.delete(id);
        clearUnusedUpdateObserver();
      };
    },
    sheetInfo(): SheetInfo {
      return parseJson(() => doc.sheetInfoJson());
    },
    usedRange(sheet: number): string | null {
      return wasmCall(() => {
        const fn = (doc as { usedRangeJson?: (sheet: number) => string }).usedRangeJson;
        if (typeof fn !== 'function') return null;
        return (JSON.parse(fn.call(doc, sheet)) as { usedRange: string | null }).usedRange;
      });
    },
    calculationStatus(): CalculationStatus {
      return wasmCall(() => {
        const fn = (doc as { calculationStatusJson?: () => string }).calculationStatusJson;
        if (typeof fn !== 'function') return { limitedCells: [] };
        return JSON.parse(fn.call(doc)) as CalculationStatus;
      });
    },
    displayList(viewport: Viewport): DisplayList {
      return parseJson(() => doc.displayListJson(JSON.stringify(viewport)));
    },
    setActiveSheet(index: number): void {
      mutatingWasmCall(() => doc.setActiveSheet(index));
    },
    editCell(sheet: number, row: number, col: number, input: string): EditResult {
      return parseJson(() => doc.editCellJson(JSON.stringify({ sheet, row, col, input })), true);
    },
    addChart(args: AddChartArgs): SheetInfo {
      return parseJson(() => doc.addChartJson(JSON.stringify(args)), true);
    },
    setChartAnchor(sheet: number, index: number, anchor: string): SheetInfo {
      return parseJson(
        () => doc.setChartAnchorJson(JSON.stringify({ sheet, index, anchor })),
        true
      );
    },
    removeChart(sheet: number, index: number): SheetInfo {
      return parseJson(() => doc.removeChartJson(JSON.stringify({ sheet, index })), true);
    },
    editCells(sheet: number, edits: CellInputEdit[]): EditResult {
      return parseJson(() => doc.editCellsJson(JSON.stringify({ sheet, edits })), true);
    },
    applyOps(ops: unknown[]): EditResult {
      return parseJson(() => doc.applyOpsJson(JSON.stringify({ ops })), true);
    },
    setFreezePane(sheet: number, row: number | null, col: number | null): EditResult {
      const rows = row ?? 0;
      const cols = col ?? 0;
      const pane =
        rows > 0 || cols > 0
          ? { rows, cols, top_left: { row: rows, col: cols } }
          : null;
      const ops = [{ type: 'setFreezePane', sheet, pane }];
      return parseJson(() => doc.applyOpsJson(JSON.stringify({ ops })), true);
    },
    sortRange(
      sheet: number,
      startRow: number,
      endRow: number,
      keyCol: number,
      ascending: boolean
    ): EditResult {
      const ops = [
        {
          type: 'sortRange',
          sheet,
          range: {
            start: { row: startRow, col: 0 },
            end: { row: endRow, col: XLSX_MAX_COL },
          },
          key_col: keyCol,
          ascending,
        },
      ];
      return parseJson(() => doc.applyOpsJson(JSON.stringify({ ops })), true);
    },
    setAutoFilter(sheet: number, filter: AutoFilterSpec | null): EditResult {
      const wire = filter
        ? {
            range: {
              start: { row: filter.range.startRow, col: filter.range.startCol },
              end: { row: filter.range.endRow, col: filter.range.endCol },
            },
            columns: filter.columns.map((c) => ({
              col: c.col,
              values: c.unsupported ? null : c.values,
              show_blanks: c.showBlanks,
              ...(c.unsupported ? { unsupported: c.unsupported } : {}),
            })),
          }
        : null;
      const ops = [{ type: 'setAutoFilter', sheet, filter: wire }];
      return parseJson(() => doc.applyOpsJson(JSON.stringify({ ops })), true);
    },
    sheetAutoFilter(sheet: number): AutoFilterSpec | null {
      return parseJson(() => doc.sheetAutoFilterJson(JSON.stringify({ sheet })));
    },
    sheetComments(sheet: number): SheetComment[] {
      const parsed = parseJson<{ comments: SheetComment[] }>(() =>
        doc.sheetCommentsJson(JSON.stringify({ sheet }))
      );
      return parsed.comments;
    },
    setComment(
      sheet: number,
      row: number,
      col: number,
      comment: CellComment | null
    ): EditResult {
      const ops = [{ type: 'setComment', sheet, cell: { row, col }, comment }];
      return parseJson(() => doc.applyOpsJson(JSON.stringify({ ops })), true);
    },
    undo(): EditResult {
      return parseJson(() => doc.undoJson(), true);
    },
    redo(): EditResult {
      return parseJson(() => doc.redoJson(), true);
    },
    cell(sheet: number, row: number, col: number): CellEdit {
      return parseJson(() => doc.cellJson(JSON.stringify({ sheet, row, col })));
    },
    cellPosition(sheet: number, row: number, col: number): CellPosition {
      return parseJson(() => doc.cellPositionJson(JSON.stringify({ sheet, row, col })));
    },
    rangeCells(sheet: number, range: string): CellEdit[][] {
      const parsed = parseJson<{ cells: CellEdit[][] }>(() =>
        doc.rangeCellsJson(JSON.stringify({ sheet, range }))
      );
      return parsed.cells;
    },
    patchRangeStyle(sheet: number, range: string, patch: RangeStylePatch): EditResult {
      return parseJson(
        () => doc.patchRangeStyleJson(JSON.stringify({ sheet, range, patch })),
        true
      );
    },
    setNumberFormat(
      sheet: number,
      range: string,
      format: NumberFormatMutation
    ): EditResult {
      const wire = typeof format === 'string' ? { type: format } : format;
      return parseJson(
        () => doc.setRangeNumberFormatJson(JSON.stringify({ sheet, range, format: wire })),
        true
      );
    },
    selectionFormatting(sheet: number, range: string): SelectionFormatting {
      return parseJson(() => doc.selectionFormattingJson(JSON.stringify({ sheet, range })));
    },
    captureFormat(sheet: number, range: string): CapturedFormat {
      return parseJson(() => doc.captureFormatJson(JSON.stringify({ sheet, range })));
    },
    applyFormat(sheet: number, range: string, format: CapturedFormat): EditResult {
      return parseJson(
        () => doc.applyFormatJson(JSON.stringify({ sheet, range, format })),
        true
      );
    },
    mergedRanges(sheet: number, range: string): MergedRange[] {
      const parsed = parseJson<{ ranges: MergedRange[] }>(() =>
        doc.mergedRangesJson(JSON.stringify({ sheet, range }))
      );
      return parsed.ranges;
    },
    historyState(): HistoryState {
      return parseJson(() => doc.historyStateJson());
    },
    renderPng(viewport: Viewport): Uint8Array {
      return wasmCall(() => {
        const fn = (doc as { renderPng?: (v: string) => Uint8Array }).renderPng;
        if (typeof fn !== 'function') throw new Error('png export not in this build');
        return fn.call(doc, JSON.stringify(viewport)).slice();
      });
    },
    renderRangePng(opts?: { range?: string; scale?: number }): Uint8Array {
      return wasmCall(() => {
        const fn = (doc as { renderRangePng?: (a: string) => Uint8Array }).renderRangePng;
        if (typeof fn !== 'function') throw new Error('png export not in this build');
        return fn.call(doc, JSON.stringify(opts ?? {})).slice();
      });
    },
    save(): Uint8Array {
      return wasmCall(() => doc.saveBytes().slice());
    },
    propose(agentId: string, note: string | null, edits: ProposalEdit[]): Proposal {
      return mutatingWasmCall(() => {
        const fn = (doc as { proposeJson?: (a: string) => string }).proposeJson;
        if (typeof fn !== 'function') throw new Error('proposals not in this build');
        const wireEdits = edits.map(({ numberFormat, ...edit }) => ({
          ...edit,
          ...(numberFormat === undefined
            ? {}
            : { numberFormat: typeof numberFormat === 'string' ? { type: numberFormat } : numberFormat }),
        }));
        return JSON.parse(
          fn.call(doc, JSON.stringify({ agentId, note, edits: wireEdits }))
        ) as Proposal;
      });
    },
    listProposals(): Proposal[] {
      return wasmCall(() => {
        const fn = (doc as { listProposalsJson?: () => string }).listProposalsJson;
        if (typeof fn !== 'function') return [];
        return (JSON.parse(fn.call(doc)) as { proposals: Proposal[] }).proposals;
      });
    },
    acceptProposal(id: string, opts?: { force?: boolean }): EditResult & { proposalId: string } {
      try {
        return mutatingWasmCall(() => {
          const fn = (doc as { acceptProposalJson?: (a: string) => string }).acceptProposalJson;
          if (typeof fn !== 'function') throw new Error('proposals not in this build');
          return JSON.parse(
            fn.call(doc, JSON.stringify({ id, force: opts?.force ?? false }))
          ) as EditResult & { proposalId: string };
        });
      } catch (e) {
        const message = e instanceof Error ? e.message : typeof e === 'string' ? e : String(e);
        if (message.startsWith(STALE_PREFIX)) throw staleErrorFrom(message);
        throw toError(e);
      }
    },
    rejectProposal(id: string): boolean {
      return mutatingWasmCall(() => {
        const fn = (doc as { rejectProposalJson?: (a: string) => string }).rejectProposalJson;
        if (typeof fn !== 'function') return false;
        return (JSON.parse(fn.call(doc, JSON.stringify({ id }))) as { removed: boolean }).removed;
      });
    },
    isProposalsAvailable(): boolean {
      return wasmCall(() => typeof (doc as { proposeJson?: unknown }).proposeJson === 'function');
    },
    dispose(): void {
      if (disposed) return;
      disposed = true;
      listeners.clear();
      pendingUpdates.length = 0;
      let disposalError: unknown;
      if (observerInstalled) {
        try {
          doc.clearUpdateObservation();
        } catch (error) {
          disposalError = error;
        }
        observerInstalled = false;
      }
      try {
        doc.free();
      } catch (error) {
        disposalError ??= error;
      }
      if (disposalError !== undefined) throw toError(disposalError);
    },
  };
  return handle;
}

function resolveCollaborativeClientId(options: OpenWorkbookOptions): number | undefined {
  if (!options.collaborative) {
    if (options.clientId !== undefined) {
      throw new TypeError('clientId requires collaborative mode');
    }
    return undefined;
  }
  if (options.clientId !== undefined) {
    if (
      !Number.isSafeInteger(options.clientId) ||
      options.clientId <= 0 ||
      options.clientId > Number.MAX_SAFE_INTEGER
    ) {
      throw new RangeError('clientId must be a nonzero safe integer');
    }
    return options.clientId;
  }

  const random = globalThis.crypto;
  if (!random || typeof random.getRandomValues !== 'function') {
    throw new Error('crypto.getRandomValues is required to generate a collaboration client ID');
  }
  const words = new Uint32Array(2);
  let value: number;
  do {
    random.getRandomValues(words);
    value = (words[0] & 0x1fffff) * 0x1_0000_0000 + words[1];
  } while (value === 0);
  return value;
}

/**
 * The crate version string, for asserting wasm/js parity.
 */
export function wasmVersion(): string {
  requireInitialized();
  return XlsxDocument.version();
}
