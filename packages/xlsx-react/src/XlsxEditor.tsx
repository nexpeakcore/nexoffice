/**
 * `<XlsxEditor />` — the editor shell. A dpr-aware canvas paints the grid; DOM
 * overlays (selection marquee, active-cell outline, in-cell editor) sit above it,
 * positioned from the same display-list geometry the painter uses. A top bar
 * holds the name box, formula bar, and save/undo/redo; an offscreen `role=grid`
 * mirror serves screen readers. All compute lives in `@betteroffice/xlsx`; this
 * layer is framework glue — keyboard/mouse wiring, focus flow, and DOM chrome.
 */

import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import {
  buildA11yGrid,
  cellAtPoint,
  cellRect,
  extendTo,
  fromTsv,
  hyperlinkAtCell,
  initWasm,
  isPngExportAvailable,
  isProposalsAvailable,
  moveFocus,
  normalizeRange,
  openWorkbook,
  paintDisplayList,
  parseHyperlinkLocation,
  rangeRect,
  selectionAt,
  selectionKeyReducer,
  safeExternalHyperlink,
  StaleProposalError,
  toTsv,
} from '@betteroffice/xlsx';
import type {
  CellAddr,
  CellComment,
  CellEdit,
  CellInputEdit,
  CapturedFormat,
  DisplayList,
  DrawCmd,
  Direction,
  EditResult,
  MergedRange,
  Proposal,
  Selection,
  SelectionLimits,
  SheetComment,
  SheetInfo,
  WorkbookHandle,
} from '@betteroffice/xlsx';
import type {
  AwarenessPeer,
  CollaborationAwareness,
  CollaborationReplica,
} from '@betteroffice/xlsx/collaboration';
import type { Translations } from '@betteroffice/xlsx-i18n';
import { LocaleProvider, useTranslation } from './i18n';
import { EditorToolbar } from './components/EditorToolbar';
import type {
  FormattingAction,
  MergeAction,
  SelectionFormatting,
} from './components/Toolbar';
import {
  ToolbarButton,
  ToolbarDropdown,
  ToolbarGroup,
  ToolbarMenuItem,
  ToolbarMenuSeparator,
} from './components/ui/ToolbarPrimitives';
import { ToolbarIcon } from './components/ui/ToolbarIcon';
import { FilterPopover } from './components/filters/FilterPopover';
import {
  collectFilterValues,
  columnCriteria,
  columnLabel,
  emptyFilterSpec,
  expandDataRegion,
  filterColumnIndices,
  hasCriteria,
  MAX_FILTER_VALUES,
  withColumnCriteria,
} from './components/filters/filterSpec';
import type { FilterSpec } from './components/filters/filterSpec';
import {
  CommentEditorPopover,
  CommentViewPopover,
} from './components/comments/CommentPopover';
import { commentAt, resolveCommentAuthor } from './components/comments/commentState';
import {
  expandRangeToMergedCells,
  PresenceStrip,
  RemoteSelections,
} from './presence/Presence';
import { ProposalsPanel } from './proposals/ProposalsPanel';

/**
 * The imperative surface handed to {@link XlsxEditorProps.onReady}: the open
 * workbook handle plus a `refreshProposals` to re-read the pending list after an
 * external caller (e.g. a demo agent) stages proposals on the same handle, and
 * the editor's own history, clipboard, selection, and zoom actions so a host
 * shell (menus, command palettes) can drive them against the live selection.
 */
export interface XlsxEditorApi {
  handle: WorkbookHandle;
  refreshProposals: () => void;
  /** Undo the last edit through the editor's history pipeline. */
  undo: () => void;
  /** Redo the last undone edit through the editor's history pipeline. */
  redo: () => void;
  /** Copy the current selection to the system clipboard as TSV. */
  copySelection: () => Promise<void>;
  /** Copy the current selection to the clipboard, then clear it. */
  cutSelection: () => Promise<void>;
  /** Paste clipboard TSV into the grid at the current selection. */
  pasteSelection: () => Promise<void>;
  /** Clear the contents of the current selection (Delete key equivalent). */
  clearSelection: () => void;
  /** Select the sheet's whole used range (Cmd/Ctrl+A equivalent). */
  selectAll: () => void;
  /** The current zoom factor (1 = 100%). */
  getZoom: () => number;
  /** Set the zoom factor, clamped to the toolbar's 25%–400% range. */
  setZoom: (zoom: number) => void;
}

export interface XlsxEditorCollaborationOptions {
  /** Peer-unique Yrs client ID. Generated securely when omitted. */
  clientId?: number;
  /** Shared Yrs state applied before the replica is exposed. */
  initialUpdate?: Uint8Array;
  /** Receive the editor-owned collaboration replica. */
  onReplica?: (replica: CollaborationReplica | null) => void;
  /** Presence-capable provider connected to the editor-owned replica. */
  provider?: CollaborationAwareness | null;
}

/**
 * Props for {@link XlsxEditor}.
 */
export interface XlsxEditorProps {
  /** Raw .xlsx bytes to open. When omitted the shell paints a demo frame. */
  file?: Uint8Array;
  /** Download name for the save button; falls back to `workbook.xlsx`. */
  fileName?: string;
  /** Receive saved bytes instead of triggering a browser download. */
  onSave?: (bytes: Uint8Array) => void;
  /** Called after every applied local edit — dirty tracking without polling. */
  onEdit?: () => void;
  /** Author name stamped on new cell comments; falls back to `User`. */
  userName?: string;
  /** Open a network-ready Yrs replica and repaint when peer updates arrive. */
  collaboration?: XlsxEditorCollaborationOptions;
  i18n?: Translations;
  /**
   * Called when a workbook opens, with a handle to stage agent proposals and a
   * way to refresh the panel afterward. Enables demo/host agents without
   * exposing the wasm object through the render tree. A returned cleanup runs
   * before the workbook is replaced or disposed.
   */
  onReady?: (api: XlsxEditorApi) => void | (() => void);
  className?: string;
}

/** the open in-cell editor: which cell it targets and its current draft text. */
interface EditState {
  row: number;
  col: number;
  value: string;
}

/** the open filter popover: its column and the distinct values read for it. */
interface OpenFilterState {
  col: number;
  values: string[];
  hasBlanks: boolean;
  truncated: boolean;
}

/** the open comment popover: its cell and whether it views or edits. */
interface OpenCommentState {
  row: number;
  col: number;
  mode: 'view' | 'edit';
}

function cellA1(row: number, col: number): string {
  return `${columnLabel(col)}${row + 1}`;
}

const XLSX_MAX_ROW = 1_048_575;
const XLSX_MAX_COL = 16_383;
const MAX_FILTER_REGION_ROWS = 10_000;
const MAX_FILTER_REGION_COLS = 256;

const COL_W = 96;
const ROW_H = 24;
const MIN_ZOOM = 0.25;
const MAX_ZOOM = 4;
const BRAND = '#217346';
const DEFAULT_XLSX_TOOLBAR_HEIGHT = 87;
const MAX_OVERLAY_MERGED_RANGES = 1024;
const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

// a placeholder grid frame so the shell paints something real when no file is
// open. real files render through the wasm display list instead.
function buildDemoDisplayList(width: number, height: number, cellText: string): DisplayList {
  const commands: DrawCmd[] = [
    { op: 'fillRect', x: 0, y: 0, w: width, h: height, color: '#ffffff' },
  ];
  const cols = Math.ceil(width / COL_W);
  const rows = Math.ceil(height / ROW_H);
  for (let c = 0; c <= cols; c++) {
    commands.push({
      op: 'line',
      x1: c * COL_W,
      y1: 0,
      x2: c * COL_W,
      y2: height,
      width: 1,
      color: '#e0e0e0',
    });
  }
  for (let r = 0; r <= rows; r++) {
    commands.push({
      op: 'line',
      x1: 0,
      y1: r * ROW_H,
      x2: width,
      y2: r * ROW_H,
      width: 1,
      color: '#e0e0e0',
    });
  }
  commands.push({ op: 'fillRect', x: 0, y: 0, w: width, h: ROW_H, color: '#f3f3f3' });
  commands.push({ op: 'fillRect', x: 0, y: 0, w: COL_W, h: height, color: '#f3f3f3' });
  commands.push({
    op: 'text',
    x: COL_W + 8,
    y: ROW_H + 18,
    text: cellText,
    fontSize: 14,
    color: '#202020',
    clip: { x: COL_W, y: ROW_H, w: COL_W * 3, h: ROW_H },
    align: 'left',
  });
  return { width, height, commands };
}

// median of the gaps between consecutive offsets, or a fallback when the window
// has no tracks. the median ignores outliers like a single very wide column, so
// the extent-to-count estimate below is not skewed by one atypical track.
function medianTrack(offsets: number[] | undefined, fallback: number): number {
  if (!offsets || offsets.length < 2) return fallback;
  const gaps: number[] = [];
  for (let i = 1; i < offsets.length; i++) gaps.push(offsets[i] - offsets[i - 1]);
  gaps.sort((a, b) => a - b);
  const mid = gaps[gaps.length >> 1];
  return mid > 0 ? mid : fallback;
}

// derive nav bounds from the scrollable extent: rows/cols estimated from the
// content size over a representative (median) track size, rowsPerPage from the
// viewport. a slack of one keeps the row/col just past the used edge reachable.
function deriveLimits(
  dl: DisplayList | null,
  info: SheetInfo,
  viewportHeight: number
): SelectionLimits {
  const rowH = medianTrack(dl?.grid?.rowOffsets, ROW_H);
  const colW = medianTrack(dl?.grid?.colOffsets, COL_W);
  const rows = Math.max(1, Math.round(info.contentHeight / rowH)) + 1;
  const cols = Math.max(1, Math.round(info.contentWidth / colW)) + 1;
  const rowsPerPage = Math.max(1, Math.floor(viewportHeight / rowH));
  return { rows, cols, rowsPerPage };
}

// trigger a browser download of a byte blob under the given name and mime type.
function downloadBytes(bytes: Uint8Array, name: string, mime: string): void {
  const blob = new Blob([new Uint8Array(bytes)], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = name;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

// the png download name derived from the workbook name: swap .xlsx for .png.
function pngName(fileName: string | undefined): string {
  return `${(fileName ?? 'workbook.xlsx').replace(/\.xlsx$/i, '')}.png`;
}

function scaledRect(rect: { x: number; y: number; w: number; h: number }, zoom: number) {
  return {
    x: rect.x * zoom,
    y: rect.y * zoom,
    w: rect.w * zoom,
    h: rect.h * zoom,
  };
}

const visuallyHidden: React.CSSProperties = {
  position: 'absolute',
  width: 1,
  height: 1,
  margin: -1,
  padding: 0,
  border: 0,
  overflow: 'hidden',
  clip: 'rect(0 0 0 0)',
  whiteSpace: 'nowrap',
};

const xlsxToolbarStyles: Record<string, React.CSSProperties> = {
  shell: {
    flex: '0 0 auto',
    minHeight: DEFAULT_XLSX_TOOLBAR_HEIGHT,
    padding: '4px 0 5px',
    borderBottom: '1px solid #e2e8f0',
    background: '#ffffff',
    color: '#0f172a',
    boxSizing: 'border-box',
  },
  rail: {
    display: 'flex',
    alignItems: 'center',
    minHeight: 32,
    margin: '0 8px',
    padding: '2px 8px',
    borderRadius: 4,
    background: '#ffffff',
    boxSizing: 'border-box',
    overflowX: 'auto',
    overflowY: 'hidden',
  },
  group: {
    display: 'inline-flex',
    alignItems: 'center',
    gap: 1,
    padding: '0 6px',
    borderRight: '1px solid rgba(226, 232, 240, 0.9)',
    flex: '0 0 auto',
  },
  formulaGroup: {
    display: 'flex',
    alignItems: 'center',
    gap: 4,
    flex: '1 1 320px',
    minWidth: 240,
    padding: '0 6px',
    borderRight: '1px solid rgba(226, 232, 240, 0.9)',
  },
  nameBox: {
    appearance: 'none',
    width: 64,
    height: 28,
    flex: '0 0 auto',
    boxSizing: 'border-box',
    border: '1px solid #e2e8f0',
    borderRadius: 6,
    background: '#f8fafc',
    color: '#0f172a',
    font: '600 12px ui-monospace, SFMono-Regular, Menlo, monospace',
    textAlign: 'center',
    outlineColor: '#2563eb',
  },
  formulaMark: {
    display: 'grid',
    placeItems: 'center',
    width: 20,
    height: 28,
    flex: '0 0 auto',
    color: '#64748b',
    font: 'italic 700 12px Georgia, serif',
    userSelect: 'none',
  },
  formulaInput: {
    appearance: 'none',
    flex: '1 1 260px',
    minWidth: 140,
    height: 28,
    boxSizing: 'border-box',
    border: '1px solid #e2e8f0',
    borderRadius: 6,
    padding: '0 8px',
    background: '#ffffff',
    color: '#0f172a',
    font: '13px ui-sans-serif, system-ui, sans-serif',
    outlineColor: '#2563eb',
  },
  proposals: {
    marginLeft: 'auto',
    paddingLeft: 6,
    flex: '0 0 auto',
  },
  count: {
    display: 'inline-grid',
    placeItems: 'center',
    minWidth: 15,
    height: 15,
    marginLeft: -3,
    padding: '0 4px',
    borderRadius: 8,
    background: '#0f172a',
    color: '#ffffff',
    fontSize: 10,
    fontWeight: 700,
    lineHeight: 1,
    boxSizing: 'border-box',
  },
};

/**
 * The xlsx editor React component.
 */
export function XlsxEditor({
  i18n,
  ...props
}: XlsxEditorProps) {
  return (
    <LocaleProvider i18n={i18n}>
      <XlsxEditorContent {...props} />
    </LocaleProvider>
  );
}

function XlsxEditorContent({
  file,
  fileName,
  onSave,
  onEdit,
  userName,
  collaboration,
  onReady,
  className,
}: Omit<XlsxEditorProps, 'i18n'>) {
  const { t } = useTranslation();
  const collaborationEnabled = collaboration !== undefined;
  const collaborationClientId = collaboration?.clientId;
  const collaborationInitialUpdate = collaboration?.initialUpdate;
  const collaborationOnReplica = collaboration?.onReplica;
  const collaborationProvider = collaboration?.provider;
  const toolbarRef = useRef<HTMLDivElement>(null);
  const scrollRef = useRef<HTMLDivElement>(null);
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const handleRef = useRef<WorkbookHandle | null>(null);
  const frameRef = useRef<DisplayList | null>(null);
  const rafRef = useRef<number | null>(null);
  const editorInputRef = useRef<HTMLInputElement>(null);
  const draggingRef = useRef(false);
  const clickStartRef = useRef<CellAddr | null>(null);
  const suppressBlurRef = useRef(false);
  const pendingSheetViewRef = useRef(false);
  // latest onReady, read (not depended on) by the open effect so a changing
  // callback identity never reopens the workbook.
  const onReadyRef = useRef(onReady);
  onReadyRef.current = onReady;
  // latest onEdit, read by applyResult so a changing callback identity never
  // invalidates every mutation callback.
  const onEditRef = useRef(onEdit);
  onEditRef.current = onEdit;
  // latest history/clipboard callbacks, read by the api handed to onReady so
  // its stable methods always act on the current selection.
  const editorActionsRef = useRef<{
    undo: () => void;
    redo: () => void;
    copySelection: () => Promise<void>;
    cutSelection: () => Promise<void>;
    pasteSelection: () => Promise<void>;
    clearSelection: () => void;
    selectAll: () => void;
    getZoom: () => number;
    setZoom: (zoom: number) => void;
  } | null>(null);

  const [sheetInfo, setSheetInfo] = useState<SheetInfo | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [frame, setFrame] = useState<DisplayList | null>(null);
  const [selection, setSelection] = useState<Selection | null>(null);
  const [editing, setEditing] = useState<EditState | null>(null);
  const [focusedCell, setFocusedCell] = useState<CellEdit | null>(null);
  const [formulaDraft, setFormulaDraft] = useState<string | null>(null);
  const [toolbarHeight, setToolbarHeight] = useState(DEFAULT_XLSX_TOOLBAR_HEIGHT);
  const [zoom, setZoom] = useState(1);
  const [revision, setRevision] = useState(0);
  const [dragging, setDragging] = useState(false);
  const [selectionFormatting, setSelectionFormatting] = useState<SelectionFormatting>({});
  const [historyState, setHistoryState] = useState({ canUndo: false, canRedo: false });
  const [mergedRanges, setMergedRanges] = useState<
    Array<{ start: CellAddr; end: CellAddr }>
  >([]);
  const [visibleMergedRanges, setVisibleMergedRanges] = useState<readonly MergedRange[]>([]);
  const [borderStyleChoice, setBorderStyleChoice] = useState<SelectionFormatting['borderStyle']>();
  const [borderColorChoice, setBorderColorChoice] = useState<string>();
  const [capturedFormat, setCapturedFormat] = useState<CapturedFormat | null>(null);
  // auto filters keyed by sheet index, seeded from the engine on open and
  // re-read after every mutation, so persisted filters keep their chrome.
  const [filterSpecs, setFilterSpecs] = useState<Record<number, FilterSpec>>({});
  const [openFilter, setOpenFilter] = useState<OpenFilterState | null>(null);
  // the active sheet's comments, re-read from the engine after every mutation.
  const [sheetComments, setSheetComments] = useState<readonly SheetComment[]>([]);
  const [openComment, setOpenComment] = useState<OpenCommentState | null>(null);
  const paintSourceRef = useRef<string | null>(null);
  const [proposals, setProposals] = useState<Proposal[]>([]);
  const [proposalsPanelOpen, setProposalsPanelOpen] = useState(false);
  const [collaborationReplica, setCollaborationReplica] =
    useState<CollaborationReplica | null>(null);
  const [awarenessPeers, setAwarenessPeers] = useState<readonly AwarenessPeer[]>([]);
  // a1 lists keyed by proposal id: cells that drifted since a proposal was
  // staged, surfaced when accepting it throws a StaleProposalError.
  const [staleFor, setStaleFor] = useState<Record<string, string[]>>({});

  const activeSheet = sheetInfo?.activeSheet ?? 0;

  // whether the embedded core was built with png export (raster cargo feature).
  // stable for the module's lifetime, so the export control can pre-disable.
  const pngExportAvailable = useMemo(() => isPngExportAvailable(), []);

  // whether the embedded core exposes the proposals api; gates all proposal
  // chrome so the editor degrades cleanly against an older module.
  const proposalsAvailable = useMemo(() => isProposalsAvailable(), []);

  useEffect(() => {
    const toolbar = toolbarRef.current;
    if (!toolbar) return;
    const updateHeight = () => setToolbarHeight(toolbar.offsetHeight);
    updateHeight();
    const observer = new ResizeObserver(updateHeight);
    observer.observe(toolbar);
    return () => observer.disconnect();
  }, []);

  // re-read the pending proposal list and queue a repaint — ghosts paint into
  // the engine frame, so every lifecycle change (propose/accept/reject) must
  // republish it. safe to call against an old core (the loader returns an
  // empty list).
  const refreshProposals = useCallback(() => {
    const handle = handleRef.current;
    if (!handle) {
      setProposals([]);
      return;
    }
    try {
      setProposals(handle.listProposals());
    } catch {
      setProposals([]);
    }
    setRevision((r) => r + 1);
  }, []);

  // re-read a sheet's persisted auto filter and comments from the engine so
  // header dropdowns and comment indicators track undo/redo, structural ops,
  // and remote updates instead of drifting on in-session state.
  const refreshSheetState = useCallback((sheet: number) => {
    const handle = handleRef.current;
    if (!handle) return;
    try {
      const spec = handle.sheetAutoFilter(sheet);
      setFilterSpecs((specs) => {
        if (spec) return { ...specs, [sheet]: spec };
        if (!(sheet in specs)) return specs;
        const { [sheet]: _removed, ...rest } = specs;
        return rest;
      });
      setSheetComments(handle.sheetComments(sheet));
    } catch {}
  }, []);

  // open the workbook when the file changes; dispose it on change/unmount and
  // reset all editing state so a dropped file starts clean.
  useEffect(() => {
    setEditing(null);
    setFormulaDraft(null);
    setProposals([]);
    setStaleFor({});
    setProposalsPanelOpen(false);
    setCollaborationReplica(null);
    setSelectionFormatting({});
    setHistoryState({ canUndo: false, canRedo: false });
    setMergedRanges([]);
    setVisibleMergedRanges([]);
    setBorderStyleChoice(undefined);
    setBorderColorChoice(undefined);
    setCapturedFormat(null);
    setFilterSpecs({});
    setOpenFilter(null);
    setSheetComments([]);
    setOpenComment(null);
    paintSourceRef.current = null;
    pendingSheetViewRef.current = false;
    if (!file) {
      handleRef.current = null;
      setSheetInfo(null);
      setSelection(null);
      setError(null);
      return;
    }
    handleRef.current = null;
    setSheetInfo(null);
    setSelection(null);
    let handle: WorkbookHandle | null = null;
    let unsubscribeUpdates = () => {};
    let cleanupReady = () => {};
    let disposed = false;
    const runReadyCleanup = () => {
      const cleanup = cleanupReady;
      cleanupReady = () => {};
      try {
        cleanup();
      } catch {}
    };
    void initWasm().then(
      () => {
        if (disposed) return;
        try {
          handle = openWorkbook(file, {
            collaborative: collaborationEnabled,
            clientId: collaborationClientId,
          });
          if (collaborationInitialUpdate) {
            handle.applyUpdate(collaborationInitialUpdate.slice());
          }
          handleRef.current = handle;
          unsubscribeUpdates = handle.onUpdate(() => {
            if (disposed || !handle) return;
            try {
              const info = handle.sheetInfo();
              setSheetInfo(info);
              refreshSheetState(info.activeSheet);
              setRevision((current) => current + 1);
              setStaleFor({});
              refreshProposals();
              setError(null);
            } catch (e) {
              setError(e instanceof Error ? e.message : String(e));
            }
          });
          pendingSheetViewRef.current = true;
          const info = handle.sheetInfo();
          setSheetInfo(info);
          const seededSpecs: Record<number, FilterSpec> = {};
          info.sheetIds.forEach((_, sheet) => {
            const spec = handle?.sheetAutoFilter(sheet);
            if (spec) seededSpecs[sheet] = spec;
          });
          setFilterSpecs(seededSpecs);
          setSheetComments(handle.sheetComments(info.activeSheet));
          setSelection(selectionAt({ row: 0, col: 0 }));
          setCollaborationReplica(handle);
          setError(null);
          refreshProposals();
          const cleanup = onReadyRef.current?.({
            handle,
            refreshProposals,
            undo: () => editorActionsRef.current?.undo(),
            redo: () => editorActionsRef.current?.redo(),
            copySelection: () =>
              editorActionsRef.current?.copySelection() ?? Promise.resolve(),
            cutSelection: () =>
              editorActionsRef.current?.cutSelection() ?? Promise.resolve(),
            pasteSelection: () =>
              editorActionsRef.current?.pasteSelection() ?? Promise.resolve(),
            clearSelection: () => editorActionsRef.current?.clearSelection(),
            selectAll: () => editorActionsRef.current?.selectAll(),
            getZoom: () => editorActionsRef.current?.getZoom() ?? 1,
            setZoom: (zoom: number) => editorActionsRef.current?.setZoom(zoom),
          });
          if (typeof cleanup === 'function') cleanupReady = cleanup;
        } catch (e) {
          runReadyCleanup();
          unsubscribeUpdates();
          unsubscribeUpdates = () => {};
          handle?.dispose();
          handle = null;
          handleRef.current = null;
          setSheetInfo(null);
          setSelection(null);
          setError(e instanceof Error ? e.message : String(e));
        }
      },
      (e: unknown) => {
        if (disposed) return;
        handleRef.current = null;
        setSheetInfo(null);
        setSelection(null);
        setError(e instanceof Error ? e.message : String(e));
      }
    );
    return () => {
      disposed = true;
      runReadyCleanup();
      unsubscribeUpdates();
      handle?.dispose();
      handleRef.current = null;
    };
  }, [
    file,
    collaborationEnabled,
    collaborationClientId,
    collaborationInitialUpdate,
    refreshProposals,
    refreshSheetState,
  ]);

  useEffect(() => {
    if (!sheetInfo || !pendingSheetViewRef.current) return;
    const scroll = scrollRef.current;
    if (!scroll) return;
    pendingSheetViewRef.current = false;
    scroll.scrollLeft = sheetInfo.initialScrollX * zoom;
    scroll.scrollTop = sheetInfo.initialScrollY * zoom;
  }, [sheetInfo, zoom]);

  useEffect(() => {
    if (!collaborationOnReplica || !collaborationReplica) return;
    collaborationOnReplica(collaborationReplica);
    return () => collaborationOnReplica(null);
  }, [collaborationOnReplica, collaborationReplica]);

  useEffect(() => {
    setAwarenessPeers([]);
    if (!collaborationProvider) return;
    return collaborationProvider.onAwareness((peers) => setAwarenessPeers([...peers]));
  }, [collaborationProvider]);

  useEffect(() => {
    return () => collaborationProvider?.setCursor(null);
  }, [collaborationProvider]);

  useEffect(() => {
    if (!collaborationProvider) return;
    const sheet = sheetInfo?.sheetIds[activeSheet];
    if (!selection || !sheet) {
      collaborationProvider.setCursor(null);
      return;
    }
    collaborationProvider.setCursor({
      sheet,
      anchor: { ...selection.anchor },
      head: { ...selection.focus },
    });
  }, [collaborationProvider, selection, sheetInfo, activeSheet, revision]);

  // paint the current scroll window into the canvas and publish the frame for
  // overlays + a11y. reads refs so it stays identity-stable across renders.
  const doPaint = useCallback(() => {
    const scroll = scrollRef.current;
    const canvas = canvasRef.current;
    if (!scroll || !canvas) return;
    const ctx = canvas.getContext('2d');
    if (!ctx) return;
    const dpr = window.devicePixelRatio || 1;
    const w = scroll.clientWidth;
    const h = scroll.clientHeight;
    if (w === 0 || h === 0) return;
    canvas.width = Math.round(w * dpr);
    canvas.height = Math.round(h * dpr);
    canvas.style.width = `${w}px`;
    canvas.style.height = `${h}px`;
    const logicalWidth = w / zoom;
    const logicalHeight = h / zoom;
    const handle = handleRef.current;
    let dl: DisplayList;
    if (handle) {
      try {
        dl = handle.displayList({
          x: scroll.scrollLeft / zoom,
          y: scroll.scrollTop / zoom,
          width: logicalWidth,
          height: logicalHeight,
        });
      } catch {
        return;
      }
    } else {
      dl = buildDemoDisplayList(logicalWidth, logicalHeight, t('editor.demoCellText'));
    }
    let nextMergedRanges: readonly MergedRange[] = [];
    const grid = dl.grid;
    const rows = (grid?.rowOffsets.length ?? 0) - 1;
    const columns = (grid?.colOffsets.length ?? 0) - 1;
    if (handle && grid && rows > 0 && columns > 0) {
      try {
        const from = handle.cell(activeSheet, grid.startRow, grid.startCol).a1;
        const to = handle.cell(
          activeSheet,
          grid.startRow + rows - 1,
          grid.startCol + columns - 1
        ).a1;
        nextMergedRanges = handle
          .mergedRanges(activeSheet, `${from}:${to}`)
          .slice(0, MAX_OVERLAY_MERGED_RANGES);
      } catch {}
    }
    paintDisplayList(ctx, dl, dpr * zoom);
    frameRef.current = dl;
    setVisibleMergedRanges(nextMergedRanges);
    setFrame(dl);
  }, [activeSheet, t, zoom]);

  // paint loop: repaint on scroll/resize (rAF-coalesced) and whenever the open
  // workbook, active sheet, or a mutation (revision) changes the pixels.
  useEffect(() => {
    const scroll = scrollRef.current;
    if (!scroll) return;
    const schedulePaint = () => {
      if (rafRef.current != null) return;
      rafRef.current = requestAnimationFrame(() => {
        rafRef.current = null;
        doPaint();
      });
    };
    doPaint();
    scroll.addEventListener('scroll', schedulePaint, { passive: true });
    const observer = new ResizeObserver(schedulePaint);
    observer.observe(scroll);
    return () => {
      scroll.removeEventListener('scroll', schedulePaint);
      observer.disconnect();
      if (rafRef.current != null) {
        cancelAnimationFrame(rafRef.current);
        rafRef.current = null;
      }
    };
  }, [doPaint, sheetInfo, error, revision]);

  // read the focused cell's editable text for the name box + formula bar; reruns
  // as the selection moves or the workbook mutates.
  useEffect(() => {
    const handle = handleRef.current;
    if (!handle || !selection || !sheetInfo) {
      setFocusedCell(null);
      return;
    }
    try {
      setFocusedCell(handle.cell(activeSheet, selection.focus.row, selection.focus.col));
    } catch {
      setFocusedCell(null);
    }
  }, [selection, sheetInfo, activeSheet, revision]);

  // clear a stuck drag if the mouse is released outside the grid.
  useEffect(() => {
    const stop = () => {
      draggingRef.current = false;
      setDragging(false);
    };
    window.addEventListener('mouseup', stop);
    return () => window.removeEventListener('mouseup', stop);
  }, []);

  useEffect(() => {
    const handle = handleRef.current;
    if (!handle || !selection || !sheetInfo) {
      setSelectionFormatting({});
      setHistoryState({ canUndo: false, canRedo: false });
      setMergedRanges([]);
      return;
    }
    const range = normalizeRange(selection);
    try {
      const from = handle.cell(activeSheet, range.top, range.left).a1;
      const to = handle.cell(activeSheet, range.bottom, range.right).a1;
      const a1 = `${from}:${to}`;
      setSelectionFormatting(handle.selectionFormatting(activeSheet, a1));
      setHistoryState(handle.historyState());
      setMergedRanges(handle.mergedRanges(activeSheet, a1));
    } catch {
      setSelectionFormatting({});
      setHistoryState({ canUndo: false, canRedo: false });
      setMergedRanges([]);
    }
  }, [selection, sheetInfo, activeSheet, revision]);

  // rebuilt from the live frame so the offscreen mirror never lags a mutation;
  // the visible window is small, so a rebuild per paint frame is cheap enough.
  const a11yGrid = useMemo(() => {
    if (!frame || !sheetInfo) return null;
    return buildA11yGrid(frame, selection, sheetInfo.sheetNames[activeSheet] ?? '', {
      gridLabel: t('a11y.gridLabel'),
      rowHeaderLabel: t('a11y.rowHeaderLabel'),
      columnHeaderLabel: t('a11y.columnHeaderLabel'),
      cellLabel: t('a11y.cellLabel'),
      cellLabelSelected: t('a11y.cellLabelSelected'),
      emptyCellLabel: t('a11y.emptyCellLabel'),
      emptyCellLabelSelected: t('a11y.emptyCellLabelSelected'),
    });
  }, [frame, selection, sheetInfo, activeSheet, t]);

  // preventScroll everywhere: the sticky overlay host sits below the full-height
  // canvas in flow, so a plain focus() scrolls the grid to bring it into view.
  const focusContainer = useCallback(() => {
    scrollRef.current?.focus({ preventScroll: true });
  }, []);

  // focus the in-cell editor when it opens, without scrolling the grid.
  useEffect(() => {
    if (editing) editorInputRef.current?.focus({ preventScroll: true });
  }, [editing]);

  // fold a mutation result back into state and queue a repaint. re-reads the
  // pending proposals because structural ops and undo/redo can drop them, and
  // the active sheet's filter + comments because any mutation can move them.
  const applyResult = useCallback(
    (result: EditResult) => {
      setSheetInfo(result.sheetInfo);
      refreshSheetState(result.sheetInfo.activeSheet);
      setRevision((r) => r + 1);
      refreshProposals();
      if (result.applied) onEditRef.current?.();
    },
    [refreshProposals, refreshSheetState]
  );

  const selectedRangeA1 = useCallback(
    (target: Selection): string | null => {
      const handle = handleRef.current;
      if (!handle) return null;
      const range = normalizeRange(target);
      const from = handle.cell(activeSheet, range.top, range.left).a1;
      const to = handle.cell(activeSheet, range.bottom, range.right).a1;
      return `${from}:${to}`;
    },
    [activeSheet]
  );

  const limits = useCallback((): SelectionLimits => {
    return deriveLimits(
      frameRef.current,
      sheetInfo!,
      (scrollRef.current?.clientHeight ?? 0) / zoom
    );
  }, [sheetInfo, zoom]);

  // map a viewport-local pointer event to a sheet cell via the frame geometry.
  const pointToCell = useCallback(
    (clientX: number, clientY: number): CellAddr | null => {
      const canvas = canvasRef.current;
      const grid = frameRef.current?.grid;
      if (!canvas || !grid) return null;
      const rect = canvas.getBoundingClientRect();
      return cellAtPoint(grid, (clientX - rect.left) / zoom, (clientY - rect.top) / zoom);
    },
    [zoom]
  );

  const openEditor = useCallback(
    (seed?: string) => {
      const handle = handleRef.current;
      if (!handle || !selection) return;
      const { row, col } = selection.focus;
      let value = seed ?? '';
      if (seed === undefined) {
        try {
          value = handle.cell(activeSheet, row, col).input;
        } catch {
          value = '';
        }
      }
      setEditing({ row, col, value });
    },
    [selection, activeSheet]
  );

  // commit the open editor, optionally stepping the selection like excel.
  const commitEditor = useCallback(
    (move?: Direction) => {
      const handle = handleRef.current;
      if (!handle || !editing) return;
      suppressBlurRef.current = true;
      const { row, col, value } = editing;
      try {
        applyResult(handle.editCell(activeSheet, row, col, value));
      } catch (e) {
        setError(e instanceof Error ? e.message : String(e));
      }
      setEditing(null);
      const base = selectionAt({ row, col });
      setSelection(move ? moveFocus(base, move, { limits: limits() }) : base);
      focusContainer();
    },
    [editing, activeSheet, applyResult, limits, focusContainer]
  );

  const cancelEditor = useCallback(() => {
    suppressBlurRef.current = true;
    setEditing(null);
    focusContainer();
  }, [focusContainer]);

  const clearCells = useCallback(() => {
    const handle = handleRef.current;
    if (!handle || !selection) return;
    const r = normalizeRange(selection);
    const edits: CellInputEdit[] = [];
    for (let row = r.top; row <= r.bottom; row++) {
      for (let col = r.left; col <= r.right; col++) edits.push({ row, col, input: '' });
    }
    try {
      applyResult(handle.editCells(activeSheet, edits));
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    }
  }, [selection, activeSheet, applyResult]);

  const selectAllCells = useCallback(() => {
    if (!sheetInfo) return;
    const bounds = limits();
    setSelection({
      anchor: { row: 0, col: 0 },
      focus: { row: bounds.rows - 1, col: bounds.cols - 1 },
    });
  }, [sheetInfo, limits]);

  const copySelection = useCallback(async () => {
    const handle = handleRef.current;
    if (!handle || !selection) return;
    const r = normalizeRange(selection);
    try {
      const from = handle.cell(activeSheet, r.top, r.left).a1;
      const to = handle.cell(activeSheet, r.bottom, r.right).a1;
      const cells = handle.rangeCells(activeSheet, `${from}:${to}`);
      const tsv = toTsv(
        cells.map((row) => row.map((c) => ({ input: c.input, isFormula: c.isFormula })))
      );
      await navigator.clipboard.writeText(tsv);
    } catch {
      // clipboard denied or read failed — nothing to paste, leave state as-is.
    }
  }, [selection, activeSheet]);

  const cutSelection = useCallback(async () => {
    await copySelection();
    clearCells();
  }, [copySelection, clearCells]);

  const pasteSelection = useCallback(async () => {
    const handle = handleRef.current;
    if (!handle || !selection) return;
    let text: string;
    try {
      text = await navigator.clipboard.readText();
    } catch {
      return;
    }
    const grid = fromTsv(text);
    if (grid.length === 0) return;
    const r = normalizeRange(selection);
    const edits: CellInputEdit[] = [];
    let width = 1;
    grid.forEach((rowArr, dr) => {
      width = Math.max(width, rowArr.length);
      rowArr.forEach((input, dc) => edits.push({ row: r.top + dr, col: r.left + dc, input }));
    });
    try {
      applyResult(handle.editCells(activeSheet, edits));
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
      return;
    }
    setSelection({
      anchor: { row: r.top, col: r.left },
      focus: { row: r.top + grid.length - 1, col: r.left + width - 1 },
    });
  }, [selection, activeSheet, applyResult]);

  const formatSelection = useCallback(
    (action: FormattingAction) => {
      const handle = handleRef.current;
      if (!handle || !selection) return;
      const range = selectedRangeA1(selection);
      if (!range) return;
      if (action === 'paintFormat') {
        if (capturedFormat) {
          setCapturedFormat(null);
          paintSourceRef.current = null;
          return;
        }
        try {
          setCapturedFormat(handle.captureFormat(activeSheet, range));
          const normalized = normalizeRange(selection);
          paintSourceRef.current = `${activeSheet}:${normalized.top}:${normalized.left}:${normalized.bottom}:${normalized.right}`;
        } catch (e) {
          setError(e instanceof Error ? e.message : String(e));
        }
        return;
      }
      try {
        let result: EditResult;
        if (action === 'currency') {
          result = handle.setNumberFormat(activeSheet, range, 'currency');
        } else if (action === 'percent') {
          result = handle.setNumberFormat(activeSheet, range, 'percent');
        } else if (action === 'increaseDecimal') {
          result = handle.setNumberFormat(activeSheet, range, 'increaseDecimal');
        } else if (action === 'decreaseDecimal') {
          result = handle.setNumberFormat(activeSheet, range, 'decreaseDecimal');
        } else if (action === 'bold') {
          result = handle.patchRangeStyle(activeSheet, range, {
            bold: !selectionFormatting.bold,
          });
        } else if (action === 'italic') {
          result = handle.patchRangeStyle(activeSheet, range, {
            italic: !selectionFormatting.italic,
          });
        } else if (action === 'strikethrough') {
          result = handle.patchRangeStyle(activeSheet, range, {
            strikethrough: !selectionFormatting.strikethrough,
          });
        } else if (action.type === 'numberFormat') {
          result =
            action.value === 'custom'
              ? handle.setNumberFormat(activeSheet, range, {
                  type: 'custom',
                  pattern: selectionFormatting.numberFormatPattern ?? '0.00',
                })
              : handle.setNumberFormat(activeSheet, range, action.value);
        } else if (action.type === 'fontFamily') {
          result = handle.patchRangeStyle(activeSheet, range, { fontFamily: action.value });
        } else if (action.type === 'fontSize') {
          result = handle.patchRangeStyle(activeSheet, range, { fontSize: action.value });
        } else if (action.type === 'textColor') {
          result = handle.patchRangeStyle(activeSheet, range, { textColor: action.value });
        } else if (action.type === 'fillColor') {
          result = handle.patchRangeStyle(activeSheet, range, { fillColor: action.value });
        } else if (action.type === 'borderPreset') {
          result = handle.patchRangeStyle(activeSheet, range, {
            border: {
              preset: action.value,
              style: borderStyleChoice ?? selectionFormatting.borderStyle ?? 'solid',
              color: borderColorChoice ?? selectionFormatting.borderColor ?? '#000000',
            },
          });
        } else if (action.type === 'borderStyle') {
          setBorderStyleChoice(action.value);
          result = handle.patchRangeStyle(activeSheet, range, {
            border: { style: action.value },
          });
        } else if (action.type === 'borderColor') {
          setBorderColorChoice(action.value);
          result = handle.patchRangeStyle(activeSheet, range, {
            border: { color: action.value },
          });
        } else if (action.type === 'horizontalAlignment') {
          result = handle.patchRangeStyle(activeSheet, range, {
            horizontalAlignment: action.value,
          });
        } else if (action.type === 'verticalAlignment') {
          result = handle.patchRangeStyle(activeSheet, range, {
            verticalAlignment: action.value,
          });
        } else {
          result = handle.patchRangeStyle(activeSheet, range, {
            textWrapping: action.value,
          });
        }
        applyResult(result);
        focusContainer();
      } catch (e) {
        setError(e instanceof Error ? e.message : String(e));
      }
    },
    [
      selection,
      selectedRangeA1,
      capturedFormat,
      activeSheet,
      selectionFormatting,
      borderStyleChoice,
      borderColorChoice,
      applyResult,
      focusContainer,
    ]
  );

  useEffect(() => {
    const handle = handleRef.current;
    if (!handle || !selection || !capturedFormat || dragging) return;
    const normalized = normalizeRange(selection);
    const key = `${activeSheet}:${normalized.top}:${normalized.left}:${normalized.bottom}:${normalized.right}`;
    if (key === paintSourceRef.current) return;
    const range = selectedRangeA1(selection);
    if (!range) return;
    try {
      applyResult(handle.applyFormat(activeSheet, range, capturedFormat));
      setCapturedFormat(null);
      paintSourceRef.current = null;
      focusContainer();
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    }
  }, [
    selection,
    capturedFormat,
    dragging,
    activeSheet,
    selectedRangeA1,
    applyResult,
    focusContainer,
  ]);

  const undo = useCallback(() => {
    const handle = handleRef.current;
    if (!handle) return;
    try {
      applyResult(handle.undo());
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    }
  }, [applyResult]);

  const redo = useCallback(() => {
    const handle = handleRef.current;
    if (!handle) return;
    try {
      applyResult(handle.redo());
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    }
  }, [applyResult]);

  editorActionsRef.current = {
    undo,
    redo,
    copySelection,
    cutSelection,
    pasteSelection,
    clearSelection: clearCells,
    selectAll: selectAllCells,
    getZoom: () => zoom,
    setZoom: (next: number) => setZoom(Math.min(MAX_ZOOM, Math.max(MIN_ZOOM, next))),
  };

  const print = useCallback(() => window.print(), []);

  const sortSelection = useCallback(
    (ascending: boolean) => {
      const handle = handleRef.current;
      if (!handle || !selection) return;
      const minRow = Math.min(selection.anchor.row, selection.focus.row);
      const maxRow = Math.max(selection.anchor.row, selection.focus.row);
      const keyCol = selection.anchor.col;
      try {
        applyResult(handle.sortRange(activeSheet, minRow, maxRow, keyCol, ascending));
      } catch (e) {
        setError(e instanceof Error ? e.message : String(e));
      }
    },
    [applyResult, selection, activeSheet]
  );

  const searchMenus = useCallback(() => undefined, []);

  const freezePanes = useCallback(
    (rows: number | null, cols: number | null) => {
      const handle = handleRef.current;
      if (!handle) return;
      try {
        applyResult(handle.setFreezePane(activeSheet, rows, cols));
        focusContainer();
      } catch (e) {
        setError(e instanceof Error ? e.message : String(e));
      }
    },
    [activeSheet, applyResult, focusContainer]
  );

  // the contiguous non-empty block around the selection (with its header row),
  // probed one boundary row/column per step through the wasm cell reader.
  const dataRegionForSelection = useCallback((): {
    top: number;
    left: number;
    bottom: number;
    right: number;
  } | null => {
    const handle = handleRef.current;
    if (!handle || !selection) return null;
    const normalized = normalizeRange(selection);
    const anyContent = (fromRow: number, fromCol: number, toRow: number, toCol: number) =>
      handle
        .rangeCells(activeSheet, `${cellA1(fromRow, fromCol)}:${cellA1(toRow, toCol)}`)
        .some((row) => row.some((cell) => cell.input !== ''));
    return expandDataRegion(
      normalized,
      {
        rowHasContent: (row, left, right) => anyContent(row, left, row, right),
        colHasContent: (col, top, bottom) => anyContent(top, col, bottom, col),
      },
      {
        maxRow: Math.min(XLSX_MAX_ROW, normalized.top + MAX_FILTER_REGION_ROWS - 1),
        maxCol: Math.min(XLSX_MAX_COL, normalized.left + MAX_FILTER_REGION_COLS - 1),
      }
    );
  }, [selection, activeSheet]);

  const toggleFilter = useCallback(() => {
    const handle = handleRef.current;
    if (!handle || !selection) return;
    try {
      if (filterSpecs[activeSheet]) {
        applyResult(handle.setAutoFilter(activeSheet, null));
      } else {
        const region = dataRegionForSelection();
        if (!region) return;
        const spec = emptyFilterSpec({
          startRow: region.top,
          startCol: region.left,
          endRow: region.bottom,
          endCol: region.right,
        });
        applyResult(handle.setAutoFilter(activeSheet, spec));
      }
      setOpenFilter(null);
      focusContainer();
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    }
  }, [
    selection,
    activeSheet,
    filterSpecs,
    dataRegionForSelection,
    applyResult,
    focusContainer,
  ]);

  // open (or toggle closed) a column's popover, reading the distinct raw cell
  // texts of its body rows — the same texts the engine matches criteria against.
  const openFilterFor = useCallback(
    (col: number) => {
      const handle = handleRef.current;
      const spec = filterSpecs[activeSheet];
      if (!handle || !spec) return;
      if (openFilter?.col === col) {
        setOpenFilter(null);
        return;
      }
      // the applied criteria values are unioned into the list so a selection
      // beyond the collection cap stays visible and survives the next Apply.
      const selected = columnCriteria(spec, col).values;
      const bodyStart = spec.range.startRow + 1;
      if (bodyStart > spec.range.endRow) {
        setOpenFilter({ col, ...collectFilterValues([], MAX_FILTER_VALUES, selected) });
        return;
      }
      try {
        const cells = handle
          .rangeCells(
            activeSheet,
            `${cellA1(bodyStart, col)}:${cellA1(spec.range.endRow, col)}`
          )
          .map((row) => row[0]);
        setOpenFilter({ col, ...collectFilterValues(cells, MAX_FILTER_VALUES, selected) });
      } catch (e) {
        setError(e instanceof Error ? e.message : String(e));
      }
    },
    [filterSpecs, activeSheet, openFilter]
  );

  const applyColumnFilter = useCallback(
    (col: number, values: string[] | null, showBlanks: boolean, clearPreserved = false) => {
      const handle = handleRef.current;
      const spec = filterSpecs[activeSheet];
      if (!handle || !spec) return;
      const next = withColumnCriteria(spec, col, values, showBlanks, clearPreserved);
      try {
        applyResult(handle.setAutoFilter(activeSheet, next));
        setOpenFilter(null);
        focusContainer();
      } catch (e) {
        setError(e instanceof Error ? e.message : String(e));
      }
    },
    [filterSpecs, activeSheet, applyResult, focusContainer]
  );

  // toggle the comment editor popover for the anchored cell.
  const toggleCommentEditor = useCallback(() => {
    if (!selection) return;
    const { row, col } = selection.focus;
    setOpenComment((open) =>
      open?.mode === 'edit' && open.row === row && open.col === col
        ? null
        : { row, col, mode: 'edit' }
    );
  }, [selection]);

  // set, replace, or (null) delete a cell's comment as one undo step.
  const setCellComment = useCallback(
    (row: number, col: number, comment: CellComment | null) => {
      const handle = handleRef.current;
      if (!handle) return;
      try {
        applyResult(handle.setComment(activeSheet, row, col, comment));
        setOpenComment(null);
        focusContainer();
      } catch (e) {
        setError(e instanceof Error ? e.message : String(e));
      }
    },
    [activeSheet, applyResult, focusContainer]
  );

  const mergeSelection = useCallback(
    (action: MergeAction) => {
      const handle = handleRef.current;
      if (!handle || !selection) return;
      const range = normalizeRange(selection);
      const cell = (row: number, col: number) => ({ row, col });
      const mergeRange = (top: number, left: number, bottom: number, right: number) => ({
        start: cell(top, left),
        end: cell(bottom, right),
      });
      const ops: unknown[] = [];
      if (action === 'all') {
        ops.push({
          type: 'mergeCells',
          sheet: activeSheet,
          range: mergeRange(range.top, range.left, range.bottom, range.right),
        });
      } else if (action === 'horizontal') {
        for (let row = range.top; row <= range.bottom; row++) {
          ops.push({
            type: 'mergeCells',
            sheet: activeSheet,
            range: mergeRange(row, range.left, row, range.right),
          });
        }
      } else if (action === 'vertical') {
        for (let col = range.left; col <= range.right; col++) {
          ops.push({
            type: 'mergeCells',
            sheet: activeSheet,
            range: mergeRange(range.top, col, range.bottom, col),
          });
        }
      } else {
        for (const merged of mergedRanges) {
          ops.push({
            type: 'unmergeCells',
            sheet: activeSheet,
            range: mergeRange(
              merged.start.row,
              merged.start.col,
              merged.end.row,
              merged.end.col
            ),
          });
        }
      }
      if (ops.length === 0) return;
      try {
        applyResult(handle.applyOps(ops));
        focusContainer();
      } catch (e) {
        setError(e instanceof Error ? e.message : String(e));
      }
    },
    [selection, activeSheet, mergedRanges, applyResult, focusContainer]
  );

  // accept a proposal (optionally forcing past drift): apply it, drop any stale
  // warning, refresh the list, repaint, and return focus to the grid.
  const acceptProposal = useCallback(
    (id: string, force?: boolean) => {
      const handle = handleRef.current;
      if (!handle) return;
      try {
        applyResult(handle.acceptProposal(id, { force }));
        setStaleFor(({ [id]: _dropped, ...rest }) => rest);
        refreshProposals();
        focusContainer();
      } catch (e) {
        if (e instanceof StaleProposalError) setStaleFor((m) => ({ ...m, [id]: e.cells }));
        else setError(e instanceof Error ? e.message : String(e));
      }
    },
    [applyResult, refreshProposals, focusContainer]
  );

  // reject a proposal: drop it and its warning, then refresh so its border
  // chrome disappears and the canvas repaints without its ghost.
  const rejectProposal = useCallback(
    (id: string) => {
      const handle = handleRef.current;
      if (!handle) return;
      try {
        handle.rejectProposal(id);
        setStaleFor(({ [id]: _dropped, ...rest }) => rest);
        refreshProposals();
      } catch (e) {
        setError(e instanceof Error ? e.message : String(e));
      }
    },
    [refreshProposals]
  );

  const save = useCallback(() => {
    const handle = handleRef.current;
    if (!handle) return;
    try {
      const bytes = handle.save();
      if (onSave) onSave(bytes);
      else downloadBytes(bytes, fileName ?? 'workbook.xlsx', XLSX_MIME);
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    }
  }, [onSave, fileName]);

  // render the current scroll window to png via the raster backend and download
  // it — the same display list the canvas paints, rasterized in the core.
  const exportPng = useCallback(() => {
    const handle = handleRef.current;
    const scroll = scrollRef.current;
    if (!handle || !scroll) return;
    try {
      const png = handle.renderPng({
        x: scroll.scrollLeft / zoom,
        y: scroll.scrollTop / zoom,
        width: scroll.clientWidth / zoom,
        height: scroll.clientHeight / zoom,
      });
      downloadBytes(png, pngName(fileName), 'image/png');
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    }
  }, [fileName, zoom]);

  // commit the formula bar draft to the focused cell.
  const commitFormula = useCallback(
    (move?: Direction) => {
      const handle = handleRef.current;
      if (!handle || !selection || formulaDraft == null) return;
      const { row, col } = selection.focus;
      try {
        applyResult(handle.editCell(activeSheet, row, col, formulaDraft));
      } catch (e) {
        setError(e instanceof Error ? e.message : String(e));
      }
      setFormulaDraft(null);
      if (move) setSelection((prev) => (prev ? moveFocus(prev, move, { limits: limits() }) : prev));
    },
    [selection, formulaDraft, activeSheet, applyResult, limits]
  );

  // grid-level keyboard: chrome shortcuts first, then the pure selection reducer.
  const onKeyDown = useCallback(
    (e: React.KeyboardEvent) => {
      const handle = handleRef.current;
      if (!handle || !selection || !sheetInfo || editing) return;
      const mod = e.metaKey || e.ctrlKey;
      const lower = e.key.toLowerCase();

      if (mod) {
        if (lower === 'c') {
          void copySelection();
          e.preventDefault();
          return;
        }
        if (lower === 'v') {
          void pasteSelection();
          e.preventDefault();
          return;
        }
        if (lower === 'x') {
          void cutSelection();
          e.preventDefault();
          return;
        }
        if (lower === 'z') {
          e.shiftKey ? redo() : undo();
          e.preventDefault();
          return;
        }
        if (lower === 'y') {
          redo();
          e.preventDefault();
          return;
        }
        if (lower === 's') {
          save();
          e.preventDefault();
          return;
        }
      }

      const action = selectionKeyReducer(
        selection,
        {
          key: e.key,
          shiftKey: e.shiftKey,
          metaKey: e.metaKey,
          ctrlKey: e.ctrlKey,
          altKey: e.altKey,
        },
        limits()
      );
      switch (action.type) {
        case 'move':
          setSelection(action.selection);
          e.preventDefault();
          break;
        case 'startEdit':
          openEditor(action.initialInput);
          e.preventDefault();
          break;
        case 'clear':
          clearCells();
          e.preventDefault();
          break;
        case 'none':
          break;
      }
    },
    [
      selection,
      sheetInfo,
      editing,
      limits,
      copySelection,
      pasteSelection,
      cutSelection,
      undo,
      redo,
      save,
      openEditor,
      clearCells,
    ]
  );

  const onMouseDown = useCallback(
    (e: React.MouseEvent) => {
      if (!selection || editing) return;
      const addr = pointToCell(e.clientX, e.clientY);
      if (!addr) return;
      if (e.shiftKey) setSelection((prev) => (prev ? extendTo(prev, addr, limits()) : prev));
      else setSelection(selectionAt(addr));
      clickStartRef.current = addr;
      draggingRef.current = true;
      setDragging(true);
      focusContainer();
    },
    [editing, selection, pointToCell, limits, focusContainer]
  );

  const onMouseMove = useCallback(
    (e: React.MouseEvent<HTMLDivElement>) => {
      const addr = pointToCell(e.clientX, e.clientY);
      if (!addr) {
        if (!draggingRef.current) {
          e.currentTarget.style.cursor = 'default';
          e.currentTarget.title = '';
        }
        return;
      }
      if (!draggingRef.current) {
        const hyperlink = frameRef.current
          ? hyperlinkAtCell(frameRef.current, addr.row, addr.col)
          : null;
        const comment = commentAt(sheetComments, addr.row, addr.col);
        e.currentTarget.style.cursor = hyperlink ? 'pointer' : 'default';
        e.currentTarget.title =
          hyperlink?.tooltip ?? (comment ? `${comment.author}: ${comment.text}` : '');
        return;
      }
      setSelection((prev) => (prev ? extendTo(prev, addr, limits()) : prev));
    },
    [pointToCell, limits, sheetComments]
  );

  const activateHyperlink = useCallback(
    (addr: CellAddr): boolean => {
      const handle = handleRef.current;
      const currentFrame = frameRef.current;
      if (!handle || !currentFrame || !sheetInfo) return false;
      const hyperlink = hyperlinkAtCell(currentFrame, addr.row, addr.col);
      if (!hyperlink) return false;
      const href = safeExternalHyperlink(hyperlink);
      if (href) {
        window.open(href, '_blank', 'noopener,noreferrer');
        return true;
      }
      if (!hyperlink.location) return true;
      const currentName = sheetInfo.sheetNames[activeSheet] ?? '';
      const destination = parseHyperlinkLocation(hyperlink.location, currentName);
      if (!destination) return true;
      const targetSheet = sheetInfo.sheetNames.findIndex(
        (name) => name.toLowerCase() === destination.sheetName.toLowerCase()
      );
      if (targetSheet < 0) return true;
      try {
        handle.setActiveSheet(targetSheet);
        const position = handle.cellPosition(targetSheet, destination.row, destination.col);
        setSheetInfo(handle.sheetInfo());
        requestAnimationFrame(() => {
          const scroll = scrollRef.current;
          if (!scroll) return;
          scroll.scrollLeft = position.x * zoom;
          scroll.scrollTop = position.y * zoom;
        });
        setSelection(selectionAt({ row: destination.row, col: destination.col }));
        setEditing(null);
        setFormulaDraft(null);
        setCapturedFormat(null);
        paintSourceRef.current = null;
        setError(null);
      } catch (error) {
        setError(error instanceof Error ? error.message : String(error));
      }
      return true;
    },
    [activeSheet, sheetInfo, zoom]
  );

  const onClick = useCallback(
    (e: React.MouseEvent) => {
      if (editing) {
        clickStartRef.current = null;
        return;
      }
      const addr = pointToCell(e.clientX, e.clientY);
      const start = clickStartRef.current;
      clickStartRef.current = null;
      if (!addr || !start || addr.row !== start.row || addr.col !== start.col) return;
      if (activateHyperlink(addr)) return;
      const comment = commentAt(sheetComments, addr.row, addr.col);
      setOpenComment(comment ? { row: addr.row, col: addr.col, mode: 'view' } : null);
    },
    [activateHyperlink, editing, pointToCell, sheetComments]
  );

  const onDoubleClick = useCallback(
    (e: React.MouseEvent) => {
      if (editing) return;
      const addr = pointToCell(e.clientX, e.clientY);
      if (
        addr &&
        frameRef.current &&
        hyperlinkAtCell(frameRef.current, addr.row, addr.col)
      ) {
        return;
      }
      if (selection) openEditor();
    },
    [editing, openEditor, pointToCell, selection]
  );

  const onMouseLeave = useCallback((e: React.MouseEvent<HTMLDivElement>) => {
    e.currentTarget.style.cursor = 'default';
    e.currentTarget.title = '';
  }, []);

  const grid = frame?.grid;
  const renderedSelection = selection
    ? expandRangeToMergedCells(normalizeRange(selection), visibleMergedRanges)
    : null;
  const renderedFocus = selection
    ? expandRangeToMergedCells(
        {
          top: selection.focus.row,
          left: selection.focus.col,
          bottom: selection.focus.row,
          right: selection.focus.col,
        },
        visibleMergedRanges
      )
    : null;
  const selRect = grid && renderedSelection ? rangeRect(grid, renderedSelection) : null;
  const focusRect = grid && renderedFocus ? rangeRect(grid, renderedFocus) : null;
  const editRect = grid && editing ? cellRect(grid, editing.row, editing.col) : null;
  const scaledSelectionRect = selRect ? scaledRect(selRect, zoom) : null;
  const scaledFocusRect = focusRect ? scaledRect(focusRect, zoom) : null;
  const scaledEditRect = editRect ? scaledRect(editRect, zoom) : null;

  const anchorRow = selection?.anchor.row ?? 0;
  const anchorCol = selection?.anchor.col ?? 0;
  const frozenRows = sheetInfo?.frozenRows ?? 0;
  const frozenCols = sheetInfo?.frozenCols ?? 0;

  const activeFilterSpec = sheetInfo ? filterSpecs[activeSheet] ?? null : null;
  const filterButtonRects =
    activeFilterSpec && grid
      ? filterColumnIndices(activeFilterSpec.range).flatMap((col) => {
          const rect = cellRect(grid, activeFilterSpec.range.startRow, col);
          return rect ? [{ col, rect: scaledRect(rect, zoom) }] : [];
        })
      : [];
  const openFilterHeaderRect =
    activeFilterSpec && grid && openFilter
      ? cellRect(grid, activeFilterSpec.range.startRow, openFilter.col)
      : null;
  const scaledOpenFilterRect = openFilterHeaderRect
    ? scaledRect(openFilterHeaderRect, zoom)
    : null;
  const openFilterCriteria =
    activeFilterSpec && openFilter ? columnCriteria(activeFilterSpec, openFilter.col) : null;

  const commentIndicators = grid
    ? sheetComments.flatMap((comment) => {
        const rect = cellRect(grid, comment.row, comment.col);
        return rect ? [{ comment, rect: scaledRect(rect, zoom) }] : [];
      })
    : [];
  const openCommentRect =
    grid && openComment ? cellRect(grid, openComment.row, openComment.col) : null;
  const scaledOpenCommentRect = openCommentRect ? scaledRect(openCommentRect, zoom) : null;
  const openCommentExisting = openComment
    ? commentAt(sheetComments, openComment.row, openComment.col)
    : null;
  const openCommentAuthor = resolveCommentAuthor(openCommentExisting, userName);

  const spacerWidth = sheetInfo ? sheetInfo.contentWidth * zoom : undefined;
  const spacerHeight = sheetInfo ? sheetInfo.contentHeight * zoom : undefined;
  const formulaValue = formulaDraft ?? focusedCell?.input ?? '';
  const normalizedSelection = selection ? normalizeRange(selection) : null;
  const selectionRows = normalizedSelection
    ? normalizedSelection.bottom - normalizedSelection.top + 1
    : 1;
  const selectionColumns = normalizedSelection
    ? normalizedSelection.right - normalizedSelection.left + 1
    : 1;

  // switch sheets: retarget the core, reset scroll + selection, reread info.
  const switchSheet = (index: number) => {
    const handle = handleRef.current;
    if (!handle) return;
    try {
      handle.setActiveSheet(index);
      pendingSheetViewRef.current = true;
      setSelection(selectionAt({ row: 0, col: 0 }));
      setEditing(null);
      setFormulaDraft(null);
      setCapturedFormat(null);
      setOpenFilter(null);
      setOpenComment(null);
      paintSourceRef.current = null;
      setSheetInfo(handle.sheetInfo());
      refreshSheetState(index);
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    }
  };

  return (
    <div
      className={className}
      role="application"
      aria-label={t('editor.appLabel')}
      style={{
        position: 'relative',
        display: 'flex',
        flexDirection: 'column',
        width: '100%',
        height: '100%',
        minWidth: 0,
        color: '#202124',
        background: '#ffffff',
        fontFamily: 'ui-sans-serif, system-ui, sans-serif',
      }}
    >
      <div ref={toolbarRef} data-testid="xlsx-toolbar" style={xlsxToolbarStyles.shell}>
        <EditorToolbar
          currentFormatting={{
            ...selectionFormatting,
            borderStyle: borderStyleChoice ?? selectionFormatting.borderStyle,
            borderColor: borderColorChoice ?? selectionFormatting.borderColor,
            paintFormat: capturedFormat !== null,
          }}
          selectionShape={{
            rows: selectionRows,
            columns: selectionColumns,
            canUnmerge: mergedRanges.length > 0,
          }}
          onSearchMenus={searchMenus}
          onUndo={undo}
          onRedo={redo}
          canUndo={historyState.canUndo}
          canRedo={historyState.canRedo}
          onPrint={print}
          zoom={zoom}
          onZoomChange={setZoom}
          onFormat={formatSelection}
          onMerge={mergeSelection}
        >
          <EditorToolbar.Toolbar />
          <div
            style={xlsxToolbarStyles.rail}
            role="group"
            aria-label={t('toolbar.formulaBarLabel')}
          >
            <ToolbarGroup
              style={{ ...xlsxToolbarStyles.group, paddingLeft: 0 }}
              label={t('toolbar.fileActionsLabel')}
            >
              <ToolbarButton
                testId="xlsx-save"
                onClick={save}
                disabled={!sheetInfo}
                title={t('toolbar.save')}
              >
                <ToolbarIcon name="save" size={18} />
              </ToolbarButton>
              <ToolbarButton
                testId="xlsx-export-png"
                onClick={exportPng}
                disabled={!sheetInfo || !pngExportAvailable}
                title={t('toolbar.exportPng')}
              >
                <ToolbarIcon name="image" size={18} />
              </ToolbarButton>
            </ToolbarGroup>
            <ToolbarGroup
              style={{ ...xlsxToolbarStyles.group }}
              label={t('toolbar.sortLabel')}
            >
              <ToolbarButton
                testId="xlsx-sort-asc"
                onClick={() => sortSelection(true)}
                disabled={!sheetInfo || !selection}
                title={t('toolbar.sortAsc')}
              >
                <ToolbarIcon name="sortAsc" size={18} />
              </ToolbarButton>
              <ToolbarButton
                testId="xlsx-sort-desc"
                onClick={() => sortSelection(false)}
                disabled={!sheetInfo || !selection}
                title={t('toolbar.sortDesc')}
              >
                <ToolbarIcon name="sortDesc" size={18} />
              </ToolbarButton>
            </ToolbarGroup>
            <ToolbarGroup
              style={{ ...xlsxToolbarStyles.group }}
              label={t('toolbar.freezeLabel')}
            >
              <ToolbarDropdown
                testId="xlsx-freeze"
                title={t('toolbar.freezeLabel')}
                disabled={!sheetInfo}
                menuWidth={240}
                trigger={
                  <>
                    <ToolbarIcon name="freeze" size={18} />
                    <ToolbarIcon name="chevronDown" size={13} />
                  </>
                }
              >
                {(close) => (
                  <>
                    <ToolbarMenuItem
                      label={t('toolbar.freezeUpToRow', { row: Math.max(1, anchorRow) })}
                      disabled={anchorRow < 1}
                      onClick={() =>
                        freezePanes(anchorRow, frozenCols > 0 ? frozenCols : null)
                      }
                      close={close}
                    />
                    <ToolbarMenuItem
                      label={t('toolbar.freezeUpToColumn', {
                        column: columnLabel(Math.max(0, anchorCol - 1)),
                      })}
                      disabled={anchorCol < 1}
                      onClick={() =>
                        freezePanes(frozenRows > 0 ? frozenRows : null, anchorCol)
                      }
                      close={close}
                    />
                    <ToolbarMenuSeparator />
                    <ToolbarMenuItem
                      label={t('toolbar.freezeFirstRow')}
                      onClick={() => freezePanes(1, null)}
                      close={close}
                    />
                    <ToolbarMenuItem
                      label={t('toolbar.freezeFirstColumn')}
                      onClick={() => freezePanes(null, 1)}
                      close={close}
                    />
                    <ToolbarMenuSeparator />
                    <ToolbarMenuItem
                      label={t('toolbar.unfreeze')}
                      disabled={frozenRows === 0 && frozenCols === 0}
                      onClick={() => freezePanes(null, null)}
                      close={close}
                    />
                  </>
                )}
              </ToolbarDropdown>
            </ToolbarGroup>
            <ToolbarGroup
              style={{ ...xlsxToolbarStyles.group }}
              label={t('toolbar.filterLabel')}
            >
              <ToolbarButton
                testId="xlsx-filter-toggle"
                onClick={toggleFilter}
                disabled={!sheetInfo || !selection}
                active={activeFilterSpec !== null}
                title={t('toolbar.filterToggle')}
              >
                <ToolbarIcon name="filter" size={18} />
              </ToolbarButton>
            </ToolbarGroup>
            <ToolbarGroup
              style={{ ...xlsxToolbarStyles.group }}
              label={t('toolbar.commentLabel')}
            >
              <span data-xlsx-comment-trigger style={{ display: 'inline-flex' }}>
                <ToolbarButton
                  testId="xlsx-comment-toggle"
                  onClick={toggleCommentEditor}
                  disabled={
                    !sheetInfo || !selection || selectionRows > 1 || selectionColumns > 1
                  }
                  active={openComment?.mode === 'edit'}
                  title={t('toolbar.commentToggle')}
                >
                  <ToolbarIcon name="comment" size={18} />
                </ToolbarButton>
              </span>
            </ToolbarGroup>
            <div
              style={xlsxToolbarStyles.formulaGroup}
              role="group"
              aria-label={t('toolbar.formulaBarLabel')}
            >
              <input
                data-testid="xlsx-name-box"
                readOnly
                value={focusedCell?.a1 ?? ''}
                placeholder={t('toolbar.nameBoxPlaceholder')}
                aria-label={t('toolbar.nameBoxPlaceholder')}
                style={xlsxToolbarStyles.nameBox}
              />
              <span style={xlsxToolbarStyles.formulaMark} aria-hidden="true">
                fx
              </span>
              <input
                data-testid="xlsx-formula-input"
                value={formulaValue}
                placeholder={t('toolbar.formulaPlaceholder')}
                aria-label={t('toolbar.formulaPlaceholder')}
                disabled={!sheetInfo}
                onChange={(e) => setFormulaDraft(e.target.value)}
                onKeyDown={(e) => {
                  if (e.key === 'Enter') {
                    commitFormula(e.shiftKey ? 'up' : 'down');
                    focusContainer();
                    e.preventDefault();
                  } else if (e.key === 'Escape') {
                    setFormulaDraft(null);
                    focusContainer();
                    e.preventDefault();
                  }
                }}
                onBlur={() => commitFormula()}
                style={xlsxToolbarStyles.formulaInput}
              />
            </div>
            <PresenceStrip
              peers={awarenessPeers}
              sheetIds={sheetInfo?.sheetIds ?? []}
              sheetNames={sheetInfo?.sheetNames ?? []}
              activeSheet={activeSheet}
            />
            {proposalsAvailable && (
              <div style={xlsxToolbarStyles.proposals}>
                <ToolbarButton
                  testId="xlsx-proposals-button"
                  onClick={() => setProposalsPanelOpen((open) => !open)}
                  disabled={!sheetInfo}
                  active={proposalsPanelOpen}
                  ariaExpanded={proposalsPanelOpen}
                  title={t('proposals.panelLabel')}
                  style={{ width: proposals.length > 0 ? 42 : 28 }}
                >
                  <ToolbarIcon name="proposals" size={18} />
                  {proposals.length > 0 && (
                    <span data-testid="xlsx-proposals-count" style={xlsxToolbarStyles.count}>
                      {proposals.length}
                    </span>
                  )}
                </ToolbarButton>
              </div>
            )}
          </div>
        </EditorToolbar>
      </div>
      {proposalsAvailable && proposalsPanelOpen && (
        <ProposalsPanel
          proposals={proposals}
          staleFor={staleFor}
          onAccept={acceptProposal}
          onReject={rejectProposal}
          style={{
            top: toolbarHeight + 4,
            right: 8,
            width: 'min(320px, calc(100% - 16px))',
            maxHeight: `min(420px, calc(100% - ${toolbarHeight + 12}px))`,
          }}
        />
      )}

      <div
        ref={scrollRef}
        data-testid="xlsx-scroll"
        tabIndex={0}
        onKeyDown={onKeyDown}
        onMouseDown={onMouseDown}
        onMouseMove={onMouseMove}
        onMouseLeave={onMouseLeave}
        onClick={onClick}
        onDoubleClick={onDoubleClick}
        style={{ position: 'relative', flex: 1, overflow: 'auto', minHeight: 0, outline: 'none' }}
      >
        <div
          style={{
            position: 'absolute',
            top: 0,
            left: 0,
            width: spacerWidth ?? '100%',
            height: spacerHeight ?? '100%',
          }}
        />
        {/* one sticky layer pins the canvas and overlays to the viewport top-left
            so overlay children share the canvas's coordinate space — a separate
            sticky sibling would sit below the full-height canvas in flow and
            scroll-jump when a child (the in-cell editor) is focused. */}
        <div style={{ position: 'sticky', top: 0, left: 0, width: 0, height: 0 }}>
          <canvas
            ref={canvasRef}
            style={{ display: 'block', position: 'absolute', top: 0, left: 0 }}
          />

          <div
            data-testid="xlsx-overlay-host"
            style={{
              position: 'absolute',
              top: 0,
              left: 0,
              width: 0,
              height: 0,
              pointerEvents: 'none',
            }}
          >
            {scaledSelectionRect && (
              <div
                data-testid="xlsx-selection"
                style={{
                  position: 'absolute',
                  left: scaledSelectionRect.x,
                  top: scaledSelectionRect.y,
                  width: scaledSelectionRect.w,
                  height: scaledSelectionRect.h,
                  boxSizing: 'border-box',
                  border: `1px solid ${BRAND}`,
                  background: 'rgba(33, 115, 70, 0.12)',
                }}
              />
            )}
            {scaledFocusRect && !editing && (
              <div
                style={{
                  position: 'absolute',
                  left: scaledFocusRect.x,
                  top: scaledFocusRect.y,
                  width: scaledFocusRect.w,
                  height: scaledFocusRect.h,
                  boxSizing: 'border-box',
                  border: `2px solid ${BRAND}`,
                }}
              />
            )}
            <RemoteSelections
              peers={awarenessPeers}
              grid={grid}
              sheetIds={sheetInfo?.sheetIds ?? []}
              activeSheet={activeSheet}
              zoom={zoom}
              mergedRanges={visibleMergedRanges}
            />
            {filterButtonRects.map(({ col, rect }) => {
              const size = Math.min(16, Math.max(10, rect.h - 6));
              const constrained = activeFilterSpec
                ? hasCriteria(columnCriteria(activeFilterSpec, col))
                : false;
              return (
                <button
                  key={col}
                  type="button"
                  data-xlsx-filter-trigger
                  data-testid={`xlsx-filter-button-${col}`}
                  aria-label={t('filter.columnButton', { column: columnLabel(col) })}
                  aria-expanded={openFilter?.col === col}
                  onMouseDown={(e) => {
                    e.preventDefault();
                    e.stopPropagation();
                  }}
                  onDoubleClick={(e) => e.stopPropagation()}
                  onClick={(e) => {
                    e.stopPropagation();
                    openFilterFor(col);
                  }}
                  style={{
                    position: 'absolute',
                    left: rect.x + rect.w - size - 2,
                    top: rect.y + (rect.h - size) / 2,
                    width: size,
                    height: size,
                    display: 'grid',
                    placeItems: 'center',
                    padding: 0,
                    border: `1px solid ${constrained ? BRAND : '#c7cacf'}`,
                    borderRadius: 3,
                    background: constrained ? BRAND : '#ffffff',
                    color: constrained ? '#ffffff' : '#3c4043',
                    cursor: 'pointer',
                    pointerEvents: 'auto',
                    boxSizing: 'border-box',
                  }}
                >
                  <ToolbarIcon name="chevronDown" size={size - 4} />
                </button>
              );
            })}
            {openFilter && scaledOpenFilterRect && openFilterCriteria && (
              <FilterPopover
                key={`${activeSheet}:${openFilter.col}`}
                columnName={columnLabel(openFilter.col)}
                values={openFilter.values}
                hasBlanks={openFilter.hasBlanks}
                truncated={openFilter.truncated}
                initialValues={openFilterCriteria.values}
                initialShowBlanks={openFilterCriteria.showBlanks}
                style={{
                  position: 'absolute',
                  left: Math.max(0, scaledOpenFilterRect.x),
                  top: scaledOpenFilterRect.y + scaledOpenFilterRect.h + 2,
                  pointerEvents: 'auto',
                }}
                onApply={(values, showBlanks) =>
                  applyColumnFilter(openFilter.col, values, showBlanks)
                }
                onClear={() => applyColumnFilter(openFilter.col, null, true, true)}
                onClose={() => setOpenFilter(null)}
              />
            )}
            {commentIndicators.map(({ comment, rect }) => (
              <div
                key={`${comment.row}:${comment.col}`}
                data-testid={`xlsx-comment-indicator-${comment.row}-${comment.col}`}
                style={{
                  position: 'absolute',
                  left: rect.x + rect.w - 7,
                  top: rect.y + 1,
                  width: 0,
                  height: 0,
                  borderTop: '6px solid #d93025',
                  borderLeft: '6px solid transparent',
                }}
              />
            ))}
            {openComment?.mode === 'view' && openCommentExisting && scaledOpenCommentRect && (
              <CommentViewPopover
                cellName={cellA1(openComment.row, openComment.col)}
                author={openCommentExisting.author}
                text={openCommentExisting.text}
                style={{
                  position: 'absolute',
                  left: Math.max(0, scaledOpenCommentRect.x + scaledOpenCommentRect.w + 2),
                  top: Math.max(0, scaledOpenCommentRect.y),
                  pointerEvents: 'auto',
                }}
                onClose={() => setOpenComment(null)}
              />
            )}
            {openComment?.mode === 'edit' && scaledOpenCommentRect && (
              <CommentEditorPopover
                key={`${activeSheet}:${openComment.row}:${openComment.col}`}
                cellName={cellA1(openComment.row, openComment.col)}
                author={openCommentAuthor}
                initialText={openCommentExisting?.text ?? ''}
                canDelete={openCommentExisting !== null}
                style={{
                  position: 'absolute',
                  left: Math.max(0, scaledOpenCommentRect.x + scaledOpenCommentRect.w + 2),
                  top: Math.max(0, scaledOpenCommentRect.y),
                  pointerEvents: 'auto',
                }}
                onSave={(text) =>
                  setCellComment(openComment.row, openComment.col, {
                    author: openCommentAuthor,
                    text,
                  })
                }
                onDelete={() => setCellComment(openComment.row, openComment.col, null)}
                onClose={() => setOpenComment(null)}
              />
            )}
            {editing && scaledEditRect && (
              <input
                ref={editorInputRef}
                data-testid="xlsx-cell-editor"
                value={editing.value}
                onChange={(e) =>
                  setEditing((prev) => (prev ? { ...prev, value: e.target.value } : prev))
                }
                onKeyDown={(e) => {
                  e.stopPropagation();
                  if (e.key === 'Enter') {
                    commitEditor(e.shiftKey ? 'up' : 'down');
                    e.preventDefault();
                  } else if (e.key === 'Tab') {
                    commitEditor(e.shiftKey ? 'left' : 'right');
                    e.preventDefault();
                  } else if (e.key === 'Escape') {
                    cancelEditor();
                    e.preventDefault();
                  }
                }}
                onBlur={() => {
                  if (suppressBlurRef.current) {
                    suppressBlurRef.current = false;
                    return;
                  }
                  commitEditor();
                }}
                style={{
                  position: 'absolute',
                  left: scaledEditRect.x,
                  top: scaledEditRect.y,
                  width: scaledEditRect.w,
                  height: scaledEditRect.h,
                  boxSizing: 'border-box',
                  border: `2px solid ${BRAND}`,
                  padding: '0 3px',
                  font: `${13 * zoom}px system-ui, sans-serif`,
                  background: '#ffffff',
                  pointerEvents: 'auto',
                  outline: 'none',
                }}
              />
            )}
          </div>
        </div>
      </div>

      {a11yGrid && (
        <div style={visuallyHidden} role="grid" aria-label={a11yGrid.label}>
          <div role="row">
            <span role="columnheader" />
            {a11yGrid.columnHeaders.map((h) => (
              <span key={h.col} role="columnheader">
                {h.label}
              </span>
            ))}
          </div>
          {a11yGrid.rows.map((r) => (
            <div key={r.row} role="row">
              <span role="rowheader">{r.header}</span>
              {r.cells.map((c) => (
                <span key={c.col} role="gridcell" aria-selected={c.selected}>
                  {c.label}
                </span>
              ))}
            </div>
          ))}
        </div>
      )}

      {error && (
        <div
          data-testid="xlsx-error"
          role="alert"
          style={{
            position: 'absolute',
            inset: 0,
            display: 'grid',
            placeItems: 'center',
            padding: 16,
            textAlign: 'center',
            color: '#b00020',
          }}
        >
          {t('editor.openError')}: {error}
        </div>
      )}

      {sheetInfo && sheetInfo.sheetNames.length > 0 && (
        <div
          data-testid="xlsx-sheet-tabs"
          role="tablist"
          aria-label={t('editor.sheetTabsLabel')}
          style={{
            display: 'flex',
            gap: 2,
            padding: '4px 6px',
            borderTop: '1px solid #e0e0e0',
            background: '#fafafa',
            overflowX: 'auto',
          }}
        >
          {sheetInfo.sheetNames.map((name, i) => {
            const active = i === sheetInfo.activeSheet;
            return (
              <button
                key={i}
                role="tab"
                aria-selected={active}
                onClick={() => switchSheet(i)}
                style={{
                  border: 'none',
                  padding: '4px 12px',
                  cursor: 'pointer',
                  borderBottom: active ? `2px solid ${BRAND}` : '2px solid transparent',
                  fontWeight: active ? 600 : 400,
                  background: active ? '#ffffff' : 'transparent',
                }}
              >
                {name}
              </button>
            );
          })}
        </div>
      )}
    </div>
  );
}
