export const IPC = {
  locale: 'app:locale',
  localeChanged: 'locale:changed',
  openFile: 'dialog:openFile',
  saveFile: 'dialog:saveFile',
  saveFileAs: 'dialog:saveFileAs',
  exportPdf: 'dialog:exportPdf',
  printDocument: 'dialog:printDocument',
  newDocument: 'file:newDocument',
  recentsList: 'file:recentsList',
  recentsRemove: 'file:recentsRemove',
  openRecent: 'file:openRecent',
  confirmUnsaved: 'dialog:confirmUnsaved',
  confirmSaveRefused: 'dialog:confirmSaveRefused',
  readFile: 'fs:readFile',
  platform: 'app:platform',
  version: 'app:version',
  menuAction: 'menu:action',
  webEditAction: 'edit:webAction',
  documentKind: 'app:documentKind',
  diagnostics: 'app:diagnostics',
  diagnosticsSample: 'app:diagnosticsSample',
  editCapabilities: 'app:editCapabilities',
  rendererReady: 'renderer:ready',
  closeRequest: 'window:closeRequest',
  closeResponse: 'window:closeResponse',
  printReady: 'print:ready',
  printRender: 'print:render',
  printRendered: 'print:rendered',
  updateEvent: 'update:event',
  updateCheck: 'update:check',
  updateInstall: 'update:install',
  agentSettingsGet: 'agent:settingsGet',
  agentSettingsSet: 'agent:settingsSet',
  agentRun: 'agent:run',
  agentCancel: 'agent:cancel',
  agentEvent: 'agent:event',
  agentToolRequest: 'agent:toolRequest',
  agentToolResult: 'agent:toolResult',
} as const

export const PRINT_PAGE_CAP = 100

export type DocumentKind = 'docx' | 'xlsx' | 'pptx'

export type UnsavedChoice = 'save' | 'discard' | 'cancel'

// The way out of a save the writer refused. Retrying is not offered: the same
// change would be refused again, so the only choices are to leave the change
// behind or to keep the document open and edit it into something writable.
export type RefusedChoice = 'discard' | 'cancel'

export interface RecentFile {
  path: string
  name: string
  kind: DocumentKind | null
  exists: boolean
}

export interface OpenedDocument {
  path: string
  name: string
  kind: DocumentKind | null
  data: Uint8Array
  /** Milliseconds the main process spent reading the bytes off disk. */
  readMs?: number
  /** `Date.now()` as the main process handed the bytes over, so the renderer
   * can charge the structured-clone transfer to the right phase. */
  sentAt?: number
}

export interface SaveRequest {
  path: string | null
  kind: DocumentKind
  data: Uint8Array
}

export interface SaveResult {
  path: string | null
  canceled: boolean
}

export interface ExportPdfRequest {
  name: string
  kind: DocumentKind
  data: Uint8Array
}

// `truncated` and `skipped` are separate outcomes and never stand in for each
// other: the export stopped at the page cap, or individual pages failed to
// render and were left out. An export can be both, or neither.
export interface ExportPdfResult {
  path: string | null
  canceled: boolean
  pages?: number
  truncated?: boolean
  skipped?: number
}

export interface PrintResult {
  printed: boolean
  pages?: number
  truncated?: boolean
  skipped?: number
}

export interface PrintJob {
  kind: DocumentKind
  data: Uint8Array
}

export type PrintRenderResult =
  | { ok: true; pages: number; truncated: boolean; skippedPages: number[] }
  | { ok: false; error: string }

export type WebEditAction = 'undo' | 'redo' | 'cut' | 'copy' | 'paste' | 'delete' | 'selectAll'

// Which Edit-menu verbs the editor showing right now can actually carry out.
// The pptx engine has no shape clipboard, so cut, copy, paste and delete
// no-op unless a caret or range is live inside a text story; a menu item that
// stays enabled there does nothing and lies about it.
export interface EditCapabilities {
  cut: boolean
  copy: boolean
  paste: boolean
  delete: boolean
  selectAll: boolean
}

// Editors that accept every verb whenever they are focused — and the state
// before any editor reports otherwise — leaving the menu as it always was.
export const ALL_EDIT_CAPABILITIES: EditCapabilities = {
  cut: true,
  copy: true,
  paste: true,
  delete: true,
  selectAll: true,
}

export function sameEditCapabilities(a: EditCapabilities, b: EditCapabilities): boolean {
  return (
    a.cut === b.cut &&
    a.copy === b.copy &&
    a.paste === b.paste &&
    a.delete === b.delete &&
    a.selectAll === b.selectAll
  )
}

// The payload crosses the IPC boundary as a plain value, so its shape is
// checked before it decides which menu items the user can reach.
export function readEditCapabilities(value: unknown): EditCapabilities | null {
  if (typeof value !== 'object' || value === null) return null
  const { cut, copy, paste, delete: remove, selectAll } = value as Partial<EditCapabilities>
  if (
    typeof cut !== 'boolean' ||
    typeof copy !== 'boolean' ||
    typeof paste !== 'boolean' ||
    typeof remove !== 'boolean' ||
    typeof selectAll !== 'boolean'
  ) {
    return null
  }
  return { cut, copy, paste, delete: remove, selectAll }
}

// ---------------------------------------------------------------------------
// Diagnostics. Chromium runs Web Workers as threads inside the renderer, so
// `app.getAppMetrics()` reports one number covering the resident layout
// engine, the image cache and the JS heap together. The renderer is the only
// place that can take those apart, so it pushes its own account over and the
// main process merges it into the report.

/** Which document the renderer is holding, for the process label. */
export interface DocumentProfile {
  kind: DocumentKind
  name: string
  /** Size of the seed bytes as opened from disk. */
  bytes: number
}

/**
 * Wall-clock cost of each phase of the last document open, in milliseconds.
 * Every phase is measured where it actually happens rather than inferred, so
 * a missing one means it was not observed — never that it was instant.
 */
export interface OpenPhaseTimings {
  /** Main-process disk read. */
  read?: number
  /** Hand-off from the main process to the renderer (structured clone). */
  transfer?: number
  /** Editor construction, parse and CRDT seed, to the first frame after the
   * document is committed to React. */
  mount?: number
  /** From that first frame to the one where the main thread is no longer
   * blocked — where a heavy document stops being frozen. */
  interactive?: number
}

export interface MemoryBreakdownRow {
  label: string
  bytes: number
}

export interface RendererDiagnostics {
  document: DocumentProfile | null
  open: OpenPhaseTimings | null
  memory: MemoryBreakdownRow[]
}

function readOpenPhaseTimings(value: unknown): OpenPhaseTimings | null {
  if (typeof value !== 'object' || value === null) return null
  const timings: OpenPhaseTimings = {}
  for (const phase of ['read', 'transfer', 'mount', 'interactive'] as const) {
    const ms = (value as Record<string, unknown>)[phase]
    if (typeof ms === 'number' && Number.isFinite(ms) && ms >= 0) timings[phase] = ms
  }
  return timings
}

// The payload crosses the IPC boundary as a plain value and is rendered into a
// dialog, so its shape is checked rather than trusted.
export function readRendererDiagnostics(value: unknown): RendererDiagnostics | null {
  if (typeof value !== 'object' || value === null) return null
  const { document, open, memory } = value as Partial<RendererDiagnostics>

  let profile: DocumentProfile | null = null
  if (typeof document === 'object' && document !== null) {
    const { kind, name, bytes } = document
    const known = kind === 'docx' || kind === 'xlsx' || kind === 'pptx'
    if (!known || typeof name !== 'string' || typeof bytes !== 'number') return null
    profile = { kind, name, bytes: Number.isFinite(bytes) && bytes >= 0 ? bytes : 0 }
  }

  const rows: MemoryBreakdownRow[] = []
  if (Array.isArray(memory)) {
    for (const row of memory) {
      if (typeof row !== 'object' || row === null) continue
      const { label, bytes } = row as Partial<MemoryBreakdownRow>
      if (typeof label !== 'string' || typeof bytes !== 'number') continue
      if (!Number.isFinite(bytes) || bytes <= 0) continue
      rows.push({ label, bytes })
    }
  }

  return { document: profile, open: readOpenPhaseTimings(open), memory: rows }
}

export type UpdateEvent =
  | { status: 'checking' }
  | { status: 'available'; version: string }
  | { status: 'progress'; percent: number; transferred: number; total: number; bytesPerSecond: number }
  | { status: 'downloaded'; version: string }
  | { status: 'none' }
  | { status: 'dev' }
  | { status: 'error'; message: string }

export type MenuAction =
  | 'file:new'
  | 'file:open'
  | 'file:save'
  | 'file:saveAs'
  | 'file:exportPdf'
  | 'file:print'
  | 'edit:undo'
  | 'edit:redo'
  | 'edit:cut'
  | 'edit:copy'
  | 'edit:paste'
  | 'edit:delete'
  | 'edit:selectAll'
  | 'edit:find'
  | 'view:wordCount'
  | 'view:spellCheck'
  | 'view:aiAssistant'
  | 'view:freezeTopRow'
  | 'view:freezeFirstColumn'
  | 'view:unfreeze'
  | 'view:zoomIn'
  | 'view:zoomOut'
  | 'view:zoomReset'

// ---------------------------------------------------------------------------
// AI assistant. The model runs in the main process (which owns the API key);
// document tools execute in the renderer against the live editor, so each tool
// call round-trips over IPC: main sends agentToolRequest, the renderer answers
// on agentToolResult with the matching id.

export type AgentProvider = 'deepseek'

export const AGENT_DEFAULT_MODELS: Record<AgentProvider, string> = {
  deepseek: 'deepseek-chat',
}

/** What the renderer may know about the configuration — never the key itself. */
export interface AgentSettings {
  provider: AgentProvider
  model: string
  hasApiKey: boolean
}

export interface AgentSettingsUpdate {
  provider?: AgentProvider
  model?: string
  /** Empty string clears the stored key; undefined leaves it untouched. */
  apiKey?: string
}

export interface AgentRunRequest {
  /** The whole visible transcript; main is stateless between runs. */
  messages: Array<{ role: 'user' | 'assistant'; content: string }>
  documentKind: DocumentKind
  documentName: string
  locale: string
}

export type AgentEvent =
  | { type: 'text'; delta: string }
  | { type: 'tool'; name: string; summary: string }
  | { type: 'done' }
  | { type: 'error'; message: string }

export interface AgentToolRequest {
  id: string
  name: string
  args: Record<string, unknown>
}

export interface AgentToolResult {
  id: string
  /** JSON-serializable payload, or an `error` string the model can read. */
  result: unknown
}

export const EXTENSIONS: Record<DocumentKind, string> = {
  docx: 'docx',
  xlsx: 'xlsx',
  pptx: 'pptx',
}

export function kindFromPath(filePath: string): DocumentKind | null {
  const ext = filePath.slice(filePath.lastIndexOf('.') + 1).toLowerCase()
  return ext === 'docx' || ext === 'xlsx' || ext === 'pptx' ? ext : null
}
