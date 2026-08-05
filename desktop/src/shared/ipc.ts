export const IPC = {
  locale: 'app:locale',
  localeChanged: 'locale:changed',
  openFile: 'dialog:openFile',
  saveFile: 'dialog:saveFile',
  saveFileAs: 'dialog:saveFileAs',
  exportPdf: 'dialog:exportPdf',
  confirmUnsaved: 'dialog:confirmUnsaved',
  readFile: 'fs:readFile',
  platform: 'app:platform',
  menuAction: 'menu:action',
  webEditAction: 'edit:webAction',
  documentKind: 'app:documentKind',
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
} as const

export const PRINT_PAGE_CAP = 100

export type DocumentKind = 'docx' | 'xlsx' | 'pptx'

export type UnsavedChoice = 'save' | 'discard' | 'cancel'

export interface OpenedDocument {
  path: string
  name: string
  kind: DocumentKind | null
  data: Uint8Array
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
  | 'view:freezeTopRow'
  | 'view:freezeFirstColumn'
  | 'view:unfreeze'
  | 'view:zoomIn'
  | 'view:zoomOut'
  | 'view:zoomReset'

export const EXTENSIONS: Record<DocumentKind, string> = {
  docx: 'docx',
  xlsx: 'xlsx',
  pptx: 'pptx',
}

export function kindFromPath(filePath: string): DocumentKind | null {
  const ext = filePath.slice(filePath.lastIndexOf('.') + 1).toLowerCase()
  return ext === 'docx' || ext === 'xlsx' || ext === 'pptx' ? ext : null
}
