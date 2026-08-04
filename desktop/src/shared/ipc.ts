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

export interface ExportPdfResult {
  path: string | null
  canceled: boolean
  pages?: number
  truncated?: boolean
}

export interface PrintJob {
  kind: DocumentKind
  data: Uint8Array
}

export type PrintRenderResult =
  | { ok: true; pages: number; truncated: boolean }
  | { ok: false; error: string }

export type WebEditAction = 'undo' | 'redo' | 'cut' | 'copy' | 'paste' | 'delete' | 'selectAll'

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
