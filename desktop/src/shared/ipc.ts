export const IPC = {
  openFile: 'dialog:openFile',
  saveFile: 'dialog:saveFile',
  saveFileAs: 'dialog:saveFileAs',
  exportPdf: 'dialog:exportPdf',
  confirmUnsaved: 'dialog:confirmUnsaved',
  readFile: 'fs:readFile',
  platform: 'app:platform',
  menuAction: 'menu:action',
  rendererReady: 'renderer:ready',
  closeRequest: 'window:closeRequest',
  closeResponse: 'window:closeResponse',
} as const

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

export type MenuAction =
  | 'file:new'
  | 'file:open'
  | 'file:save'
  | 'file:saveAs'
  | 'file:exportPdf'
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
