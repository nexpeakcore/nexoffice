import { contextBridge, ipcRenderer } from 'electron'
import {
  IPC,
  type DocumentKind,
  type EditCapabilities,
  type ExportPdfResult,
  type MenuAction,
  type OpenedDocument,
  type PrintJob,
  type PrintRenderResult,
  type RefusedChoice,
  type SaveResult,
  type UnsavedChoice,
  type UpdateEvent,
  type WebEditAction,
} from '../shared/ipc.js'
import type { LocaleCode } from '../i18n/index.js'

const api = {
  platform: (): Promise<NodeJS.Platform> => ipcRenderer.invoke(IPC.platform),

  getLocale: (): Promise<LocaleCode> => ipcRenderer.invoke(IPC.locale),

  onLocaleChanged: (handler: (locale: LocaleCode) => void): (() => void) => {
    const listener = (_event: Electron.IpcRendererEvent, locale: LocaleCode) => handler(locale)
    ipcRenderer.on(IPC.localeChanged, listener)
    return () => ipcRenderer.removeListener(IPC.localeChanged, listener)
  },

  openFile: (): Promise<OpenedDocument | null> => ipcRenderer.invoke(IPC.openFile),

  readFile: (path: string): Promise<OpenedDocument> => ipcRenderer.invoke(IPC.readFile, path),

  saveFile: (path: string | null, kind: DocumentKind, data: Uint8Array): Promise<SaveResult> =>
    ipcRenderer.invoke(IPC.saveFile, { path, kind, data }),

  saveFileAs: (path: string | null, kind: DocumentKind, data: Uint8Array): Promise<SaveResult> =>
    ipcRenderer.invoke(IPC.saveFileAs, { path, kind, data }),

  exportPdf: (name: string, kind: DocumentKind, data: Uint8Array): Promise<ExportPdfResult> =>
    ipcRenderer.invoke(IPC.exportPdf, { name, kind, data }),

  printReady: (): void => ipcRenderer.send(IPC.printReady),

  printRendered: (result: PrintRenderResult): void => ipcRenderer.send(IPC.printRendered, result),

  onPrintRender: (handler: (job: PrintJob) => void): (() => void) => {
    const listener = (_event: Electron.IpcRendererEvent, job: PrintJob) => handler(job)
    ipcRenderer.on(IPC.printRender, listener)
    return () => ipcRenderer.removeListener(IPC.printRender, listener)
  },

  confirmUnsaved: (name: string): Promise<UnsavedChoice> =>
    ipcRenderer.invoke(IPC.confirmUnsaved, name),

  confirmSaveRefused: (name: string, message: string): Promise<RefusedChoice> =>
    ipcRenderer.invoke(IPC.confirmSaveRefused, { name, message }),

  webEditAction: (action: WebEditAction): void => ipcRenderer.send(IPC.webEditAction, action),

  setDocumentKind: (kind: DocumentKind | null): void => ipcRenderer.send(IPC.documentKind, kind),

  setEditCapabilities: (capabilities: EditCapabilities): void =>
    ipcRenderer.send(IPC.editCapabilities, capabilities),

  rendererReady: (): void => ipcRenderer.send(IPC.rendererReady),

  resolveClose: (proceed: boolean): void => ipcRenderer.send(IPC.closeResponse, proceed),

  onCloseRequest: (handler: () => void): (() => void) => {
    const listener = () => handler()
    ipcRenderer.on(IPC.closeRequest, listener)
    return () => ipcRenderer.removeListener(IPC.closeRequest, listener)
  },

  onMenuAction: (handler: (action: MenuAction) => void): (() => void) => {
    const listener = (_event: Electron.IpcRendererEvent, action: MenuAction) => handler(action)
    ipcRenderer.on(IPC.menuAction, listener)
    return () => ipcRenderer.removeListener(IPC.menuAction, listener)
  },

  onFileOpened: (handler: (document: OpenedDocument) => void): (() => void) => {
    const listener = (_event: Electron.IpcRendererEvent, document: OpenedDocument) => handler(document)
    ipcRenderer.on(IPC.readFile, listener)
    return () => ipcRenderer.removeListener(IPC.readFile, listener)
  },

  checkForUpdates: (): Promise<void> => ipcRenderer.invoke(IPC.updateCheck),

  installUpdate: (): void => ipcRenderer.send(IPC.updateInstall),

  onUpdateEvent: (handler: (event: UpdateEvent) => void): (() => void) => {
    const listener = (_event: Electron.IpcRendererEvent, update: UpdateEvent) => handler(update)
    ipcRenderer.on(IPC.updateEvent, listener)
    return () => ipcRenderer.removeListener(IPC.updateEvent, listener)
  },
}

export type NexOfficeApi = typeof api

contextBridge.exposeInMainWorld('nexoffice', api)
