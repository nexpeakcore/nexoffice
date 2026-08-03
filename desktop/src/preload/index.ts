import { contextBridge, ipcRenderer } from 'electron'
import {
  IPC,
  type DocumentKind,
  type ExportPdfResult,
  type MenuAction,
  type OpenedDocument,
  type PrintJob,
  type PrintRenderResult,
  type SaveResult,
  type UnsavedChoice,
  type WebEditAction,
} from '../shared/ipc.js'

const api = {
  platform: (): Promise<NodeJS.Platform> => ipcRenderer.invoke(IPC.platform),

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

  webEditAction: (action: WebEditAction): void => ipcRenderer.send(IPC.webEditAction, action),

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
}

export type NexOfficeApi = typeof api

contextBridge.exposeInMainWorld('nexoffice', api)
