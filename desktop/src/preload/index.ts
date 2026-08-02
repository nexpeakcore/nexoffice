import { contextBridge, ipcRenderer } from 'electron'
import {
  IPC,
  type DocumentKind,
  type MenuAction,
  type OpenedDocument,
  type SaveResult,
  type UnsavedChoice,
} from '../shared/ipc.js'

const api = {
  platform: (): Promise<NodeJS.Platform> => ipcRenderer.invoke(IPC.platform),

  openFile: (): Promise<OpenedDocument | null> => ipcRenderer.invoke(IPC.openFile),

  readFile: (path: string): Promise<OpenedDocument> => ipcRenderer.invoke(IPC.readFile, path),

  saveFile: (path: string | null, kind: DocumentKind, data: Uint8Array): Promise<SaveResult> =>
    ipcRenderer.invoke(IPC.saveFile, { path, kind, data }),

  saveFileAs: (path: string | null, kind: DocumentKind, data: Uint8Array): Promise<SaveResult> =>
    ipcRenderer.invoke(IPC.saveFileAs, { path, kind, data }),

  exportPdf: (defaultName: string): Promise<{ path: string | null; canceled: boolean }> =>
    ipcRenderer.invoke(IPC.exportPdf, defaultName),

  confirmUnsaved: (name: string): Promise<UnsavedChoice> =>
    ipcRenderer.invoke(IPC.confirmUnsaved, name),

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
