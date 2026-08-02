import { basename, extname, join, resolve, sep } from 'node:path'
import { readFile, writeFile } from 'node:fs/promises'
import { app, BrowserWindow, dialog, ipcMain, protocol, shell } from 'electron'
import { buildMenu } from './menu.js'
import {
  EXTENSIONS,
  IPC,
  kindFromPath,
  type DocumentKind,
  type MenuAction,
  type OpenedDocument,
  type SaveRequest,
  type SaveResult,
  type UnsavedChoice,
} from '../shared/ipc.js'

const isDev = !app.isPackaged
let mainWindow: BrowserWindow | null = null
let rendererReady = false
let quitting = false
const pendingOpenPaths: string[] = []
const grantedPaths = new Set<string>()

const APP_SCHEME = 'app'
const APP_ORIGIN = `${APP_SCHEME}://bundle`

const MIME_TYPES: Record<string, string> = {
  '.html': 'text/html',
  '.js': 'text/javascript',
  '.mjs': 'text/javascript',
  '.css': 'text/css',
  '.json': 'application/json',
  '.map': 'application/json',
  '.wasm': 'application/wasm',
  '.svg': 'image/svg+xml',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.gif': 'image/gif',
  '.ico': 'image/x-icon',
  '.woff': 'font/woff',
  '.woff2': 'font/woff2',
  '.ttf': 'font/ttf',
  '.otf': 'font/otf',
  '.txt': 'text/plain',
  '.dic': 'text/plain',
  '.aff': 'text/plain',
}

protocol.registerSchemesAsPrivileged([
  { scheme: APP_SCHEME, privileges: { standard: true, secure: true, supportFetchAPI: true } },
])

const FILTERS: Record<DocumentKind, Electron.FileFilter> = {
  docx: { name: 'Word Document', extensions: ['docx'] },
  xlsx: { name: 'Excel Workbook', extensions: ['xlsx'] },
  pptx: { name: 'PowerPoint Presentation', extensions: ['pptx'] },
}

function registerAppProtocol(): void {
  const rendererRoot = join(import.meta.dirname, '../renderer')

  protocol.handle(APP_SCHEME, async (request) => {
    const { pathname } = new URL(request.url)
    const relative = decodeURIComponent(pathname)
    const target = resolve(rendererRoot, relative === '/' || relative === '' ? 'index.html' : `.${relative}`)
    if (target !== rendererRoot && !target.startsWith(rendererRoot + sep)) {
      return new Response('Forbidden', { status: 403 })
    }

    try {
      const data = await readFile(target)
      const mime = MIME_TYPES[extname(target).toLowerCase()] ?? 'application/octet-stream'
      return new Response(new Uint8Array(data), { headers: { 'Content-Type': mime } })
    } catch {
      return new Response('Not Found', { status: 404 })
    }
  })
}

function createWindow(): BrowserWindow {
  rendererReady = false

  const window = new BrowserWindow({
    width: 1440,
    height: 900,
    minWidth: 960,
    minHeight: 600,
    show: false,
    backgroundColor: '#ffffff',
    titleBarStyle: process.platform === 'darwin' ? 'hiddenInset' : 'default',
    webPreferences: {
      preload: join(import.meta.dirname, '../preload/index.cjs'),
      sandbox: true,
      contextIsolation: true,
      nodeIntegration: false,
    },
  })

  window.once('ready-to-show', () => window.show())

  window.webContents.setWindowOpenHandler(({ url }) => {
    let scheme: string | null = null
    try {
      scheme = new URL(url).protocol
    } catch {
      scheme = null
    }
    if (scheme === 'https:' || scheme === 'mailto:') void shell.openExternal(url)
    return { action: 'deny' }
  })

  window.webContents.on('will-navigate', (event) => event.preventDefault())

  window.on('close', (event) => {
    if (!rendererReady || window.webContents.isDestroyed()) return
    event.preventDefault()
    window.webContents.send(IPC.closeRequest)
  })

  window.on('closed', () => {
    if (mainWindow === window) {
      mainWindow = null
      rendererReady = false
    }
  })

  const devServerUrl = process.env['ELECTRON_RENDERER_URL']
  if (isDev && devServerUrl) {
    void window.loadURL(devServerUrl)
  } else {
    void window.loadURL(`${APP_ORIGIN}/index.html`)
  }

  return window
}

async function readDocument(filePath: string): Promise<OpenedDocument> {
  const data = await readFile(filePath)
  return {
    path: filePath,
    name: basename(filePath),
    kind: kindFromPath(filePath),
    data: new Uint8Array(data),
  }
}

function sendMenuAction(action: MenuAction): void {
  if (!mainWindow || mainWindow.isDestroyed()) return
  mainWindow.webContents.send(IPC.menuAction, action)
}

async function promptOpen(): Promise<OpenedDocument | null> {
  const owner = mainWindow
  if (!owner || owner.isDestroyed()) return null

  const { canceled, filePaths } = await dialog.showOpenDialog(owner, {
    properties: ['openFile'],
    filters: [
      { name: 'Office Documents', extensions: ['docx', 'xlsx', 'pptx'] },
      FILTERS.docx,
      FILTERS.xlsx,
      FILTERS.pptx,
      { name: 'All Files', extensions: ['*'] },
    ],
  })

  const selected = filePaths[0]
  if (canceled || !selected) return null

  grantedPaths.add(selected)
  const document = await readDocument(selected)
  app.addRecentDocument(selected)
  return document
}

async function promptSaveAs(kind: DocumentKind, suggestedPath: string | null): Promise<string | null> {
  const owner = mainWindow
  if (!owner || owner.isDestroyed()) return null

  const { canceled, filePath } = await dialog.showSaveDialog(owner, {
    defaultPath: suggestedPath ?? `Untitled.${EXTENSIONS[kind]}`,
    filters: [FILTERS[kind]],
  })

  if (canceled || !filePath) return null
  grantedPaths.add(filePath)
  return filePath
}

function registerIpc(): void {
  ipcMain.handle(IPC.openFile, () => promptOpen())

  ipcMain.handle(IPC.readFile, (_event, filePath: string) => {
    if (!grantedPaths.has(filePath)) throw new Error(`Access denied: ${filePath}`)
    return readDocument(filePath)
  })

  ipcMain.handle(IPC.platform, () => process.platform)

  ipcMain.handle(IPC.saveFile, async (_event, request: SaveRequest): Promise<SaveResult> => {
    if (request.path && !grantedPaths.has(request.path)) {
      throw new Error(`Access denied: ${request.path}`)
    }
    const target = request.path ?? (await promptSaveAs(request.kind, null))
    if (!target) return { path: null, canceled: true }

    await writeFile(target, request.data)
    app.addRecentDocument(target)
    return { path: target, canceled: false }
  })

  ipcMain.handle(IPC.saveFileAs, async (_event, request: SaveRequest): Promise<SaveResult> => {
    const target = await promptSaveAs(request.kind, request.path)
    if (!target) return { path: null, canceled: true }

    await writeFile(target, request.data)
    app.addRecentDocument(target)
    return { path: target, canceled: false }
  })

  ipcMain.handle(IPC.confirmUnsaved, async (_event, name: string): Promise<UnsavedChoice> => {
    const owner = mainWindow
    if (!owner || owner.isDestroyed()) return 'discard'

    const { response } = await dialog.showMessageBox(owner, {
      type: 'warning',
      buttons: ['Save', "Don't Save", 'Cancel'],
      defaultId: 0,
      cancelId: 2,
      message: `Do you want to save the changes you made to ${name}?`,
      detail: "Your changes will be lost if you don't save them.",
    })
    return response === 0 ? 'save' : response === 1 ? 'discard' : 'cancel'
  })

  ipcMain.handle(IPC.exportPdf, async (_event, defaultName: string): Promise<{ path: string | null; canceled: boolean }> => {
    const owner = mainWindow
    if (!owner || owner.isDestroyed()) return { path: null, canceled: true }

    const base = defaultName.replace(/\.(docx|xlsx|pptx)$/i, '') || 'Untitled'
    const { canceled, filePath } = await dialog.showSaveDialog(owner, {
      defaultPath: `${base}.pdf`,
      filters: [{ name: 'PDF Document', extensions: ['pdf'] }],
    })
    if (canceled || !filePath) return { path: null, canceled: true }

    const data = await owner.webContents.printToPDF({
      pageSize: 'A4',
      printBackground: true,
      margins: { top: 0.4, bottom: 0.4, left: 0.4, right: 0.4 },
    })
    await writeFile(filePath, data)
    return { path: filePath, canceled: false }
  })

  ipcMain.on(IPC.rendererReady, (event) => {
    if (!mainWindow || event.sender !== mainWindow.webContents) return
    rendererReady = true
    for (const filePath of pendingOpenPaths.splice(0)) openPathInWindow(filePath)
  })

  ipcMain.on(IPC.closeResponse, (event, proceed: boolean) => {
    if (!mainWindow || event.sender !== mainWindow.webContents) return
    if (!proceed) {
      quitting = false
      return
    }
    const window = mainWindow
    mainWindow = null
    rendererReady = false
    window.destroy()
    if (quitting) app.quit()
  })
}

function openPathInWindow(filePath: string): void {
  grantedPaths.add(filePath)
  if (!mainWindow || mainWindow.isDestroyed() || !rendererReady) {
    pendingOpenPaths.push(filePath)
    return
  }
  void readDocument(filePath)
    .then((document) => {
      if (!mainWindow || mainWindow.isDestroyed()) return
      mainWindow.webContents.send(IPC.readFile, document)
      app.addRecentDocument(filePath)
    })
    .catch((error: unknown) => {
      const message = error instanceof Error ? error.message : String(error)
      dialog.showErrorBox('Could Not Open File', `${filePath}\n\n${message}`)
    })
}

app.on('open-file', (event, filePath) => {
  event.preventDefault()
  openPathInWindow(filePath)
})

app.on('before-quit', () => {
  quitting = true
})

if (!app.requestSingleInstanceLock()) {
  app.quit()
} else {
  app.on('second-instance', (_event, argv) => {
    if (mainWindow && !mainWindow.isDestroyed()) {
      if (mainWindow.isMinimized()) mainWindow.restore()
      mainWindow.focus()
    }
    for (const arg of argv.slice(1)) {
      if (/\.(docx|xlsx|pptx)$/i.test(arg)) openPathInWindow(arg)
    }
  })

  void app.whenReady().then(() => {
    registerAppProtocol()
    registerIpc()
    mainWindow = createWindow()
    buildMenu(sendMenuAction)

    app.on('activate', () => {
      if (BrowserWindow.getAllWindows().length === 0) {
        mainWindow = createWindow()
      }
    })
  })
}

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit()
})
