import { basename, extname, join, resolve, sep } from 'node:path'
import { access, readFile, writeFile } from 'node:fs/promises'
import { existsSync } from 'node:fs'
import {
  app,
  BrowserWindow,
  clipboard,
  dialog,
  ipcMain,
  protocol,
  screen,
  shell,
  type Rectangle,
} from 'electron'
import { closeDecision } from './closePolicy.js'
import { buildMenu } from './menu.js'
import {
  collectProcessReport,
  documentRendererLabel,
  formatProcessReport,
  labelWebContents,
} from './processReport.js'
import {
  addRecent,
  clearRecents,
  getRecents,
  getStoredLocale,
  getWindowState,
  removeRecent,
  setStoredLocale,
  setWindowState,
} from './store.js'
import {
  createTranslator,
  DEFAULT_LOCALE,
  isSupportedLocale,
  matchLocale,
  type LocaleCode,
  type Translator,
} from '../i18n/index.js'
import { checkForUpdatesManually, installDownloadedUpdate, setupAutoUpdater } from './updater.js'
import { registerAgent } from './agent/index.js'
import {
  ALL_EDIT_CAPABILITIES,
  EXTENSIONS,
  IPC,
  kindFromPath,
  PRINT_PAGE_CAP,
  readEditCapabilities,
  readRendererDiagnostics,
  sameEditCapabilities,
  type DocumentKind,
  type EditCapabilities,
  type ExportPdfRequest,
  type ExportPdfResult,
  type MenuAction,
  type OpenedDocument,
  type PrintJob,
  type PrintResult,
  type PrintRenderResult,
  type RefusedChoice,
  type RendererDiagnostics,
  type SaveRequest,
  type SaveResult,
  type UnsavedChoice,
  type WebEditAction,
  type RecentFile,
} from '../shared/ipc.js'

const isDev = !app.isPackaged
let mainWindow: BrowserWindow | null = null
let rendererReady = false
let quitting = false
/** Whether a close is already waiting on the renderer's answer. */
let closeRequested = false
/** Whether the renderer has stopped answering input, per Electron. */
let rendererUnresponsive = false
let activeDocumentKind: DocumentKind | null = null
/** The renderer's last account of its own memory, and when it landed. */
let rendererDiagnostics: RendererDiagnostics | null = null
let diagnosticsAt = 0
let activeEditCapabilities: EditCapabilities = ALL_EDIT_CAPABILITIES
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

let activeLocale: LocaleCode = DEFAULT_LOCALE
let t: Translator = createTranslator(activeLocale)

function applyLocale(locale: LocaleCode): void {
  activeLocale = locale
  t = createTranslator(locale)
}

function fileFilters(): Record<DocumentKind, Electron.FileFilter> {
  return {
    docx: { name: t('dialog.filters.wordDocument'), extensions: ['docx'] },
    xlsx: { name: t('dialog.filters.excelWorkbook'), extensions: ['xlsx'] },
    pptx: { name: t('dialog.filters.powerpointPresentation'), extensions: ['pptx'] },
  }
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

const MIN_WIDTH = 960
const MIN_HEIGHT = 600
const COPYRIGHT = 'Copyright © 2026 NexOffice. Licensed under Apache-2.0.'

function restoredWindowState(): { bounds: Rectangle | null; maximized: boolean } {
  const saved = getWindowState()
  if (!saved) return { bounds: null, maximized: false }
  const area = screen.getDisplayMatching(saved.bounds).workArea
  const width = Math.min(Math.max(saved.bounds.width, MIN_WIDTH), area.width)
  const height = Math.min(Math.max(saved.bounds.height, MIN_HEIGHT), area.height)
  const x = Math.min(Math.max(saved.bounds.x, area.x), area.x + area.width - width)
  const y = Math.min(Math.max(saved.bounds.y, area.y), area.y + area.height - height)
  return { bounds: { x, y, width, height }, maximized: saved.maximized }
}

function trackWindowState(window: BrowserWindow): void {
  const persist = (): void => {
    if (window.isDestroyed()) return
    setWindowState({ bounds: window.getNormalBounds(), maximized: window.isMaximized() })
  }
  let timer: ReturnType<typeof setTimeout> | null = null
  const schedulePersist = (): void => {
    if (timer) clearTimeout(timer)
    timer = setTimeout(persist, 500)
  }
  window.on('resize', schedulePersist)
  window.on('move', schedulePersist)
  window.on('maximize', persist)
  window.on('unmaximize', persist)
  window.on('close', () => {
    if (timer) clearTimeout(timer)
    persist()
  })
}

function createWindow(): BrowserWindow {
  rendererReady = false
  // A fresh window holds none of the previous one's memory, so its numbers
  // must not survive into the next report as though they still did.
  rendererDiagnostics = null
  diagnosticsAt = 0

  const { bounds, maximized } = restoredWindowState()

  const window = new BrowserWindow({
    width: bounds?.width ?? 1440,
    height: bounds?.height ?? 900,
    ...(bounds ? { x: bounds.x, y: bounds.y } : {}),
    minWidth: MIN_WIDTH,
    minHeight: MIN_HEIGHT,
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

  labelWebContents(window.webContents, documentRendererLabel(null))

  window.once('ready-to-show', () => {
    if (maximized) window.maximize()
    window.show()
  })

  trackWindowState(window)

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

  // Vetoing the close is how the renderer gets to ask about unsaved work, and
  // it is also how the app becomes impossible to quit if the renderer never
  // answers. The wait ends when the renderer has demonstrably stopped
  // answering — not when the user asks a second time, and not on a timer.
  //
  // Both of the cheaper rules lose documents. A second attempt cannot be told
  // apart from impatience: a save takes seconds on a large deck, and a user who
  // presses again during one would have the window destroyed mid-write. A
  // timer cannot be told apart from a prompt sitting on screen waiting to be
  // read. `unresponsive` is the one signal that means the renderer itself has
  // stopped, so it is the only one allowed to end the wait.
  window.on('close', (event) => {
    const decision = closeDecision({
      rendererReady,
      webContentsDestroyed: window.webContents.isDestroyed(),
      closeRequested,
      rendererUnresponsive,
    })
    if (decision === 'close') return
    event.preventDefault()
    if (decision === 'wait') return
    closeRequested = true
    window.webContents.send(IPC.closeRequest)
  })

  // Renderer console into the dev server's terminal. Diagnosing the editor
  // otherwise means asking whoever is testing to open devtools and read it
  // back, which loses everything that happened before they thought to look.
  if (isDev) {
    window.webContents.on('console-message', (_event, level, message, line, sourceId) => {
      if (level < 2) return
      const where = sourceId ? ` (${sourceId}:${line})` : ''
      console[level === 2 ? 'warn' : 'error'](`[renderer] ${message}${where}`)
    })
  }

  window.webContents.on('unresponsive', () => {
    rendererUnresponsive = true
  })
  window.webContents.on('responsive', () => {
    rendererUnresponsive = false
  })

  // A dead renderer will never answer the handshake, so the window stops
  // waiting on one. It is not destroyed here: quitting the app out from under
  // the user with no word of why is its own failure, and on Windows and Linux
  // destroying the last window quits. Marking it not-ready is enough — the
  // next close goes straight through.
  window.webContents.on('render-process-gone', (_event, details) => {
    if (details.reason === 'clean-exit') return
    rendererReady = false
    rendererUnresponsive = false
    closeRequested = false
  })

  window.on('closed', () => {
    if (mainWindow === window) {
      mainWindow = null
      rendererReady = false
      closeRequested = false
      rendererUnresponsive = false
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
  const started = Date.now()
  const data = await readFile(filePath)
  return {
    path: filePath,
    name: basename(filePath),
    kind: kindFromPath(filePath),
    data: new Uint8Array(data),
    readMs: Date.now() - started,
    sentAt: Date.now(),
  }
}

function sendMenuAction(action: MenuAction): void {
  if (!mainWindow || mainWindow.isDestroyed()) return
  mainWindow.webContents.send(IPC.menuAction, action)
}

function rebuildMenu(): void {
  buildMenu({
    t,
    locale: activeLocale,
    dispatch: sendMenuAction,
    documentKind: activeDocumentKind,
    editCapabilities: activeEditCapabilities,
    recents: getRecents(),
    onOpenRecent: (filePath) => void openRecentPath(filePath),
    onClearRecents: () => {
      clearRecents()
      app.clearRecentDocuments()
      rebuildMenu()
    },
    onSelectLocale: selectLocale,
    onCheckForUpdates: () => checkForUpdatesManually(),
    onShowAbout: showAboutDialog,
    onShowProcesses: showProcessReportDialog,
  })
}

function selectLocale(locale: LocaleCode): void {
  if (locale === activeLocale) return
  applyLocale(locale)
  setStoredLocale(locale)
  rebuildMenu()
  if (mainWindow && !mainWindow.isDestroyed()) {
    mainWindow.webContents.send(IPC.localeChanged, locale)
  }
}

function noteRecent(filePath: string): void {
  addRecent(filePath)
  app.addRecentDocument(filePath)
  rebuildMenu()
}

function listRecentFiles(): RecentFile[] {
  return getRecents().map((filePath) => {
    const name = basename(filePath)
    const extension = name.split('.').pop()?.toLowerCase()
    const kind: DocumentKind | null =
      extension === 'docx' || extension === 'xlsx' || extension === 'pptx' ? extension : null
    return { path: filePath, name, kind, exists: existsSync(filePath) }
  })
}

async function openRecentPath(filePath: string): Promise<void> {
  try {
    await access(filePath)
  } catch {
    removeRecent(filePath)
    rebuildMenu()
    const owner = mainWindow && !mainWindow.isDestroyed() ? mainWindow : null
    const options: Electron.MessageBoxOptions = {
      type: 'info',
      message: t('dialog.fileNotFound.message'),
      detail: t('dialog.fileNotFound.detail', { path: filePath }),
    }
    void (owner ? dialog.showMessageBox(owner, options) : dialog.showMessageBox(options))
    return
  }
  openPathInWindow(filePath)
}

function showAboutDialog(): void {
  const options: Electron.MessageBoxOptions = {
    type: 'info',
    title: t('dialog.about.title'),
    message: 'NexOffice',
    detail: `${t('dialog.about.version', { version: app.getVersion() })}\nElectron ${process.versions.electron}\nChromium ${process.versions.chrome}\n\n${COPYRIGHT}`,
  }
  const owner = mainWindow && !mainWindow.isDestroyed() ? mainWindow : null
  void (owner ? dialog.showMessageBox(owner, options) : dialog.showMessageBox(options))
}

function requestDiagnosticsSample(): void {
  if (!mainWindow || mainWindow.isDestroyed() || !rendererReady) return
  mainWindow.webContents.send(IPC.diagnosticsSample)
}

function showProcessReportDialog(): void {
  // The renderer takes a fresh sample on the way in, so the Refresh button —
  // not this first render — is what shows current numbers. Asking and then
  // waiting would hang the dialog exactly when it is most wanted: while the
  // renderer is busy laying out the heavy document being diagnosed.
  requestDiagnosticsSample()
  const report = formatProcessReport(collectProcessReport(), {
    diagnostics: rendererDiagnostics,
    sampleAgeMs: rendererDiagnostics ? Date.now() - diagnosticsAt : null,
  })
  const owner = mainWindow && !mainWindow.isDestroyed() ? mainWindow : null
  const options: Electron.MessageBoxOptions = {
    type: 'info',
    title: t('dialog.processes.title'),
    message: t('dialog.processes.message'),
    detail: report,
    buttons: [t('dialog.processes.refresh'), t('dialog.processes.copy'), t('dialog.processes.close')],
    defaultId: 2,
    cancelId: 2,
    noLink: true,
  }
  const show = owner ? dialog.showMessageBox(owner, options) : dialog.showMessageBox(options)
  void show.then(({ response }) => {
    if (response === 0) showProcessReportDialog()
    else if (response === 1) clipboard.writeText(report)
  })
}

async function promptOpen(): Promise<OpenedDocument | null> {
  const owner = mainWindow
  if (!owner || owner.isDestroyed()) return null

  const filters = fileFilters()
  const { canceled, filePaths } = await dialog.showOpenDialog(owner, {
    properties: ['openFile'],
    filters: [
      { name: t('dialog.filters.officeDocuments'), extensions: ['docx', 'xlsx', 'pptx'] },
      filters.docx,
      filters.xlsx,
      filters.pptx,
      { name: t('dialog.filters.allFiles'), extensions: ['*'] },
    ],
  })

  const selected = filePaths[0]
  if (canceled || !selected) return null

  grantedPaths.add(selected)
  const document = await readDocument(selected)
  noteRecent(selected)
  return document
}

async function promptSaveAs(kind: DocumentKind, suggestedPath: string | null): Promise<string | null> {
  const owner = mainWindow
  if (!owner || owner.isDestroyed()) return null

  const { canceled, filePath } = await dialog.showSaveDialog(owner, {
    defaultPath: suggestedPath ?? `Untitled.${EXTENSIONS[kind]}`,
    filters: [fileFilters()[kind]],
  })

  if (canceled || !filePath) return null
  grantedPaths.add(filePath)
  return filePath
}

const PRINT_QUERY = 'offscreenReplay=0&pageWindow=0'
const PRINT_TIMEOUT_MS = 120_000

// A deck can fail on more pages than a message box can carry, so the notice
// names the first few and the main-process log keeps the rest.
const MAX_LISTED_PAGES = 12

function formatPageList(numbers: number[]): string {
  if (numbers.length <= MAX_LISTED_PAGES) return numbers.join(', ')
  return `${numbers.slice(0, MAX_LISTED_PAGES).join(', ')}…`
}

function createPrintWindow(): BrowserWindow {
  const window = new BrowserWindow({
    width: 1000,
    height: 1400,
    show: false,
    webPreferences: {
      preload: join(import.meta.dirname, '../preload/index.cjs'),
      sandbox: true,
      contextIsolation: true,
      nodeIntegration: false,
    },
  })

  labelWebContents(window.webContents, 'Renderer · Print/PDF (transient)')

  window.webContents.setWindowOpenHandler(() => ({ action: 'deny' }))
  window.webContents.on('will-navigate', (event) => event.preventDefault())

  const devServerUrl = process.env['ELECTRON_RENDERER_URL']
  if (isDev && devServerUrl) {
    void window.loadURL(`${devServerUrl}/print.html?${PRINT_QUERY}`)
  } else {
    void window.loadURL(`${APP_ORIGIN}/print.html?${PRINT_QUERY}`)
  }
  return window
}

function renderPrintJob(
  window: BrowserWindow,
  job: PrintJob,
): Promise<{ pages: number; truncated: boolean; skippedPages: number[] }> {
  return new Promise((resolvePromise, rejectPromise) => {
    const cleanup = (): void => {
      clearTimeout(timer)
      ipcMain.removeListener(IPC.printReady, onReady)
      ipcMain.removeListener(IPC.printRendered, onRendered)
      window.removeListener('closed', onClosed)
    }
    const fail = (message: string): void => {
      cleanup()
      rejectPromise(new Error(message))
    }
    const timer = setTimeout(() => fail('PDF rendering timed out'), PRINT_TIMEOUT_MS)
    const onClosed = (): void => fail('Print window closed before rendering finished')
    const onReady = (event: Electron.IpcMainEvent): void => {
      if (window.isDestroyed() || event.sender !== window.webContents) return
      event.sender.send(IPC.printRender, job)
    }
    const onRendered = (event: Electron.IpcMainEvent, result: PrintRenderResult): void => {
      if (window.isDestroyed() || event.sender !== window.webContents) return
      if (result.ok) {
        cleanup()
        resolvePromise({
          pages: result.pages,
          truncated: result.truncated,
          skippedPages: result.skippedPages,
        })
      } else {
        fail(result.error)
      }
    }
    ipcMain.on(IPC.printReady, onReady)
    ipcMain.on(IPC.printRendered, onRendered)
    window.on('closed', onClosed)
    window.webContents.on('render-process-gone', (_event, details) =>
      fail(`Print renderer crashed: ${details.reason}`),
    )
  })
}

function registerIpc(): void {
  ipcMain.handle(IPC.openFile, () => promptOpen())

  ipcMain.handle(IPC.readFile, (_event, filePath: string) => {
    if (!grantedPaths.has(filePath)) throw new Error(`Access denied: ${filePath}`)
    return readDocument(filePath)
  })

  ipcMain.handle(IPC.newDocument, async (_event, kind: DocumentKind): Promise<OpenedDocument> => {
    if (kind !== 'docx' && kind !== 'xlsx' && kind !== 'pptx') {
      throw new Error(`Unknown document kind: ${String(kind)}`)
    }
    const templates = app.isPackaged
      ? join(process.resourcesPath, 'templates')
      : join(app.getAppPath(), 'resources', 'templates')
    const data = await readFile(join(templates, `blank.${kind}`))
    return { path: '', name: `Untitled.${kind}`, kind, data: new Uint8Array(data) }
  })

  ipcMain.handle(IPC.recentsList, (): RecentFile[] => listRecentFiles())

  ipcMain.handle(IPC.recentsRemove, (_event, filePath: string): RecentFile[] => {
    removeRecent(filePath)
    rebuildMenu()
    return listRecentFiles()
  })

  ipcMain.handle(IPC.openRecent, async (_event, filePath: string) => {
    if (!getRecents().includes(filePath)) throw new Error(`Not a recent file: ${filePath}`)
    await openRecentPath(filePath)
  })

  ipcMain.handle(IPC.platform, () => process.platform)
  ipcMain.handle(IPC.version, () => app.getVersion())

  ipcMain.handle(IPC.locale, () => activeLocale)

  ipcMain.handle(IPC.saveFile, async (_event, request: SaveRequest): Promise<SaveResult> => {
    if (request.path && !grantedPaths.has(request.path)) {
      throw new Error(`Access denied: ${request.path}`)
    }
    const target = request.path ?? (await promptSaveAs(request.kind, null))
    if (!target) return { path: null, canceled: true }

    await writeFile(target, request.data)
    noteRecent(target)
    return { path: target, canceled: false }
  })

  ipcMain.handle(IPC.saveFileAs, async (_event, request: SaveRequest): Promise<SaveResult> => {
    const target = await promptSaveAs(request.kind, request.path)
    if (!target) return { path: null, canceled: true }

    await writeFile(target, request.data)
    noteRecent(target)
    return { path: target, canceled: false }
  })

  ipcMain.handle(IPC.confirmUnsaved, async (_event, name: string): Promise<UnsavedChoice> => {
    const owner = mainWindow
    if (!owner || owner.isDestroyed()) return 'discard'

    const { response } = await dialog.showMessageBox(owner, {
      type: 'warning',
      buttons: [t('dialog.unsaved.save'), t('dialog.unsaved.dontSave'), t('dialog.unsaved.cancel')],
      defaultId: 0,
      cancelId: 2,
      message: t('dialog.unsaved.message', { name }),
      detail: t('dialog.unsaved.detail'),
    })
    return response === 0 ? 'save' : response === 1 ? 'discard' : 'cancel'
  })

  // Shown only after a save the writer refused, so the buttons are the two
  // moves that exist: leave the change behind, or keep the document open. The
  // refusal names what could not be written and is quoted whole — it is the
  // only thing that tells the user which change to take back.
  //
  // Keeping the document open is `buttons[0]`. macOS lays buttons out from
  // there rightwards, which is where the unsaved-changes prompt just put
  // "Save" — so the spot the hand is already moving to keeps the work rather
  // than throwing it away.
  ipcMain.handle(
    IPC.confirmSaveRefused,
    async (_event, request: { name: string; message: string }): Promise<RefusedChoice> => {
      const owner = mainWindow
      // Matches `confirmUnsaved`: with no window there is no document left to
      // keep open, and answering 'cancel' here would veto a quit the user
      // asked for without ever telling them why.
      if (!owner || owner.isDestroyed()) return 'discard'

      const { response } = await dialog.showMessageBox(owner, {
        type: 'warning',
        buttons: [t('dialog.saveRefused.keepEditing'), t('dialog.saveRefused.discard')],
        defaultId: 0,
        cancelId: 0,
        message: t('dialog.saveRefused.message', { name: request.name }),
        detail: t('dialog.saveRefused.detail', { message: request.message }),
      })
      return response === 1 ? 'discard' : 'cancel'
    },
  )

  ipcMain.handle(IPC.exportPdf, async (_event, request: ExportPdfRequest): Promise<ExportPdfResult> => {
    const owner = mainWindow
    if (!owner || owner.isDestroyed()) return { path: null, canceled: true }

    const base = request.name.replace(/\.(docx|xlsx|pptx)$/i, '') || 'Untitled'
    const { canceled, filePath } = await dialog.showSaveDialog(owner, {
      defaultPath: `${base}.pdf`,
      filters: [{ name: t('dialog.filters.pdfDocument'), extensions: ['pdf'] }],
    })
    if (canceled || !filePath) return { path: null, canceled: true }
    grantedPaths.add(filePath)

    const printWindow = createPrintWindow()
    try {
      const { pages, truncated, skippedPages } = await renderPrintJob(printWindow, {
        kind: request.kind,
        data: request.data,
      })
      const pdf = await printWindow.webContents.printToPDF({
        preferCSSPageSize: true,
        printBackground: true,
        margins: { top: 0, bottom: 0, left: 0, right: 0 },
      })
      await writeFile(filePath, pdf)
      if (truncated && !owner.isDestroyed()) {
        void dialog.showMessageBox(owner, {
          type: 'warning',
          message: t('dialog.pdfTruncated.message'),
          detail: t('dialog.pdfTruncated.detail', { cap: PRINT_PAGE_CAP, pages }),
        })
      }
      if (skippedPages.length > 0) {
        // The only other record of which slides failed is a console warning
        // inside the hidden print window, which nothing outlives.
        console.warn(
          `PDF export left out ${skippedPages.length} unrenderable page(s) of ${basename(filePath)}: ${skippedPages.join(', ')}`,
        )
        if (!owner.isDestroyed()) {
          const slides = formatPageList(skippedPages)
          void dialog.showMessageBox(owner, {
            type: 'warning',
            message: t('dialog.pdfSkipped.message'),
            detail:
              skippedPages.length === 1
                ? t('dialog.pdfSkipped.detailOne', { slides, pages })
                : t('dialog.pdfSkipped.detailMany', {
                    count: skippedPages.length,
                    slides,
                    pages,
                  }),
          })
        }
      }
      return { path: filePath, canceled: false, pages, truncated, skipped: skippedPages.length }
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error)
      dialog.showErrorBox(t('dialog.pdfFailed.title'), message)
      return { path: null, canceled: true }
    } finally {
      if (!printWindow.isDestroyed()) printWindow.destroy()
    }
  })

  ipcMain.handle(IPC.printDocument, async (_event, request: ExportPdfRequest): Promise<PrintResult> => {
    const owner = mainWindow
    if (!owner || owner.isDestroyed()) return { printed: false }

    const printWindow = createPrintWindow()
    try {
      const { pages, truncated, skippedPages } = await renderPrintJob(printWindow, {
        kind: request.kind,
        data: request.data,
      })
      const { printed, failureReason } = await new Promise<{
        printed: boolean
        failureReason: string
      }>((resolve) => {
        printWindow.webContents.print({ printBackground: true }, (success, reason) =>
          resolve({ printed: success, failureReason: reason }),
        )
      })
      if (!printed && failureReason && failureReason !== 'cancelled' && !owner.isDestroyed()) {
        dialog.showErrorBox(t('dialog.printFailed.title'), failureReason)
      }
      if (printed && truncated && !owner.isDestroyed()) {
        void dialog.showMessageBox(owner, {
          type: 'warning',
          message: t('dialog.pdfTruncated.message'),
          detail: t('dialog.pdfTruncated.detail', { cap: PRINT_PAGE_CAP, pages }),
        })
      }
      if (printed && skippedPages.length > 0 && !owner.isDestroyed()) {
        const slides = formatPageList(skippedPages)
        void dialog.showMessageBox(owner, {
          type: 'warning',
          message: t('dialog.pdfSkipped.message'),
          detail:
            skippedPages.length === 1
              ? t('dialog.pdfSkipped.detailOne', { slides, pages })
              : t('dialog.pdfSkipped.detailMany', {
                  count: skippedPages.length,
                  slides,
                  pages,
                }),
        })
      }
      return {
        printed,
        pages,
        truncated,
        skipped: skippedPages.length,
      }
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error)
      dialog.showErrorBox(t('dialog.printFailed.title'), message)
      return { printed: false }
    } finally {
      if (!printWindow.isDestroyed()) printWindow.destroy()
    }
  })

  ipcMain.on(IPC.webEditAction, (event, action: WebEditAction) => {
    if (!mainWindow || event.sender !== mainWindow.webContents) return
    const contents = mainWindow.webContents
    if (action === 'undo') contents.undo()
    else if (action === 'redo') contents.redo()
    else if (action === 'cut') contents.cut()
    else if (action === 'copy') contents.copy()
    else if (action === 'delete') contents.delete()
    else if (action === 'selectAll') contents.selectAll()
    else contents.paste()
  })

  ipcMain.on(IPC.diagnostics, (event, payload: unknown) => {
    if (!mainWindow || event.sender !== mainWindow.webContents) return
    const next = readRendererDiagnostics(payload)
    if (!next) return
    rendererDiagnostics = next
    diagnosticsAt = Date.now()
    labelWebContents(mainWindow.webContents, documentRendererLabel(next.document))
  })

  ipcMain.on(IPC.documentKind, (event, kind: DocumentKind | null) => {
    if (!mainWindow || event.sender !== mainWindow.webContents) return
    if (kind !== null && kind !== 'docx' && kind !== 'xlsx' && kind !== 'pptx') return
    if (kind === activeDocumentKind) return
    activeDocumentKind = kind
    // Only the pptx editor reports narrower capabilities, so leaving its last
    // report in place would disable the Edit menu over the next document.
    if (kind !== 'pptx') activeEditCapabilities = ALL_EDIT_CAPABILITIES
    rebuildMenu()
  })

  // The renderer only sends this when a capability actually flips, but the
  // menu is rebuilt from scratch here, so the comparison is repeated rather
  // than trusted: a caret moving inside a paragraph must not rebuild a menu.
  ipcMain.on(IPC.editCapabilities, (event, capabilities: EditCapabilities) => {
    if (!mainWindow || event.sender !== mainWindow.webContents) return
    const next = readEditCapabilities(capabilities)
    if (!next || sameEditCapabilities(next, activeEditCapabilities)) return
    activeEditCapabilities = next
    rebuildMenu()
  })

  ipcMain.on(IPC.rendererReady, (event) => {
    if (!mainWindow || event.sender !== mainWindow.webContents) return
    rendererReady = true
    for (const filePath of pendingOpenPaths.splice(0)) openPathInWindow(filePath)
  })

  ipcMain.handle(IPC.updateCheck, (event) => {
    if (!mainWindow || event.sender !== mainWindow.webContents) return
    checkForUpdatesManually()
  })

  ipcMain.on(IPC.updateInstall, (event) => {
    if (!mainWindow || event.sender !== mainWindow.webContents) return
    quitting = true
    rendererReady = false
    if (!installDownloadedUpdate()) {
      quitting = false
      rendererReady = true
    }
  })

  ipcMain.on(IPC.closeResponse, (event, proceed: boolean) => {
    if (!mainWindow || event.sender !== mainWindow.webContents) return
    closeRequested = false
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

function documentPathsFromArgv(argv: string[], cwd: string): string[] {
  return argv
    .slice(1)
    .filter((arg) => !arg.startsWith('-') && /\.(docx|xlsx|pptx)$/i.test(arg))
    .map((arg) => resolve(cwd, arg))
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
      noteRecent(filePath)
    })
    .catch((error: unknown) => {
      const message = error instanceof Error ? error.message : String(error)
      dialog.showErrorBox(t('dialog.openFailed.title'), `${filePath}\n\n${message}`)
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
  app.on('second-instance', (_event, argv, workingDirectory) => {
    if (mainWindow && !mainWindow.isDestroyed()) {
      if (mainWindow.isMinimized()) mainWindow.restore()
      mainWindow.focus()
    }
    for (const filePath of documentPathsFromArgv(argv, workingDirectory)) {
      openPathInWindow(filePath)
    }
  })

  void app.whenReady().then(() => {
    const stored = getStoredLocale()
    applyLocale(stored && isSupportedLocale(stored) ? stored : matchLocale(app.getLocale()))
    if (process.platform === 'darwin') {
      app.setAboutPanelOptions({
        applicationName: 'NexOffice',
        applicationVersion: app.getVersion(),
        version: `Electron ${process.versions.electron} · Chromium ${process.versions.chrome}`,
        copyright: COPYRIGHT,
      })
    }
    registerAppProtocol()
    registerIpc()
    mainWindow = createWindow()
    rebuildMenu()
    setupAutoUpdater(() => mainWindow)
    registerAgent(() => mainWindow)

    for (const filePath of documentPathsFromArgv(process.argv, process.cwd())) {
      openPathInWindow(filePath)
    }

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
