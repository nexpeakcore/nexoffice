import { app, type BrowserWindow } from 'electron'
import { autoUpdater } from 'electron-updater'
import { IPC, type UpdateEvent } from '../shared/ipc.js'

// Feed configuration (publish: generic → Azure Blob, channel `latest`) comes
// from app-update.yml, which electron-builder writes into the packaged app.
const CHECK_INTERVAL_MS = 4 * 60 * 60 * 1000
const FIRST_CHECK_DELAY_MS = 5000

type GetWindow = () => BrowserWindow | null

let sendEvent: (event: UpdateEvent) => void = () => {}
let downloading = false
let downloadedVersion: string | null = null

function makeSender(getWindow: GetWindow): (event: UpdateEvent) => void {
  return (event) => {
    const window = getWindow()
    if (!window || window.isDestroyed()) return
    window.webContents.send(IPC.updateEvent, event)
  }
}

export function setupAutoUpdater(getWindow: GetWindow): void {
  sendEvent = makeSender(getWindow)

  // Unpackaged dev runs have no app-update.yml and no installer to swap in.
  if (!app.isPackaged) {
    console.log('[updater] dev build — auto-update disabled')
    return
  }

  autoUpdater.autoDownload = true
  autoUpdater.autoInstallOnAppQuit = true

  autoUpdater.on('checking-for-update', () => sendEvent({ status: 'checking' }))

  autoUpdater.on('update-available', (info) => {
    downloading = true
    sendEvent({ status: 'available', version: info.version })
  })

  autoUpdater.on('update-not-available', () => sendEvent({ status: 'none' }))

  autoUpdater.on('download-progress', (progress) => {
    getWindow()?.setProgressBar(progress.percent / 100)
    sendEvent({
      status: 'progress',
      percent: Math.round(progress.percent),
      transferred: progress.transferred,
      total: progress.total,
      bytesPerSecond: progress.bytesPerSecond,
    })
  })

  autoUpdater.on('update-downloaded', (info) => {
    downloading = false
    downloadedVersion = info.version
    getWindow()?.setProgressBar(-1)
    sendEvent({ status: 'downloaded', version: info.version })
  })

  autoUpdater.on('error', (error) => {
    downloading = false
    getWindow()?.setProgressBar(-1)
    console.error('[updater] error:', error.message)
    sendEvent({ status: 'error', message: error.message })
  })

  const check = (): void => {
    if (downloading || downloadedVersion) return
    autoUpdater.checkForUpdates().catch((error: unknown) => {
      console.error('[updater] check failed:', error instanceof Error ? error.message : String(error))
    })
  }

  setTimeout(check, FIRST_CHECK_DELAY_MS)
  setInterval(check, CHECK_INTERVAL_MS)
}

export function checkForUpdatesManually(): void {
  if (!app.isPackaged) {
    console.log('[updater] dev build — manual check is a no-op')
    sendEvent({ status: 'dev' })
    return
  }
  if (downloadedVersion) {
    sendEvent({ status: 'downloaded', version: downloadedVersion })
    return
  }
  if (downloading) return
  autoUpdater.checkForUpdates().catch((error: unknown) => {
    sendEvent({ status: 'error', message: error instanceof Error ? error.message : String(error) })
  })
}

export function installDownloadedUpdate(): boolean {
  if (!app.isPackaged || !downloadedVersion) return false
  autoUpdater.quitAndInstall()
  return true
}
