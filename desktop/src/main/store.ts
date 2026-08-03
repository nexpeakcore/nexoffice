import { join } from 'node:path'
import { mkdirSync, readFileSync, renameSync, writeFileSync } from 'node:fs'
import { app, type Rectangle } from 'electron'

export const MAX_RECENTS = 10

export interface WindowState {
  bounds: Rectangle
  maximized: boolean
}

interface PersistedState {
  recents: string[]
  window: WindowState | null
}

let state: PersistedState | null = null

function storePath(): string {
  return join(app.getPath('userData'), 'state.json')
}

function isRectangle(value: unknown): value is Rectangle {
  if (typeof value !== 'object' || value === null) return false
  const rect = value as Record<string, unknown>
  return (
    typeof rect['x'] === 'number' &&
    typeof rect['y'] === 'number' &&
    typeof rect['width'] === 'number' &&
    typeof rect['height'] === 'number'
  )
}

function load(): PersistedState {
  if (state) return state
  let recents: string[] = []
  let window: WindowState | null = null
  try {
    const raw = JSON.parse(readFileSync(storePath(), 'utf8')) as Record<string, unknown>
    if (Array.isArray(raw['recents'])) {
      recents = raw['recents'].filter((entry): entry is string => typeof entry === 'string').slice(0, MAX_RECENTS)
    }
    const savedWindow = raw['window'] as Record<string, unknown> | null | undefined
    if (savedWindow && isRectangle(savedWindow['bounds'])) {
      window = { bounds: savedWindow['bounds'], maximized: savedWindow['maximized'] === true }
    }
  } catch {
    // Missing or corrupt store starts fresh.
  }
  state = { recents, window }
  return state
}

function persist(): void {
  const current = load()
  const target = storePath()
  const temp = `${target}.tmp`
  try {
    mkdirSync(app.getPath('userData'), { recursive: true })
    writeFileSync(temp, JSON.stringify(current, null, 2))
    renameSync(temp, target)
  } catch (error) {
    console.error('[store] failed to persist state:', error)
  }
}

export function getRecents(): readonly string[] {
  return load().recents
}

export function addRecent(filePath: string): void {
  const current = load()
  current.recents = [filePath, ...current.recents.filter((entry) => entry !== filePath)].slice(0, MAX_RECENTS)
  persist()
}

export function removeRecent(filePath: string): void {
  const current = load()
  current.recents = current.recents.filter((entry) => entry !== filePath)
  persist()
}

export function clearRecents(): void {
  load().recents = []
  persist()
}

export function getWindowState(): WindowState | null {
  return load().window
}

export function setWindowState(windowState: WindowState): void {
  load().window = windowState
  persist()
}
