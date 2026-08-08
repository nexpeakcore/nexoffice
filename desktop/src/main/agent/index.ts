import { randomUUID } from 'node:crypto'
import { join } from 'node:path'
import { mkdirSync, readFileSync, writeFileSync, rmSync } from 'node:fs'
import { app, ipcMain, safeStorage, type BrowserWindow } from 'electron'
import {
  AGENT_DEFAULT_MODELS,
  IPC,
  type AgentEvent,
  type AgentProvider,
  type AgentRunRequest,
  type AgentSettings,
  type AgentSettingsUpdate,
  type AgentToolRequest,
  type AgentToolResult,
} from '../../shared/ipc.js'
import { runAgentLoop } from './loop.js'
import { streamChatTurn, type ChatMessage } from './openaiStream.js'
import { agentSystemPrompt, APPROVAL_TOOLS, toolsForDocument } from './tools.js'

const PROVIDER_BASE_URLS: Record<AgentProvider, string> = {
  deepseek: 'https://api.deepseek.com',
}

function providerBaseUrl(provider: AgentProvider): string {
  return process.env['NEXOFFICE_AGENT_BASE_URL'] ?? PROVIDER_BASE_URLS[provider]
}

const TOOL_TIMEOUT_MS = 15_000
// Approval-gated tools wait on a human decision, not on computation.
const APPROVAL_TIMEOUT_MS = 5 * 60_000

interface AgentConfig {
  provider: AgentProvider
  model: string
}

function configPath(): string {
  return join(app.getPath('userData'), 'agent.json')
}

function keyPath(): string {
  return join(app.getPath('userData'), 'agent.key')
}

function loadConfig(): AgentConfig {
  try {
    const raw = JSON.parse(readFileSync(configPath(), 'utf8')) as Record<string, unknown>
    const provider: AgentProvider = raw['provider'] === 'deepseek' ? 'deepseek' : 'deepseek'
    const model =
      typeof raw['model'] === 'string' && raw['model'].trim()
        ? raw['model']
        : AGENT_DEFAULT_MODELS[provider]
    return { provider, model }
  } catch {
    return { provider: 'deepseek', model: AGENT_DEFAULT_MODELS.deepseek }
  }
}

function saveConfig(config: AgentConfig): void {
  mkdirSync(app.getPath('userData'), { recursive: true })
  writeFileSync(configPath(), JSON.stringify(config, null, 2))
}

// The key never crosses to the renderer and never lands on disk in the clear:
// encrypted with the OS keychain-backed safeStorage when available, refused
// otherwise (a plaintext fallback would silently downgrade every install).
function saveApiKey(key: string): void {
  if (!safeStorage.isEncryptionAvailable()) {
    throw new Error('secure key storage is unavailable on this system')
  }
  mkdirSync(app.getPath('userData'), { recursive: true })
  writeFileSync(keyPath(), safeStorage.encryptString(key))
}

function loadApiKey(): string | null {
  try {
    const encrypted = readFileSync(keyPath())
    if (!safeStorage.isEncryptionAvailable()) return null
    return safeStorage.decryptString(encrypted)
  } catch {
    return null
  }
}

function clearApiKey(): void {
  rmSync(keyPath(), { force: true })
}

export function registerAgent(getWindow: () => BrowserWindow | null): void {
  let abort: AbortController | null = null
  const pendingTools = new Map<
    string,
    { resolve: (value: unknown) => void; timer: NodeJS.Timeout }
  >()

  const fromMainWindow = (event: Electron.IpcMainEvent | Electron.IpcMainInvokeEvent): boolean => {
    const window = getWindow()
    return window !== null && !window.isDestroyed() && event.sender === window.webContents
  }

  ipcMain.handle(IPC.agentSettingsGet, (event): AgentSettings | null => {
    if (!fromMainWindow(event)) return null
    const config = loadConfig()
    return { ...config, hasApiKey: loadApiKey() !== null }
  })

  ipcMain.handle(IPC.agentSettingsSet, (event, update: AgentSettingsUpdate): AgentSettings | null => {
    if (!fromMainWindow(event)) return null
    const config = loadConfig()
    if (update.provider) config.provider = update.provider
    if (typeof update.model === 'string' && update.model.trim()) config.model = update.model.trim()
    saveConfig(config)
    if (typeof update.apiKey === 'string') {
      if (update.apiKey === '') clearApiKey()
      else saveApiKey(update.apiKey)
    }
    return { ...config, hasApiKey: loadApiKey() !== null }
  })

  ipcMain.on(IPC.agentToolResult, (event, payload: AgentToolResult) => {
    if (!fromMainWindow(event)) return
    const pending = pendingTools.get(payload.id)
    if (!pending) return
    pendingTools.delete(payload.id)
    clearTimeout(pending.timer)
    pending.resolve(payload.result)
  })

  ipcMain.on(IPC.agentCancel, (event) => {
    if (!fromMainWindow(event)) return
    abort?.abort()
  })

  ipcMain.handle(IPC.agentRun, async (event, request: AgentRunRequest): Promise<void> => {
    if (!fromMainWindow(event)) return
    const window = getWindow()
    if (!window) return
    const emit = (agentEvent: AgentEvent): void => {
      if (!window.isDestroyed()) window.webContents.send(IPC.agentEvent, agentEvent)
    }

    const apiKey = loadApiKey()
    if (!apiKey) {
      emit({ type: 'error', message: 'missing-api-key' })
      return
    }

    abort?.abort()
    const controller = new AbortController()
    abort = controller
    const config = loadConfig()

    const executeTool = (name: string, args: Record<string, unknown>): Promise<unknown> =>
      new Promise((resolve) => {
        const id = randomUUID()
        const timer = setTimeout(
          () => {
            pendingTools.delete(id)
            resolve({ error: 'tool call timed out' })
          },
          APPROVAL_TOOLS.has(name) ? APPROVAL_TIMEOUT_MS : TOOL_TIMEOUT_MS
        )
        pendingTools.set(id, { resolve, timer })
        const toolRequest: AgentToolRequest = { id, name, args }
        if (window.isDestroyed()) {
          clearTimeout(timer)
          pendingTools.delete(id)
          resolve({ error: 'window closed' })
          return
        }
        window.webContents.send(IPC.agentToolRequest, toolRequest)
      })

    const transcript: ChatMessage[] = [
      {
        role: 'system',
        content: agentSystemPrompt(request.documentKind, request.documentName, request.locale),
      },
      ...request.messages.map(
        (message): ChatMessage => ({ role: message.role, content: message.content })
      ),
    ]

    try {
      await runAgentLoop(transcript, {
        tools: toolsForDocument(request.documentKind),
        callModel: (messages, tools, onTextDelta) =>
          streamChatTurn(
            {
              baseUrl: providerBaseUrl(config.provider),
              apiKey,
              model: config.model,
            },
            messages,
            tools,
            { onTextDelta },
            controller.signal
          ),
        executeTool,
        emit,
      })
    } catch (error) {
      if (!controller.signal.aborted) {
        emit({ type: 'error', message: error instanceof Error ? error.message : String(error) })
      } else {
        emit({ type: 'done' })
      }
    } finally {
      if (abort === controller) abort = null
    }
  })
}
