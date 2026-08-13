import type { Misspelling, SpellCheckEngine } from './spellcheckEngine.js'
import type {
  SpellCheckRequest,
  SpellCheckRequestWithoutId,
  SpellCheckResponse,
} from './spellcheckProtocol.js'

export type { Misspelling }

type Settled = SpellCheckResponse & { ok: true }

interface PendingRequest {
  resolve: (response: Settled) => void
  reject: (error: Error) => void
}

/**
 * Renderer-side proofing client. The dictionary itself lives in a dedicated
 * worker: expanding en_US costs ~40MB of heap and blocks whichever thread does
 * it, neither of which belongs on the thread that paints the editor. If the
 * worker cannot start, the engine is loaded inline so proofing still works.
 */
class SpellCheckServiceImpl {
  private worker: Worker | null = null
  private inline: SpellCheckEngine | null = null
  private starting: Promise<void> | null = null
  private pending = new Map<number, PendingRequest>()
  private nextId = 1
  ready = false

  init(_lang: string): Promise<void> {
    if (this.ready) return Promise.resolve()
    this.starting ??= this.start().catch((error: unknown) => {
      this.starting = null
      throw error instanceof Error ? error : new Error(String(error))
    })
    return this.starting
  }

  async check(text: string): Promise<Misspelling[]> {
    if (!this.ready) return []
    if (this.inline) return this.inline.check(text)
    return (await this.request({ type: 'check', text })).misspellings ?? []
  }

  async suggest(word: string): Promise<string[]> {
    if (!this.ready) return []
    if (this.inline) return this.inline.suggest(word)
    return (await this.request({ type: 'suggest', word })).suggestions ?? []
  }

  private async start(): Promise<void> {
    try {
      const worker = new Worker(new URL('./spellcheckWorker.ts', import.meta.url), {
        type: 'module',
        name: 'nexoffice-spellcheck',
      })
      worker.onmessage = (event: MessageEvent<SpellCheckResponse>) => {
        const response = event.data
        const pending = this.pending.get(response.id)
        if (!pending) return
        this.pending.delete(response.id)
        if (response.ok) pending.resolve(response)
        else pending.reject(new Error(response.error))
      }
      worker.onerror = (event) => {
        this.failAll(new Error(event.message || 'Spell-check worker failed'))
      }
      this.worker = worker
      await this.request({ type: 'init' })
    } catch {
      // A renderer that cannot host the worker still gets proofing; it just
      // pays the dictionary's heap on this thread.
      this.worker?.terminate()
      this.worker = null
      this.failAll(new Error('Spell-check worker unavailable'))
      const { SpellCheckEngine } = await import('./spellcheckEngine.js')
      const engine = new SpellCheckEngine()
      await engine.init()
      this.inline = engine
    }
    this.ready = true
  }

  private request(request: SpellCheckRequestWithoutId): Promise<Settled> {
    const worker = this.worker
    if (!worker) return Promise.reject(new Error('Spell-check worker is not running'))
    const id = this.nextId++
    return new Promise<Settled>((resolve, reject) => {
      this.pending.set(id, { resolve, reject })
      worker.postMessage({ ...request, id } as SpellCheckRequest)
    })
  }

  private failAll(error: Error): void {
    for (const pending of this.pending.values()) pending.reject(error)
    this.pending.clear()
  }
}

export const spellCheckService = new SpellCheckServiceImpl()
