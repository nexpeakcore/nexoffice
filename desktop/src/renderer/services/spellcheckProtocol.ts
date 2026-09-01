import type { Misspelling } from './spellcheckEngine.js'

export type SpellCheckRequest =
  | { id: number; type: 'init' }
  | { id: number; type: 'check'; text: string }
  | { id: number; type: 'suggest'; word: string }

export type SpellCheckRequestWithoutId = SpellCheckRequest extends infer Request
  ? Request extends { id: number }
    ? Omit<Request, 'id'>
    : never
  : never

/** Answer payload, before the worker stamps the request id onto it. */
export interface SpellCheckResult {
  misspellings?: Misspelling[]
  suggestions?: string[]
  /** The worker isolate's JS heap, where the ~40MB dictionary lives. The main
   * thread's `performance.memory` covers its own isolate only, so this is the
   * only way that memory can be told apart from the renderer's. */
  heapBytes?: number
}

export type SpellCheckResponse =
  | ({ id: number; ok: true } & SpellCheckResult)
  | { id: number; ok: false; error: string }
