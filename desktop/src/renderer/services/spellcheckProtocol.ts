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
}

export type SpellCheckResponse =
  | ({ id: number; ok: true } & SpellCheckResult)
  | { id: number; ok: false; error: string }
