/// <reference lib="webworker" />

import { SpellCheckEngine } from './spellcheckEngine.js'
import type {
  SpellCheckRequest,
  SpellCheckResponse,
  SpellCheckResult,
} from './spellcheckProtocol.js'

const scope = self as unknown as DedicatedWorkerGlobalScope
const engine = new SpellCheckEngine()

scope.onmessage = (event: MessageEvent<SpellCheckRequest>) => {
  const { id } = event.data
  void handle(event.data)
    .then((result) => {
      scope.postMessage({ id, ok: true, ...result } satisfies SpellCheckResponse)
    })
    .catch((error: unknown) => {
      const message = error instanceof Error ? error.message : String(error)
      scope.postMessage({ id, ok: false, error: message } satisfies SpellCheckResponse)
    })
}

async function handle(request: SpellCheckRequest): Promise<SpellCheckResult> {
  switch (request.type) {
    case 'init':
      await engine.init()
      return {}
    case 'check':
      return { misspellings: engine.check(request.text) }
    case 'suggest':
      return { suggestions: engine.suggest(request.word) }
  }
}
