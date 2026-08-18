import { beforeEach, describe, expect, test } from 'bun:test'
import { spellCheckService, type Misspelling } from './spellcheck.js'

interface Engine {
  check(text: string): Misspelling[]
  suggest(word: string): string[]
}

// The crash path is only reachable through private state: a worker cannot be
// made to die on demand under the test runner.
const service = spellCheckService as unknown as {
  ready: boolean
  inline: Engine | null
  fallback: Promise<void> | null
  check(text: string): Promise<Misspelling[]>
  suggest(word: string): Promise<string[]>
}

const engine: Engine = {
  check: () => [{ word: 'teh' }],
  suggest: () => ['the'],
}

function crashedWithFallbackInFlight(): void {
  service.ready = false
  service.inline = null
  service.fallback = new Promise<void>((resolve) => {
    setTimeout(() => {
      service.inline = engine
      service.ready = true
      resolve()
    }, 5)
  })
}

describe('spellCheckService', () => {
  beforeEach(() => {
    service.ready = false
    service.inline = null
    service.fallback = null
  })

  test('waits for an in-flight inline fallback instead of answering empty', async () => {
    crashedWithFallbackInFlight()

    expect(await service.check('teh')).toEqual([{ word: 'teh' }])
  })

  test('waits for an in-flight inline fallback before suggesting', async () => {
    crashedWithFallbackInFlight()

    expect(await service.suggest('teh')).toEqual(['the'])
  })

  test('answers empty when proofing never started', async () => {
    expect(await service.check('teh')).toEqual([])
    expect(await service.suggest('teh')).toEqual([])
  })

  test('answers empty when the inline fallback failed to load', async () => {
    service.fallback = Promise.resolve()

    expect(await service.check('teh')).toEqual([])
  })
})
