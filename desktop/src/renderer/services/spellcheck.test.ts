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
  starting: Promise<void> | null
  epoch: number
  dispose(): void
  check(text: string): Promise<Misspelling[]>
  suggest(word: string): Promise<string[]>
  loadInline(): Promise<void>
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

  test('dispose retires the dictionary and later checks answer empty', async () => {
    service.ready = true
    service.inline = engine
    service.starting = Promise.resolve()

    service.dispose()

    expect(service.ready).toBe(false)
    expect(service.inline).toBeNull()
    expect(service.starting).toBeNull()
    expect(service.fallback).toBeNull()
    expect(await service.check('teh')).toEqual([])
  })

  test('an inline load in flight at dispose cannot resurrect the engine', async () => {
    const savedFetch = globalThis.fetch
    globalThis.fetch = (() =>
      Promise.resolve(new Response('SET UTF-8\n'))) as unknown as typeof globalThis.fetch
    try {
      const load = service.loadInline()
      service.dispose()
      await load

      expect(service.inline).toBeNull()
      expect(service.ready).toBe(false)
    } finally {
      globalThis.fetch = savedFetch
    }
  })

  test('an undisturbed inline load installs the engine', async () => {
    const savedFetch = globalThis.fetch
    globalThis.fetch = (() =>
      Promise.resolve(new Response('SET UTF-8\n'))) as unknown as typeof globalThis.fetch
    try {
      await service.loadInline()

      expect(service.inline).not.toBeNull()
      expect(service.ready).toBe(true)
    } finally {
      globalThis.fetch = savedFetch
      service.dispose()
    }
  })
})
