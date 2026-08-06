import { describe, expect, test } from 'bun:test'
import { shareInFlight, type FlightSlot } from './singleFlight.js'

function deferred<T>(): { promise: Promise<T>; resolve: (value: T) => void } {
  let resolve!: (value: T) => void
  const promise = new Promise<T>((settle) => {
    resolve = settle
  })
  return { promise, resolve }
}

describe('shareInFlight', () => {
  test('runs the operation once for callers that overlap', async () => {
    const slot: FlightSlot<boolean> = { current: null }
    const gate = deferred<boolean>()
    let runs = 0

    const first = shareInFlight(slot, () => {
      runs += 1
      return gate.promise
    })
    const second = shareInFlight(slot, () => {
      runs += 1
      return Promise.resolve(false)
    })

    expect(runs).toBe(1)
    gate.resolve(true)
    expect(await first).toBe(true)
    // The second caller gets the first caller's answer, not its own operation's
    // — asking the user twice and acting on both replies is the bug.
    expect(await second).toBe(true)
    expect(runs).toBe(1)
  })

  test('runs again once the first run has settled', async () => {
    const slot: FlightSlot<number> = { current: null }
    let runs = 0
    const run = (): Promise<number> => shareInFlight(slot, () => Promise.resolve(++runs))

    expect(await run()).toBe(1)
    expect(await run()).toBe(2)
    expect(slot.current).toBeNull()
  })

  test('frees the slot when the operation rejects', async () => {
    const slot: FlightSlot<boolean> = { current: null }
    await expect(shareInFlight(slot, () => Promise.reject(new Error('boom')))).rejects.toThrow(
      'boom',
    )
    expect(slot.current).toBeNull()
    expect(await shareInFlight(slot, () => Promise.resolve(true))).toBe(true)
  })

  // A rejection reaches every caller that joined the run, so none of them
  // silently proceeds as though the question had been answered.
  test('gives a joining caller the same rejection', async () => {
    const slot: FlightSlot<boolean> = { current: null }
    const gate = deferred<boolean>()
    const first = shareInFlight(slot, () => gate.promise)
    const second = shareInFlight(slot, () => Promise.resolve(true))
    const failure = new Error('boom')
    ;(gate as unknown as { resolve: (v: Promise<boolean>) => void }).resolve(
      Promise.reject(failure),
    )

    await expect(first).rejects.toThrow('boom')
    await expect(second).rejects.toThrow('boom')
  })

  // A run that finishes after being replaced must not clear the slot the
  // replacement is using, or a third caller would start a duplicate.
  test('a settled run never clears a slot it no longer owns', async () => {
    const slot: FlightSlot<string> = { current: null }
    const stale = Promise.resolve('stale')
    shareInFlight(slot, () => stale)
    const replacement = Promise.resolve('replacement')
    slot.current = replacement

    await stale
    await Promise.resolve()
    expect(slot.current).toBe(replacement)
  })
})
