import { describe, expect, test } from 'bun:test'
import { closeDecision, type CloseState } from './closePolicy.js'

const live: CloseState = {
  rendererReady: true,
  webContentsDestroyed: false,
  closeRequested: false,
  rendererUnresponsive: false,
}

describe('closeDecision', () => {
  test('asks the renderer the first time a live window is closed', () => {
    expect(closeDecision(live)).toBe('ask')
  })

  test('closes a window whose renderer never came up', () => {
    expect(closeDecision({ ...live, rendererReady: false })).toBe('close')
    expect(closeDecision({ ...live, webContentsDestroyed: true })).toBe('close')
  })

  // The regression this exists for. Pressing close again while a save runs is
  // impatience, and reading it as permission destroys the window mid-write —
  // the document is gone and the file may be half-written.
  test('keeps waiting when asked again while the renderer still answers', () => {
    expect(closeDecision({ ...live, closeRequested: true })).toBe('wait')
  })

  test('closes once the renderer has stopped answering input', () => {
    expect(closeDecision({ ...live, closeRequested: true, rendererUnresponsive: true })).toBe(
      'close',
    )
  })

  // Being unresponsive is not on its own a reason to skip the question: the
  // renderer may still come back, and it has not been asked yet.
  test('still asks an unresponsive renderer that has not been asked', () => {
    expect(closeDecision({ ...live, rendererUnresponsive: true })).toBe('ask')
  })

  // A dead renderer outranks a pending question — nothing is coming.
  test('closes a torn-down window even with a question outstanding', () => {
    expect(closeDecision({ ...live, rendererReady: false, closeRequested: true })).toBe('close')
  })

  test('never returns anything but the three decisions', () => {
    const flags = [false, true]
    for (const rendererReady of flags) {
      for (const webContentsDestroyed of flags) {
        for (const closeRequested of flags) {
          for (const rendererUnresponsive of flags) {
            const decision = closeDecision({
              rendererReady,
              webContentsDestroyed,
              closeRequested,
              rendererUnresponsive,
            })
            expect(['close', 'ask', 'wait']).toContain(decision)
          }
        }
      }
    }
  })

  // Whatever else changes, a live renderer that was asked and is still
  // answering must never have its window taken away.
  test('a live, answering renderer is never closed out from under', () => {
    for (const closeRequested of [false, true]) {
      expect(closeDecision({ ...live, closeRequested })).not.toBe('close')
    }
  })
})
