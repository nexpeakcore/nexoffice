/** What the main process knows when a window is asked to close. */
export interface CloseState {
  /** Whether the renderer has said it is running and listening. */
  rendererReady: boolean
  /** Whether its web contents have already been torn down. */
  webContentsDestroyed: boolean
  /** Whether a close is already waiting on the renderer's answer. */
  closeRequested: boolean
  /** Whether Electron has reported the renderer as not answering input. */
  rendererUnresponsive: boolean
}

export type CloseDecision =
  /** Let the window close now. */
  | 'close'
  /** Veto the close and ask the renderer about unsaved work. */
  | 'ask'
  /** Veto the close; an answer to the question already asked is still coming. */
  | 'wait'

/**
 * Whether this close may proceed.
 *
 * The renderer is asked once and its answer is waited for. Only two things end
 * that wait: an answer, or Electron reporting that the renderer has stopped
 * answering input at all.
 *
 * Asking again is deliberately not one of them. A save takes seconds on a large
 * deck, and pressing close again during one is impatience, not a diagnosis —
 * treating it as a diagnosis destroys the window mid-write. A timer is no
 * better: it cannot tell a wedged renderer from a prompt sitting on screen
 * waiting to be read.
 */
export function closeDecision(state: CloseState): CloseDecision {
  if (!state.rendererReady || state.webContentsDestroyed) return 'close'
  if (state.closeRequested) return state.rendererUnresponsive ? 'close' : 'wait'
  return 'ask'
}
