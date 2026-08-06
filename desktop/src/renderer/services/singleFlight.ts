/** Somewhere to keep the run in progress, so callers can find it. */
export interface FlightSlot<T> {
  current: Promise<T> | null
}

/**
 * Runs `operation` one at a time.
 *
 * A caller arriving while a run is in flight waits on that run and gets its
 * answer, rather than starting a second one. This is for work that owns
 * something exclusive — the screen, in the case of a modal prompt — where two
 * runs would not merely duplicate effort but ask the user the same question
 * twice and then act on both replies.
 *
 * The slot is only cleared by the run that filled it, so a run that finishes
 * late cannot cancel the one that replaced it.
 */
export function shareInFlight<T>(slot: FlightSlot<T>, operation: () => Promise<T>): Promise<T> {
  const running = slot.current
  if (running) return running
  const started = operation().finally(() => {
    if (slot.current === started) slot.current = null
  })
  slot.current = started
  return started
}
