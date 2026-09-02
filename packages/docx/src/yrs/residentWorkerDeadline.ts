import type { ResidentEngineWorkerStage } from './residentEngineWorkerProtocol';

/**
 * How long the host waits for a worker that is opening a document.
 *
 * A single completion deadline cannot tell a broken worker from a slow one, and
 * treating the two alike is a cliff: a worker a moment past the line is killed
 * along with everything it had already computed, and the main thread — on a
 * machine already too busy to have met the deadline — starts the whole document
 * over. Missing by a moment costs many times what finishing late would have.
 *
 * So the budget follows what the worker last reported. A worker that has not
 * yet loaded the document has to prove it is alive quickly, because that is
 * where a genuinely broken one fails: wasm that will not instantiate, a module
 * that never arrives, state that will not decode. Once it reports that layout
 * has begun it has proved it is running, and the only remaining question is how
 * big the document is — so it is given room. The ceiling is not a deadline for
 * the work; it is the point past which a worker is wedged rather than busy.
 */
export const RESIDENT_WORKER_STARTUP_BUDGET_MS = 20_000;

/** Room for the one long wasm call, during which no worker can answer. */
export const RESIDENT_WORKER_LAYOUT_BUDGET_MS = 120_000;

/** Silence allowed before the next sign of life, given the last one seen. */
export function residentWorkerSilenceBudgetMs(stage: ResidentEngineWorkerStage | null): number {
  switch (stage) {
    case null:
    case 'received':
    case 'sessionReady':
    case 'stateLoaded':
      return RESIDENT_WORKER_STARTUP_BUDGET_MS;
    case 'layingOut':
    case 'laidOut':
      return RESIDENT_WORKER_LAYOUT_BUDGET_MS;
  }
}
