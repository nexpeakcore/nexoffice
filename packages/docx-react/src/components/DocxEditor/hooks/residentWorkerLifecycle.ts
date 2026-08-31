/**
 * A resident worker holds its own replica of one document, bootstrapped from
 * one host engine. It serves that engine and no other.
 *
 * The rule needs stating because the failure is silent: a worker retained from
 * a previous document still reports itself ready and still answers input, and
 * `applyResidentInput` reports any answer — including a failure — as applied.
 * The compatibility path is then skipped, so the keystroke is committed to the
 * previous document and never reaches the one on screen.
 */
export function workerServesEngine<E>(
  worker: { engine: E } | null | undefined,
  engine: E | null | undefined
): boolean {
  return worker != null && engine != null && worker.engine === engine;
}
