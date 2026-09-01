/**
 * The renderer's account of its own memory and of what opening a document
 * cost.
 *
 * Chromium runs Web Workers as threads inside the renderer process, so the
 * resident layout engine's wasm heap, the image cache and the JS heap all land
 * in one `getAppMetrics()` number. Only this side can take them apart, so the
 * breakdown is assembled here and pushed to the main process, which merges it
 * into the process report.
 */

import { heapStages, memoryReport, onHeapStage } from '@betteroffice/docx/diagnostics'

// The registry backing the breakdown. Re-exported so anything in the renderer
// registers through this module rather than reaching for the docx package,
// which happens to own it but does not own what it measures.
export { registerMemoryReader, unregisterMemoryReader } from '@betteroffice/docx/diagnostics'

import type {
  DocumentProfile,
  MemoryBreakdownRow,
  OpenPhaseTimings,
  OpenedDocument,
  RendererDiagnostics,
} from '../../shared/ipc.js'

/** Chromium's JS heap counter, which the DOM lib does not declare. */
interface PerformanceMemory {
  usedJSHeapSize?: number
}

function jsHeapBytes(): number {
  const { memory } = performance as Performance & { memory?: PerformanceMemory }
  const used = memory?.usedJSHeapSize
  return typeof used === 'number' && Number.isFinite(used) ? used : 0
}

let profile: DocumentProfile | null = null
let timings: OpenPhaseTimings | null = null
/** Phases fixed at open time; the frame loop extends a copy of these. */
let transportTimings: OpenPhaseTimings = {}
let openedAt = 0
/** Bumped per open so a still-running frame loop cannot publish over a newer one. */
let openEpoch = 0

function collect(): RendererDiagnostics {
  const rows: MemoryBreakdownRow[] = memoryReport()
  const heap = jsHeapBytes()
  if (heap > 0) rows.push({ label: 'JS heap', bytes: heap })
  return { document: profile, open: timings, memory: rows, heapStages: heapStages() }
}

// A phase that blocks the thread cannot answer a sample request while it is
// blocking, so each completed phase pushes its own snapshot instead.
onHeapStage(() => publishDiagnostics())

export function publishDiagnostics(): void {
  window.nexoffice.reportDiagnostics(collect())
}

/** Answer the main process's request for a fresh sample. */
export function serveDiagnosticsSamples(): () => void {
  return window.nexoffice.onDiagnosticsSample(publishDiagnostics)
}

export function clearDocumentProfile(): void {
  openEpoch += 1
  contentPaintedAt = null
  awaitingContent = false
  profile = null
  timings = null
  transportTimings = {}
  publishDiagnostics()
}

/**
 * Record which document was opened and what reaching the renderer cost. Call
 * synchronously as the bytes are handed to the editor; {@link trackOpenSettle}
 * then measures what the editor does with them.
 */
export function noteDocumentOpened(opened: OpenedDocument, kind: DocumentProfile['kind']): void {
  openEpoch += 1
  contentPaintedAt = null
  awaitingContent = false
  openedAt = performance.now()
  profile = { kind, name: opened.name, bytes: opened.data.byteLength }
  const next: OpenPhaseTimings = {}
  if (opened.readMs !== undefined) next.read = opened.readMs
  if (opened.sentAt !== undefined) next.transfer = Math.max(0, Date.now() - opened.sentAt)
  transportTimings = next
  timings = next
  publishDiagnostics()
}

/**
 * When the editor's own content reached the screen. An editor that lays its
 * document out on a worker leaves this thread idle while it does, so an
 * unblocked frame no longer means the document is up — without this the open
 * measurement reports the shell mounting and calls it done.
 */
let contentPaintedAt: number | null = null;
let awaitingContent = false;

/**
 * Declare that this editor will report when its content lands. Registered on
 * mount, so it is in place before {@link trackOpenSettle} sees its first
 * frame. An editor that never reports leaves the open measurement without an
 * `interactive` figure rather than with a wrong one.
 */
export function expectDocumentContent(): () => void {
  awaitingContent = true;
  return () => {
    awaitingContent = false;
  };
}

/** The editor's first content is on screen. */
export function noteDocumentContentPainted(): void {
  if (contentPaintedAt === null) contentPaintedAt = performance.now();
}

/** A frame this long means the main thread was busy through it rather than
 * idling between paints — two missed frames at 60Hz. */
const BLOCKED_FRAME_MS = 32
/** Stop watching rather than follow an animation that never settles. */
const MAX_OBSERVED_FRAMES = 900

/**
 * Measure the editor's share of the open, from the commit that mounted it.
 *
 * `mount` ends at the first painted frame — the editor shell is up, though a
 * heavy document is still parsing. `interactive` ends at the first unblocked
 * frame at or after the editor's content reaches the screen, which is where
 * the document stops feeling frozen. Both halves are needed: an unblocked
 * frame alone says only that this thread is idle, which it also is while a
 * worker lays the document out. A frame budget is used rather than a settle
 * timer so the number tracks the document's real cost instead of the timer's.
 */
export function trackOpenSettle(): () => void {
  const epoch = openEpoch
  const committed = openedAt
  let mountedAt: number | null = null
  let previousFrame = performance.now()
  let frames = 0
  let handle = 0
  let stopped = false

  const finish = (interactive: number | null): void => {
    if (mountedAt === null) return
    const next: OpenPhaseTimings = { ...transportTimings, mount: mountedAt - committed }
    if (interactive !== null) next.interactive = interactive
    timings = next
    publishDiagnostics()
  }

  const observe = (): void => {
    if (stopped || epoch !== openEpoch) return
    const now = performance.now()
    const frameCost = now - previousFrame
    previousFrame = now
    frames += 1

    if (mountedAt === null) {
      mountedAt = now
      timings = { ...transportTimings, mount: now - committed }
      handle = requestAnimationFrame(observe)
      return
    }
    const settled = frameCost <= BLOCKED_FRAME_MS && (!awaitingContent || contentPaintedAt !== null)
    if (!settled && frames < MAX_OBSERVED_FRAMES) {
      handle = requestAnimationFrame(observe)
      return
    }
    finish(settled ? now - mountedAt : null)
  }

  handle = requestAnimationFrame(observe)
  return () => {
    stopped = true
    cancelAnimationFrame(handle)
  }
}
