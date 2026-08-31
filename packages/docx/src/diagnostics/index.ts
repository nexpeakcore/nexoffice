/**
 * Memory attribution for the document renderer.
 *
 * Web Workers are threads inside the renderer process, so Electron's
 * per-process metrics fold the resident engine, the image cache and the JS
 * heap into one number. Producers register a byte reader here; a host reads
 * them back to say which of them is holding a heavy document's memory.
 */

export interface MemoryRow {
  label: string;
  bytes: number;
}

const readers = new Map<string, () => number>();

/** Register (or replace) the byte reader published under `label`. */
export function registerMemoryReader(label: string, read: () => number): void {
  readers.set(label, read);
}

/**
 * `reader` makes this safe to call during teardown: a later owner of the same
 * label has already replaced the entry, and removing the label outright would
 * take the live reader with it.
 */
export function unregisterMemoryReader(label: string, reader?: () => number): void {
  if (reader && readers.get(label) !== reader) return;
  readers.delete(label);
}

/** Every registered reader that currently holds bytes, largest first. */
export function memoryReport(): MemoryRow[] {
  const rows: MemoryRow[] = [];
  for (const [label, read] of readers) {
    let bytes = 0;
    try {
      bytes = read();
    } catch {
      continue;
    }
    if (bytes > 0) rows.push({ label, bytes });
  }
  return rows.sort((a, b) => b.bytes - a.bytes);
}

export function memoryTotalBytes(): number {
  return memoryReport().reduce((sum, row) => sum + row.bytes, 0);
}

// ---------------------------------------------------------------------------
// Heap stages.
//
// `WebAssembly.Memory.buffer.byteLength` — what the readers above report for a
// wasm module — is the linear memory the allocator has claimed. It includes
// free lists, never shrinks, and so cannot answer how much of the wasm32
// address space a document actually needs, or which phase of opening it spends
// that space in. The editing core counts its own allocator instead; this is
// where a host samples those counters at phase boundaries.

export interface HeapStage {
  /** Phase that just finished: `seed`, `layout`, `display`. */
  label: string;
  /** Bytes still held when the phase ended. */
  liveBytes: number;
  /** Highest figure reached during the phase — transient cost the retained
   * number hides, and what actually decides whether a ceiling is hit. */
  peakBytes: number;
  /** Milliseconds from the first mark of this open. */
  atMs: number;
}

export interface HeapProbe {
  liveBytes(): number;
  peakBytes(): number;
  /** Restart the high-water mark so the next phase's peak is its own. */
  resetPeak(): void;
}

let probe: HeapProbe | null = null;
let stages: HeapStage[] = [];
let firstMarkAt = 0;
const stageListeners = new Set<() => void>();

/**
 * Called after every mark. A host that ships these somewhere must re-send on
 * this signal rather than sample on demand: the phases worth measuring are the
 * ones that block the thread, so a request made during them is served late,
 * and the report ends up missing exactly the phases it was opened to explain.
 */
export function onHeapStage(listener: () => void): () => void {
  stageListeners.add(listener);
  return () => {
    stageListeners.delete(listener);
  };
}

/** Publish the editing core's allocator counters. Replaces any previous probe. */
export function registerHeapProbe(next: HeapProbe | null): void {
  probe = next;
}

/**
 * Record what the phase that just ended holds and peaked at, then hand the
 * next phase a fresh high-water mark. A no-op without a probe, so callers can
 * mark unconditionally.
 */
export function markHeapStage(label: string): void {
  // Listeners fire either way. They exist so a host re-sends its report at
  // phase boundaries, and a build without the counters still needs that: the
  // phases worth reporting are the ones that block the thread, and the figures
  // the host does have go stale across them just the same.
  if (probe) {
    recordStage(label);
  }
  for (const listener of stageListeners) {
    try {
      listener();
    } catch {
      // Diagnostics must never break the open they are measuring.
    }
  }
}

function recordStage(label: string): void {
  if (!probe) return;
  try {
    const now = performance.now();
    if (stages.length === 0) firstMarkAt = now;
    stages.push({
      label,
      liveBytes: probe.liveBytes(),
      peakBytes: probe.peakBytes(),
      atMs: now - firstMarkAt,
    });
    probe.resetPeak();
  } catch {
    // As above.
  }
}

export function heapStages(): HeapStage[] {
  return [...stages];
}

/** Drop the previous document's phases; call when a new open begins. */
export function clearHeapStages(): void {
  stages = [];
  firstMarkAt = 0;
  try {
    probe?.resetPeak();
  } catch {
    // As above.
  }
}
