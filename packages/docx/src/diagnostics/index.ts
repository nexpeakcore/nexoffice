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

export function unregisterMemoryReader(label: string): void {
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
