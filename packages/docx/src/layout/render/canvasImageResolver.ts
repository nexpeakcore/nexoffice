/**
 * Image resolver for the canvas replay backend, shared by both adapters.
 *
 * v0 display lists carry the flow-block image `src` as the primitive's
 * relId — for embedded media that is a blob:/data: URL minted by the parser.
 * Shape picture fills resolve through the same gate via
 * `fillPaint.pictureSrc`. Only those schemes are decoded; anything else
 * (notably remote http urls from external-mode relationships, or a raw
 * unresolved `rId`) resolves to null so opening a document never triggers a
 * network fetch (the no-zero-click-external-fetch security contract). Decode
 * results are cached per source so repaints reuse the same HTMLImageElement.
 *
 * The cache is bounded: each entry is charged its decoded-bitmap size
 * (width × height × 4) once the load settles, and the least recently used
 * settled entries are dropped when the total passes the budget. An evicted
 * source simply decodes again on its next repaint — blob:/data: URLs are
 * local, so re-resolving costs decode time, never a fetch.
 */

import type { ImageResolver } from './canvasBackend';

/** Roughly a screenful of large photos; big enough that repaints never churn. */
const DEFAULT_BUDGET_BYTES = 64 * 1024 * 1024;

interface CacheEntry {
  promise: Promise<CanvasImageSource | null>;
  /** False while the decode is in flight; in-flight entries are never evicted. */
  settled: boolean;
  /** Decoded RGBA bytes; 0 until settled (and for failed decodes). */
  size: number;
}

function decodedSize(source: CanvasImageSource | null): number {
  if (!source) return 0;
  const { naturalWidth, naturalHeight } = source as {
    naturalWidth?: unknown;
    naturalHeight?: unknown;
  };
  if (typeof naturalWidth === 'number' && typeof naturalHeight === 'number') {
    return naturalWidth * naturalHeight * 4;
  }
  return 0;
}

export function createCanvasImageResolver(
  budgetBytes: number = DEFAULT_BUDGET_BYTES
): ImageResolver {
  const cache = new Map<string, CacheEntry>();
  let cachedBytes = 0;

  const trim = (keep: string): void => {
    if (cachedBytes <= budgetBytes) return;
    for (const [key, entry] of cache) {
      if (cachedBytes <= budgetBytes) return;
      if (key === keep || !entry.settled) continue;
      cache.delete(key);
      cachedBytes -= entry.size;
    }
  };

  return (relId: string) => {
    if (!relId.startsWith('blob:') && !relId.startsWith('data:')) return null;
    const cached = cache.get(relId);
    if (cached) {
      // Re-insert so Map iteration order stays least-recently-used first.
      cache.delete(relId);
      cache.set(relId, cached);
      return cached.promise;
    }
    const entry: CacheEntry = {
      promise: new Promise<CanvasImageSource | null>((resolve) => {
        const img = new Image();
        img.onload = () => resolve(img);
        img.onerror = () => resolve(null);
        img.src = relId;
      }),
      settled: false,
      size: 0,
    };
    cache.set(relId, entry);
    void entry.promise.then((source) => {
      entry.settled = true;
      // Charge only entries still cached; a cleared/evicted entry must not
      // count toward the budget it no longer occupies.
      if (cache.get(relId) !== entry) return;
      entry.size = decodedSize(source);
      cachedBytes += entry.size;
      trim(relId);
    });
    return entry.promise;
  };
}
