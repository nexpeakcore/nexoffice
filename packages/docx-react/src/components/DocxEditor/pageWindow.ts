/**
 * Which pages carry live content.
 *
 * Pages within the window keep live bitmaps and the positioned a11y mirror;
 * everything farther keeps its canvas ELEMENT — stable identity, exact geometry
 * for pointer routing, overlays and scroll math — but releases its backing
 * store and drops to a text-only outline mirror. Page structure and accessible
 * text are never windowed away.
 *
 * The engine reads the same window, so a frame carries content only for pages
 * that can be seen. One definition, because the view and the open path both
 * need it and two would drift.
 */

const PAGE_WINDOW_BUFFER = 2;
/** Documents at or below this page count never window. */
const PAGE_WINDOW_MIN_PAGES = 4;
/** A mounted page stays mounted until it drifts this far beyond the band, so
 * slow scrolling at a boundary cannot thrash mount/unmount. */
const PAGE_WINDOW_HYSTERESIS = 1;

export { PAGE_WINDOW_MIN_PAGES };

export interface PageWindowRange {
  start: number;
  end: number;
}

export function nextPageWindow(
  previous: PageWindowRange | null,
  firstVisible: number,
  lastVisible: number,
  totalPages: number
): PageWindowRange {
  const mountStart = Math.max(0, firstVisible - PAGE_WINDOW_BUFFER);
  const mountEnd = Math.min(totalPages - 1, lastVisible + PAGE_WINDOW_BUFFER);
  if (!previous) return { start: mountStart, end: mountEnd };
  const keepStart = Math.max(0, mountStart - PAGE_WINDOW_HYSTERESIS);
  const keepEnd = Math.min(totalPages - 1, mountEnd + PAGE_WINDOW_HYSTERESIS);
  const start = Math.min(mountStart, Math.max(previous.start, keepStart));
  const end = Math.max(mountEnd, Math.min(previous.end, keepEnd));
  if (start === previous.start && end === previous.end) return previous;
  return { start, end };
}

/**
 * The window the first frame carries, decided before any view exists to
 * measure. A document opens at its first page, so those are the pages worth
 * building; the view widens this as soon as it knows its own viewport.
 *
 * Generous enough to cover a tall viewport at a low zoom, because the cost of
 * being wrong differs by direction: too many pages costs a few milliseconds of
 * frame build, too few shows blank paper until the next frame.
 */
export function openingPageWindow(): { start: number; count: number } {
  return { start: 0, count: PAGE_WINDOW_MIN_PAGES + PAGE_WINDOW_BUFFER * 4 };
}
