export type PointerRegion = 'body' | 'header' | 'footer' | null;

export type MousedownRegionAction =
  /** Place or drag a body caret as usual. */
  | 'body'
  /** Leave header/footer edit mode and let the body take the click. */
  | 'exit-hf'
  /** Do nothing: this region's interactions belong to the click handler. */
  | 'ignore';

/**
 * What a mousedown may do, given where it landed and which region is being
 * edited.
 *
 * Outside edit mode a header/footer mousedown is always `ignore`, whatever the
 * click count. Entering the header belongs to the click handler, and the click
 * handler re-hit-tests the same screen point — so the mousedown must leave the
 * scroll position alone. The old exception for double-clicks let the caret
 * path run with no body position under the point, which defaulted the caret to
 * the end of the document; the view scrolled after it, the click handler then
 * found body content under the cursor, and the header never opened.
 */
export function mousedownRegionAction(
  region: PointerRegion,
  hfEditMode: 'header' | 'footer' | null | undefined
): MousedownRegionAction {
  if (hfEditMode) return region === hfEditMode ? 'body' : 'exit-hf';
  if (region === 'header' || region === 'footer') return 'ignore';
  return 'body';
}
