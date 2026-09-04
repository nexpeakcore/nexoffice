import type { BundledFontProvider } from '@betteroffice/docx/layout';

/** Regular, bold and italic — what a body of text asks a family for. */
const MEASUREMENT_STYLES: ReadonlyArray<readonly [boolean, boolean]> = [
  [false, false],
  [true, false],
  [false, true],
];

/**
 * Fetch the bundled faces this document's families resolve to, before
 * measurement asks for them.
 *
 * Measurement meets a family only when it reaches text using it, so the bytes
 * are fetched in the middle of the first layout pass and the pass waits on
 * them — 280ms of a 2.2s open on a 343-page document. The families are known
 * as soon as the document is parsed, and a family with no bundled substitute
 * resolves to nothing, so this fetches what would have been fetched anyway.
 */
export function warmMeasurementFaces(
  provider: BundledFontProvider,
  families: readonly string[]
): void {
  const started = new Set<string>();
  for (const family of families) {
    for (const [bold, italic] of MEASUREMENT_STYLES) {
      const load = provider.resolve(family, bold, italic);
      const key = load?.faceKey;
      if (!load || key === undefined || started.has(key)) continue;
      started.add(key);
      // A failure here is not one: the layout pass asks again and reports it.
      void load().catch(() => {});
    }
  }
}
