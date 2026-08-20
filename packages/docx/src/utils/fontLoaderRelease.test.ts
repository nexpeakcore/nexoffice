import { afterAll, describe, expect, test } from 'bun:test';
import { Window } from 'happy-dom';

// fontLoader reads `document` at module scope time-of-use and no-ops without
// one, so the DOM has to exist before the import.
const window = new Window();
const globals = globalThis as Record<string, unknown>;
const saved: Record<string, unknown> = {};
const revoked: string[] = [];
let nextUrl = 0;

for (const key of ['document', 'Blob', 'Event', 'HTMLStyleElement']) {
  saved[key] = globals[key];
  globals[key] = (window as unknown as Record<string, unknown>)[key];
}
const savedCreate = URL.createObjectURL;
const savedRevoke = URL.revokeObjectURL;
URL.createObjectURL = (() => `blob:font/${nextUrl++}`) as typeof URL.createObjectURL;
URL.revokeObjectURL = ((url: string) => {
  revoked.push(url);
}) as typeof URL.revokeObjectURL;

const {
  createBufferFontOwner,
  loadFontFromBuffer,
  loadFontFromUrl,
  loadFontWithMapping,
  releaseBufferFontFaces,
  isFontLoaded,
} = await import('./fontLoader');

const doc = globals.document as Document;

function styleRules(): string {
  return [...doc.head.querySelectorAll('style')].map((el) => el.textContent ?? '').join('\n');
}

function bytes(): ArrayBuffer {
  return new Uint8Array([0, 1, 2, 3]).buffer;
}

/** The Google Fonts <link> loadFontWithMapping injects, if it got that far. */
function googleLinks(name: string): HTMLLinkElement[] {
  return [...doc.head.querySelectorAll('link')].filter((el) =>
    (el.getAttribute('href') ?? '').includes(name)
  ) as HTMLLinkElement[];
}

describe('releaseBufferFontFaces', () => {
  afterAll(() => {
    URL.createObjectURL = savedCreate;
    URL.revokeObjectURL = savedRevoke;
    for (const [key, value] of Object.entries(saved)) globals[key] = value;
  });

  test('registering an embedded face injects a rule and an object URL', async () => {
    await loadFontFromBuffer('EmbeddedAlpha', bytes(), { weight: 400 });

    expect(styleRules()).toContain('EmbeddedAlpha');
    expect(isFontLoaded('EmbeddedAlpha')).toBe(true);
  });

  test('releasing removes the rule and revokes the bytes', async () => {
    const owner = createBufferFontOwner();
    await loadFontFromBuffer('EmbeddedBeta', bytes(), { weight: 400, owner });
    const before = revoked.length;

    releaseBufferFontFaces(owner);

    expect(styleRules()).not.toContain('EmbeddedBeta');
    expect(revoked.length).toBe(before + 1);
    expect(isFontLoaded('EmbeddedBeta')).toBe(false);
  });

  test('every weight that owner registered is dropped', async () => {
    const owner = createBufferFontOwner();
    await loadFontFromBuffer('EmbeddedGamma', bytes(), { weight: 400, owner });
    await loadFontFromBuffer('EmbeddedGamma', bytes(), { weight: 700, owner });
    const before = revoked.length;

    releaseBufferFontFaces(owner);

    expect(revoked.length).toBe(before + 2);
    expect(styleRules()).not.toContain('EmbeddedGamma');
  });

  test('a released family registers again on the next document', async () => {
    const first = createBufferFontOwner();
    await loadFontFromBuffer('EmbeddedDelta', bytes(), { weight: 400, owner: first });
    releaseBufferFontFaces(first);
    expect(styleRules()).not.toContain('EmbeddedDelta');

    // The face key must have been cleared, or this call short-circuits and
    // the document renders against a rule that no longer exists.
    await loadFontFromBuffer('EmbeddedDelta', bytes(), {
      weight: 400,
      owner: createBufferFontOwner(),
    });

    expect(styleRules()).toContain('EmbeddedDelta');
    expect(isFontLoaded('EmbeddedDelta')).toBe(true);
  });

  test('a face two documents embed survives the first release', async () => {
    const stale = createBufferFontOwner();
    const current = createBufferFontOwner();
    await loadFontFromBuffer('EmbeddedShared', bytes(), { weight: 400, owner: stale });
    await loadFontFromBuffer('EmbeddedShared', bytes(), { weight: 400, owner: current });
    const before = revoked.length;

    releaseBufferFontFaces(stale);

    expect(styleRules()).toContain('EmbeddedShared');
    expect(isFontLoaded('EmbeddedShared')).toBe(true);
    expect(revoked.length).toBe(before);

    releaseBufferFontFaces(current);

    expect(styleRules()).not.toContain('EmbeddedShared');
    expect(revoked.length).toBe(before + 1);
  });

  test('a document that joins an in-flight load still holds the face', async () => {
    const stale = createBufferFontOwner();
    const current = createBufferFontOwner();
    const first = loadFontFromBuffer('EmbeddedRacing', bytes(), { weight: 400, owner: stale });
    const second = loadFontFromBuffer('EmbeddedRacing', bytes(), { weight: 400, owner: current });
    await Promise.all([first, second]);

    releaseBufferFontFaces(stale);

    expect(styleRules()).toContain('EmbeddedRacing');
    expect(isFontLoaded('EmbeddedRacing')).toBe(true);
  });

  test('leaves consumer-hosted faces of the same family alone', async () => {
    const owner = createBufferFontOwner();
    await loadFontFromUrl('SharedName', 'https://example.test/shared.woff2', { weight: 700 });
    await loadFontFromBuffer('SharedName', bytes(), { weight: 400, owner });
    const before = revoked.length;

    releaseBufferFontFaces(owner);

    // The url-registered face's rule survives; only the buffer face's bytes
    // were revoked, and the family stays "loaded" because a face remains.
    expect(revoked.length).toBe(before + 1);
    expect(styleRules()).toContain('shared.woff2');
    expect(isFontLoaded('SharedName')).toBe(true);
  });

  test('a consumer face of the same weight and style survives the release', async () => {
    const owner = createBufferFontOwner();
    await loadFontFromBuffer('ExactName', bytes(), { weight: 400, owner });
    // Same family, weight and style as the embedded face: the consumer's rule
    // must still be injected rather than deduped against the document's.
    await loadFontFromUrl('ExactName', 'https://example.test/exact.woff2', { weight: 400 });
    expect(styleRules()).toContain('exact.woff2');

    releaseBufferFontFaces(owner);

    expect(styleRules()).toContain('exact.woff2');
    expect(isFontLoaded('ExactName')).toBe(true);
  });

  test('an embedded face registers even when a consumer face already has the key', async () => {
    await loadFontFromUrl('UrlFirst', 'https://example.test/urlfirst.woff2', { weight: 400 });
    const owner = createBufferFontOwner();

    await loadFontFromBuffer('UrlFirst', bytes(), { weight: 400, owner });

    // Two rules for the family now: the document's own bytes, and the
    // consumer's URL. Releasing the document leaves the consumer's.
    expect(styleRules()).toContain('UrlFirst');
    releaseBufferFontFaces(owner);
    expect(styleRules()).toContain('urlfirst.woff2');
    expect(isFontLoaded('UrlFirst')).toBe(true);
  });

  test('a release keeps family provenance while a same-family face is loading', async () => {
    const owner = createBufferFontOwner();
    await loadFontFromBuffer('Calibri', bytes(), { weight: 400, owner });
    // A second face of the family, still in flight when the document closes.
    const pending = loadFontFromUrl('Calibri', 'https://example.test/calibri-bold.woff2', {
      weight: 700,
    });

    releaseBufferFontFaces(owner);
    await pending;

    // Calibri is still a registered — so possibly subsetted — family, which is
    // what keeps its metric-compatible Google equivalent being fetched as the
    // glyph-coverage fallback. Dropping the marker mid-load loses that.
    expect(googleLinks('Carlito')).toHaveLength(0);
    const mapping = loadFontWithMapping('Calibri');
    const injected = googleLinks('Carlito');
    expect(injected).toHaveLength(1);
    injected[0].dispatchEvent(new (globals.Event as typeof Event)('error'));
    await mapping;
  });

  test('releasing an owner that registered nothing is a no-op', async () => {
    await loadFontFromBuffer('EmbeddedEpsilon', bytes(), { weight: 400 });
    const before = revoked.length;

    releaseBufferFontFaces(createBufferFontOwner());

    expect(revoked.length).toBe(before);
    expect(styleRules()).toContain('EmbeddedEpsilon');
  });
});
