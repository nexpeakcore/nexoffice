import { afterAll, describe, expect, test } from 'bun:test';
import { Window } from 'happy-dom';

// fontLoader reads `document` at module scope time-of-use and no-ops without
// one, so the DOM has to exist before the import.
const window = new Window();
const globals = globalThis as Record<string, unknown>;
const saved: Record<string, unknown> = {};
const revoked: string[] = [];
let nextUrl = 0;

for (const key of ['document', 'Blob', 'HTMLStyleElement']) {
  saved[key] = globals[key];
  globals[key] = (window as unknown as Record<string, unknown>)[key];
}
const savedCreate = URL.createObjectURL;
const savedRevoke = URL.revokeObjectURL;
URL.createObjectURL = (() => `blob:font/${nextUrl++}`) as typeof URL.createObjectURL;
URL.revokeObjectURL = ((url: string) => {
  revoked.push(url);
}) as typeof URL.revokeObjectURL;

const { loadFontFromBuffer, loadFontFromUrl, releaseBufferFontFaces, isFontLoaded } = await import(
  './fontLoader'
);

const doc = globals.document as Document;

function styleRules(): string {
  return [...doc.head.querySelectorAll('style')].map((el) => el.textContent ?? '').join('\n');
}

function bytes(): ArrayBuffer {
  return new Uint8Array([0, 1, 2, 3]).buffer;
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
    await loadFontFromBuffer('EmbeddedBeta', bytes(), { weight: 400 });
    const before = revoked.length;

    releaseBufferFontFaces(['EmbeddedBeta']);

    expect(styleRules()).not.toContain('EmbeddedBeta');
    expect(revoked.length).toBe(before + 1);
    expect(isFontLoaded('EmbeddedBeta')).toBe(false);
  });

  test('every weight of a released family is dropped', async () => {
    await loadFontFromBuffer('EmbeddedGamma', bytes(), { weight: 400 });
    await loadFontFromBuffer('EmbeddedGamma', bytes(), { weight: 700 });
    const before = revoked.length;

    releaseBufferFontFaces(['EmbeddedGamma']);

    expect(revoked.length).toBe(before + 2);
    expect(styleRules()).not.toContain('EmbeddedGamma');
  });

  test('a released family registers again on the next document', async () => {
    await loadFontFromBuffer('EmbeddedDelta', bytes(), { weight: 400 });
    releaseBufferFontFaces(['EmbeddedDelta']);
    expect(styleRules()).not.toContain('EmbeddedDelta');

    // The face key must have been cleared, or this call short-circuits and
    // the document renders against a rule that no longer exists.
    await loadFontFromBuffer('EmbeddedDelta', bytes(), { weight: 400 });

    expect(styleRules()).toContain('EmbeddedDelta');
    expect(isFontLoaded('EmbeddedDelta')).toBe(true);
  });

  test('leaves consumer-hosted faces of the same family alone', async () => {
    await loadFontFromUrl('SharedName', 'https://example.test/shared.woff2', { weight: 700 });
    await loadFontFromBuffer('SharedName', bytes(), { weight: 400 });
    const before = revoked.length;

    releaseBufferFontFaces(['SharedName']);

    // The url-registered face's rule survives; only the buffer face's bytes
    // were revoked, and the family stays "loaded" because a face remains.
    expect(revoked.length).toBe(before + 1);
    expect(styleRules()).toContain('shared.woff2');
    expect(isFontLoaded('SharedName')).toBe(true);
  });

  test('releasing an unknown family is a no-op', () => {
    const before = revoked.length;

    releaseBufferFontFaces(['NeverRegistered']);

    expect(revoked.length).toBe(before);
  });

  test('releasing nothing touches nothing', async () => {
    await loadFontFromBuffer('EmbeddedEpsilon', bytes(), { weight: 400 });
    const before = revoked.length;

    releaseBufferFontFaces([]);

    expect(revoked.length).toBe(before);
    expect(styleRules()).toContain('EmbeddedEpsilon');
  });
});
