import { afterAll, describe, expect, test } from 'bun:test';

// registerBundledFontFace decides at call time whether a DOM exists, so the
// stubs go in before the import. No happy-dom: the module touches exactly
// three globals, and faking them keeps the font bytes out of the test.
class FakeFontFace {
  constructor(
    readonly family: string,
    readonly source: ArrayBuffer,
    readonly descriptors: Record<string, string>
  ) {}
  load(): Promise<this> {
    return Promise.resolve(this);
  }
}

const installed = new Set<unknown>();
const fontSet = {
  add(face: unknown): void {
    installed.add(face);
  },
  delete(face: unknown): boolean {
    return installed.delete(face);
  },
};

const globals = globalThis as Record<string, unknown>;
const saved = {
  document: globals.document,
  FontFace: globals.FontFace,
  fetch: globals.fetch,
};

/** Set to hold the next fetch open, so a release can race an in-flight load. */
let hold: (() => void) | null = null;

globals.document = { fonts: fontSet };
globals.FontFace = FakeFontFace;
globals.fetch = (() => {
  const body = { ok: true, arrayBuffer: () => Promise.resolve(new ArrayBuffer(8)) } as Response;
  if (!hold) return Promise.resolve(body);
  return new Promise<Response>((resolve) => {
    hold = () => resolve(body);
  });
}) as unknown as typeof globalThis.fetch;

const {
  createFontFaceOwner,
  registerBundledFontFace,
  releaseBundledFontFaces,
  resolveMetricCompatFace,
} = await import('./index');

const arial = resolveMetricCompatFace('Arial', false, false)!;
const times = resolveMetricCompatFace('Times New Roman', false, false)!;
const courier = resolveMetricCompatFace('Courier New', false, false)!;

function familiesInstalled(): string[] {
  return [...installed].map((face) => (face as FakeFontFace).family);
}

describe('releaseBundledFontFaces', () => {
  afterAll(() => {
    for (const [key, value] of Object.entries(saved)) globals[key] = value;
  });

  test('registering installs the face under the requested family', async () => {
    await registerBundledFontFace(arial, 'DeckAlpha', createFontFaceOwner());

    expect(familiesInstalled()).toContain('DeckAlpha');
  });

  test('releasing uninstalls it', async () => {
    const owner = createFontFaceOwner();
    await registerBundledFontFace(arial, 'DeckBeta', owner);
    expect(familiesInstalled()).toContain('DeckBeta');

    releaseBundledFontFaces(owner);

    expect(familiesInstalled()).not.toContain('DeckBeta');
  });

  test('releases only what that owner registered', async () => {
    const keeper = createFontFaceOwner();
    const dropper = createFontFaceOwner();
    await registerBundledFontFace(arial, 'DeckKeep', keeper);
    await registerBundledFontFace(times, 'DeckDrop', dropper);

    releaseBundledFontFaces(dropper);

    expect(familiesInstalled()).toContain('DeckKeep');
    expect(familiesInstalled()).not.toContain('DeckDrop');
  });

  test('the same face under many families is one registration each', async () => {
    const owner = createFontFaceOwner();
    await registerBundledFontFace(arial, 'DeckMulti1', owner);
    await registerBundledFontFace(arial, 'DeckMulti2', owner);
    expect(familiesInstalled().filter((f) => f.startsWith('DeckMulti'))).toHaveLength(2);

    releaseBundledFontFaces(owner);

    expect(familiesInstalled().filter((f) => f.startsWith('DeckMulti'))).toHaveLength(0);
  });

  test('a family two owners registered survives the first release', async () => {
    const first = createFontFaceOwner();
    const second = createFontFaceOwner();
    await registerBundledFontFace(arial, 'DeckShared', first);
    await registerBundledFontFace(arial, 'DeckShared', second);

    releaseBundledFontFaces(first);
    expect(familiesInstalled()).toContain('DeckShared');

    releaseBundledFontFaces(second);
    expect(familiesInstalled()).not.toContain('DeckShared');
  });

  test('an owner that registered the same face twice releases it once', async () => {
    const owner = createFontFaceOwner();
    await registerBundledFontFace(arial, 'DeckTwice', owner);
    await registerBundledFontFace(arial, 'DeckTwice', owner);

    releaseBundledFontFaces(owner);

    expect(familiesInstalled()).not.toContain('DeckTwice');
  });

  test('a released family registers again on the next deck', async () => {
    const owner = createFontFaceOwner();
    await registerBundledFontFace(arial, 'DeckReopen', owner);
    releaseBundledFontFaces(owner);
    expect(familiesInstalled()).not.toContain('DeckReopen');

    // The memo must have been cleared, or this resolves against a registration
    // whose face is no longer installed.
    await registerBundledFontFace(arial, 'DeckReopen', createFontFaceOwner());

    expect(familiesInstalled()).toContain('DeckReopen');
  });

  test('releasing an owner that registered nothing is a no-op', () => {
    const before = installed.size;

    releaseBundledFontFaces(createFontFaceOwner());

    expect(installed.size).toBe(before);
  });

  test('a release during an in-flight load stops the face being installed', async () => {
    hold = () => {};
    const owner = createFontFaceOwner();
    const pending = registerBundledFontFace(courier, 'DeckRacing', owner);

    releaseBundledFontFaces(owner);
    hold?.();
    hold = null;
    await pending;

    expect(familiesInstalled()).not.toContain('DeckRacing');
  });

  test('a deck that joins an in-flight load keeps the face when the first lets go', async () => {
    hold = () => {};
    const stale = createFontFaceOwner();
    const current = createFontFaceOwner();
    const first = registerBundledFontFace(times, 'DeckHandover', stale);
    const second = registerBundledFontFace(times, 'DeckHandover', current);

    hold?.();
    hold = null;
    await Promise.all([first, second]);
    releaseBundledFontFaces(stale);

    expect(familiesInstalled()).toContain('DeckHandover');

    releaseBundledFontFaces(current);
    expect(familiesInstalled()).not.toContain('DeckHandover');
  });
});
