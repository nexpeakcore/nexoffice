import { afterEach, beforeEach, describe, expect, test } from 'bun:test';
import { createCanvasImageResolver } from './canvasImageResolver';

/** Controllable stand-in for the DOM Image: tests decide when a load lands. */
class FakeImage {
  onload: (() => void) | null = null;
  onerror: (() => void) | null = null;
  naturalWidth = 0;
  naturalHeight = 0;
  set src(value: string) {
    loads.push({ image: this, src: value });
  }
}

let loads: Array<{ image: FakeImage; src: string }> = [];
const globals = globalThis as { Image?: unknown };
let savedImage: unknown;

function settle(src: string, pixels: number): void {
  const load = loads.find((candidate) => candidate.src === src);
  if (!load) throw new Error(`no load in flight for ${src}`);
  load.image.naturalWidth = pixels;
  load.image.naturalHeight = 1;
  load.image.onload?.();
}

async function flush(): Promise<void> {
  await Promise.resolve();
  await Promise.resolve();
}

describe('createCanvasImageResolver', () => {
  beforeEach(() => {
    loads = [];
    savedImage = globals.Image;
    globals.Image = FakeImage;
  });

  afterEach(() => {
    globals.Image = savedImage;
  });

  test('refuses non-local schemes without caching anything', () => {
    const resolve = createCanvasImageResolver();
    expect(resolve('https://example.com/x.png')).toBeNull();
    expect(resolve('rId7')).toBeNull();
    expect(loads).toHaveLength(0);
  });

  test('reuses the in-flight promise for a repeated source', () => {
    const resolve = createCanvasImageResolver();
    const first = resolve('blob:doc/one');
    const second = resolve('blob:doc/one');
    expect(second).toBe(first);
    expect(loads).toHaveLength(1);
  });

  test('evicts the least recently used settled entry past the byte budget', async () => {
    // Budget of 10 pixels' worth: 40 bytes.
    const resolve = createCanvasImageResolver(40);
    void resolve('blob:doc/a');
    settle('blob:doc/a', 6);
    await flush();
    void resolve('blob:doc/b');
    settle('blob:doc/b', 6);
    await flush();

    // 48 bytes cached: a (oldest) must have been dropped, b kept.
    loads = [];
    void resolve('blob:doc/b');
    expect(loads).toHaveLength(0);
    void resolve('blob:doc/a');
    expect(loads).toHaveLength(1);
  });

  test('a cache hit refreshes recency', async () => {
    const resolve = createCanvasImageResolver(56);
    void resolve('blob:doc/a');
    settle('blob:doc/a', 6);
    await flush();
    void resolve('blob:doc/b');
    settle('blob:doc/b', 6);
    await flush();
    // Touch a so b becomes the least recently used.
    void resolve('blob:doc/a');
    void resolve('blob:doc/c');
    settle('blob:doc/c', 6);
    await flush();

    loads = [];
    void resolve('blob:doc/a');
    void resolve('blob:doc/c');
    expect(loads).toHaveLength(0);
    void resolve('blob:doc/b');
    expect(loads).toHaveLength(1);
  });

  test('never evicts an in-flight load', async () => {
    const resolve = createCanvasImageResolver(40);
    void resolve('blob:doc/pending');
    const pendingLoad = loads.find((candidate) => candidate.src === 'blob:doc/pending');
    void resolve('blob:doc/big');
    settle('blob:doc/big', 100);
    await flush();

    // The pending entry survived the oversized settle beside it.
    loads = [];
    const again = resolve('blob:doc/pending');
    expect(loads).toHaveLength(0);
    pendingLoad!.image.naturalWidth = 1;
    pendingLoad!.image.naturalHeight = 1;
    pendingLoad!.image.onload?.();
    const source = (await again) as { naturalWidth: number } | null;
    expect(source?.naturalWidth).toBe(1);
  });

  test('does not retain an image larger than the whole budget', async () => {
    const resolve = createCanvasImageResolver(40);
    const pending = resolve('blob:doc/huge');
    settle('blob:doc/huge', 100);
    await flush();

    // The caller still gets its decode; the cache just does not keep it.
    expect(await pending).not.toBeNull();
    loads = [];
    void resolve('blob:doc/huge');
    expect(loads).toHaveLength(1);
  });

  test('an oversized image does not evict what fits after it', async () => {
    const resolve = createCanvasImageResolver(40);
    void resolve('blob:doc/small');
    settle('blob:doc/small', 6);
    await flush();
    void resolve('blob:doc/huge');
    settle('blob:doc/huge', 100);
    await flush();

    // huge went instead of evicting small on its way to not fitting either.
    loads = [];
    void resolve('blob:doc/small');
    expect(loads).toHaveLength(0);
  });

  test('a failed decode resolves null and stays cached weightless', async () => {
    const resolve = createCanvasImageResolver(40);
    const pending = resolve('blob:doc/broken');
    const load = loads.find((candidate) => candidate.src === 'blob:doc/broken');
    load?.image.onerror?.();
    expect(await pending).toBeNull();

    loads = [];
    void resolve('blob:doc/broken');
    expect(loads).toHaveLength(0);
  });
});
