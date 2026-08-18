import { describe, expect, test } from 'bun:test';
import {
  TextMeasureFontRegistry,
  type BundledFaceLoader,
  type BundledFontProvider,
} from './fontRegistry';

function sink() {
  const registered: Uint8Array[] = [];
  return {
    registered,
    registerFont(bytes: Uint8Array): number {
      registered.push(bytes);
      return registered.length;
    },
  };
}

/** Fresh buffer per call, as a byte cache does after evicting the face. */
function evictingLoader(file: string): BundledFaceLoader {
  const load = () => Promise.resolve(new ArrayBuffer(8));
  load.faceKey = file;
  return load;
}

describe('TextMeasureFontRegistry', () => {
  test('registers a face once even when every load yields a fresh buffer', async () => {
    const engine = sink();
    const bundled: BundledFontProvider = {
      resolve: () => evictingLoader('Liberation-Sans.ttf'),
      resolveLastResort: () => evictingLoader('Liberation-Sans.ttf'),
    };
    const registry = new TextMeasureFontRegistry(engine, { bundled });

    const arial = await registry.getFontIdChain('Arial', false, false);
    const helvetica = await registry.getFontIdChain('Helvetica', false, false);

    expect(engine.registered).toHaveLength(1);
    expect(arial).toEqual([1]);
    expect(helvetica).toEqual([1]);
  });

  test('keeps distinct faces apart', async () => {
    const engine = sink();
    const bundled: BundledFontProvider = {
      resolve: (family) =>
        evictingLoader(family === 'Arial' ? 'Liberation-Sans.ttf' : 'Liberation-Serif.ttf'),
    };
    const registry = new TextMeasureFontRegistry(engine, { bundled });

    await registry.getFontIdChain('Arial', false, false);
    await registry.getFontIdChain('Times New Roman', false, false);

    expect(engine.registered).toHaveLength(2);
  });

  test('falls back to buffer identity when a loader carries no face key', async () => {
    const engine = sink();
    const bytes = new ArrayBuffer(8);
    const bundled: BundledFontProvider = { resolve: () => () => Promise.resolve(bytes) };
    const registry = new TextMeasureFontRegistry(engine, { bundled });

    const arial = await registry.getFontIdChain('Arial', false, false);
    const helvetica = await registry.getFontIdChain('Helvetica', false, false);

    expect(engine.registered).toHaveLength(1);
    expect(arial).toEqual(helvetica);
  });
});
