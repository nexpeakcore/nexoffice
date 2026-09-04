import { describe, expect, test } from 'bun:test';
import type { BundledFontProvider } from '@betteroffice/docx/layout';
import { warmMeasurementFaces } from './warmMeasurementFaces';

function providerFor(bundled: Record<string, string>): {
  provider: BundledFontProvider;
  fetched: string[];
} {
  const fetched: string[] = [];
  const provider: BundledFontProvider = {
    resolve(family, bold, italic) {
      const substitute = bundled[family.toLowerCase()];
      if (!substitute) return undefined;
      const face = `${substitute}-${bold ? 'Bold' : italic ? 'Italic' : 'Regular'}`;
      const load = () => {
        fetched.push(face);
        return Promise.resolve(new ArrayBuffer(1));
      };
      load.faceKey = face;
      return load;
    },
    resolveLastResort(_family, _bold, _italic) {
      throw new Error('the last resort is warmed by the host, not from a family');
    },
  } as BundledFontProvider;
  return { provider, fetched };
}

describe('warmMeasurementFaces', () => {
  test('fetches the faces a document’s families resolve to', () => {
    // Without this the bytes are fetched mid-layout, once measurement reaches
    // text in the family, and the layout pass waits on them.
    const { provider, fetched } = providerFor({ calibri: 'Carlito' });
    warmMeasurementFaces(provider, ['Calibri']);
    expect(fetched.sort()).toEqual(['Carlito-Bold', 'Carlito-Italic', 'Carlito-Regular']);
  });

  test('asks for each face once, however many families reach it', () => {
    const { provider, fetched } = providerFor({ calibri: 'Carlito', 'calibri light': 'Carlito' });
    warmMeasurementFaces(provider, ['Calibri', 'Calibri Light', 'Calibri']);
    expect(fetched).toHaveLength(3);
  });

  test('passes over a family with no bundled substitute', () => {
    const { provider, fetched } = providerFor({ calibri: 'Carlito' });
    warmMeasurementFaces(provider, ['Some Foundry Display']);
    expect(fetched).toEqual([]);
  });

  test('a face that fails to load is not the open failing', () => {
    const provider = {
      resolve() {
        const load = () => Promise.reject(new Error('offline'));
        load.faceKey = 'Broken-Regular';
        return load;
      },
    } as unknown as BundledFontProvider;
    expect(() => warmMeasurementFaces(provider, ['Broken'])).not.toThrow();
  });
});
