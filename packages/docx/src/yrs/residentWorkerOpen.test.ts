import { beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { parseDocx } from '../docx';
import { buildResidentRegionLayoutRequest } from '../editor/computeLayout';
import { applyFrameDelta, decodeFrameDelta } from '../layout/render/frameDelta';
import type { Layout } from '../layout/pagination';
import { preloadEditWasm } from '../wasm/edit';
import { createYrsSession, type YrsSession } from './index';

const WASM = resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm');
const DOCX = resolve(import.meta.dir, '__fixtures__/titlepg-first-footer.docx');
const FONT = resolve(import.meta.dir, '../../../fonts/assets/LiberationSans-Regular.ttf');

const RENDER_ENV = { themeColors: {}, numericIds: {} };

function measurementFor(fontId: number) {
  return {
    fontChains: { 'liberation sans|0|0': [fontId] },
    defaults: { fontSize: 24, fontFamily: 'Liberation Sans' },
    authoritativeShaping: true,
  };
}

describe('handing a document to a worker to lay out', () => {
  beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(WASM))));

  it('paginates from state alone, with nothing replayed', async () => {
    const bytes = new Uint8Array(readFileSync(DOCX));
    const font = new Uint8Array(readFileSync(FONT));
    const parsed = await parseDocx(bytes.buffer as ArrayBuffer, { preloadFonts: false });

    const main = await createYrsSession({ clientId: 71001 });
    // A fresh client id, as the worker uses: it is a genuine peer, not a copy.
    const worker = await createYrsSession({ clientId: 71002 });
    try {
      main.seedFromDocx(bytes);
      const layoutInput = JSON.stringify({
        ...buildResidentRegionLayoutRequest(parsed, 24, RENDER_ENV),
        measurement: measurementFor(main.registerFont(font)),
      });

      const open = main.residentWorkerOpen();
      expect(open.state.byteLength).toBeGreaterThan(0);
      expect(open.fonts).toHaveLength(1);

      // Exactly what the worker's `open` handler does — no render inputs and
      // no measure inputs, because nobody has lowered or measured this yet.
      worker.loadState(open.state);
      worker.clearFonts();
      for (const bytes of open.fonts) worker.registerFont(bytes);
      worker.layoutDocumentWithRegionsVoid(layoutInput);

      const delta = decodeFrameDelta(worker.buildDisplayListFrame('{}', 0));
      expect(delta.full).toBe(true);
      const frame = applyFrameDelta(null, delta);
      expect(frame.displayList.pages.length).toBeGreaterThan(0);

      // The page geometry the shell now reads off the frame has to be the
      // geometry a pagination pass on this thread would have produced.
      const onMain = JSON.parse(main.layoutDocumentWithRegionsSlimJson(layoutInput)) as {
        layout: Layout;
      };
      expect(frame.displayList.pages.map((page) => page.height)).toEqual(
        onMain.layout.pages.map((page) => page.size.h)
      );
    } finally {
      main.destroy();
      worker.destroy();
    }
  });

  it('re-paginates from updates it was sent, without state being shipped again', async () => {
    const bytes = new Uint8Array(readFileSync(DOCX));
    const font = new Uint8Array(readFileSync(FONT));
    const parsed = await parseDocx(bytes.buffer as ArrayBuffer, { preloadFonts: false });

    const main = await createYrsSession({ clientId: 72001 });
    const worker = await createYrsSession({ clientId: 72002 });
    try {
      main.seedFromDocx(bytes);
      const layoutInput = JSON.stringify({
        ...buildResidentRegionLayoutRequest(parsed, 24, RENDER_ENV),
        measurement: measurementFor(main.registerFont(font)),
      });

      const open = main.residentWorkerOpen();
      worker.loadState(open.state);
      worker.clearFonts();
      for (const bytes of open.fonts) worker.registerFont(bytes);
      worker.layoutDocumentWithRegionsVoid(layoutInput);
      const before = applyFrameDelta(null, decodeFrameDelta(worker.buildDisplayListFrame('{}', 0)));

      // Enough text to change the page stack, delivered the only way a
      // worker-owned document receives one: as an update.
      const vector = worker.encodeStateVector();
      const paragraph = main.paragraphs('body')[0];
      main.insertText(
        { story: 'body', paraId: paragraph.paraId, offset: 0 },
        'lorem ipsum dolor sit amet '.repeat(400)
      );
      worker.applyUpdate(main.encodeStateAsUpdate(vector));

      // `relayout`: the region request and nothing else.
      worker.layoutDocumentWithRegionsVoid(layoutInput);
      const after = applyFrameDelta(
        before,
        decodeFrameDelta(worker.buildDisplayListFrame('{}', before.frameEpoch))
      );

      expect(after.displayList.pages.length).toBeGreaterThan(before.displayList.pages.length);
      const onMain = JSON.parse(main.layoutDocumentWithRegionsSlimJson(layoutInput)) as {
        layout: Layout;
      };
      expect(after.displayList.pages.map((page) => page.height)).toEqual(
        onMain.layout.pages.map((page) => page.size.h)
      );
    } finally {
      main.destroy();
      worker.destroy();
    }
  });
});
