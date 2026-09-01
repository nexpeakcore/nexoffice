import { beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { parseDocx } from '../docx';
import { buildResidentRegionLayoutRequest } from '../editor/computeLayout';
import { applyFrameDelta, decodeFrameDelta } from '../layout/render/frameDelta';
import { preloadEditWasm } from '../wasm/edit';
import { createYrsSession } from './index';

const WASM = resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm');
const DOCX = resolve(import.meta.dir, '__fixtures__/titlepg-first-footer.docx');
const FONT = resolve(import.meta.dir, '../../../fonts/assets/LiberationSans-Regular.ttf');

describe('typing through a page boundary', () => {
  beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(WASM))));

  /**
   * Incremental pagination can land on a different page count — that is what
   * typing text onto a new page is. The incremental display update cannot
   * follow that, and used to say so by failing, which drops the frame and
   * sends the host back to paginating the document on its own thread.
   */
  it('keeps building frames as the page count grows', async () => {
    const bytes = new Uint8Array(readFileSync(DOCX));
    const parsed = await parseDocx(bytes.buffer as ArrayBuffer, { preloadFonts: false });
    const session = await createYrsSession({ clientId: 61001 });
    try {
      session.seedFromDocx(bytes);
      const layoutInput = JSON.stringify({
        ...buildResidentRegionLayoutRequest(parsed, 24, { themeColors: {}, numericIds: {} }),
        measurement: {
          fontChains: {
            'liberation sans|0|0': [session.registerFont(new Uint8Array(readFileSync(FONT)))],
          },
          defaults: { fontSize: 24, fontFamily: 'Liberation Sans' },
          authoritativeShaping: true,
        },
      });
      session.layoutDocumentWithRegionsVoid(layoutInput);

      let frame = applyFrameDelta(null, decodeFrameDelta(session.buildDisplayListFrame('{}', 0)));
      const startPages = frame.displayList.pages.length;
      const paragraph = session.paragraphs('body')[0];
      session.setSelection({ story: 'body', paraId: paragraph.paraId, offset: 0 });

      // One character at a time, as a person types. The first page boundary on
      // this fixture arrives inside 200 keystrokes.
      for (let typed = 0; typed < 200; typed += 1) {
        frame = applyFrameDelta(frame, decodeFrameDelta(session.applyInput('x', frame.frameEpoch)));
      }

      expect(frame.displayList.pages.length).toBeGreaterThan(startPages);
    } finally {
      session.destroy();
    }
  }, 30_000);
});
