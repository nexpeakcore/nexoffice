import { beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { parseDocx } from '../docx';
import { buildResidentRegionLayoutRequest } from './computeLayout';
import type { Layout } from '../layout/pagination';
import { preloadEditWasm } from '../wasm/edit';
import { createYrsSession } from '../yrs/index';

const WASM = resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm');
const DOCX = resolve(import.meta.dir, '../yrs/__fixtures__/titlepg-first-footer.docx');
const FONT = resolve(import.meta.dir, '../../../fonts/assets/LiberationSans-Regular.ttf');

describe('the region request a document asks for', () => {
  beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(WASM))));

  /**
   * The last modelled section already carries the body's final section
   * properties, so listing them again describes a section that owns no
   * content. It changes no page, but it does make the document look
   * multi-section to the engine, which turns off the region fast path — and
   * every keystroke then repaginates the whole document.
   */
  it('does not repeat the last section', async () => {
    const bytes = new Uint8Array(readFileSync(DOCX));
    const parsed = await parseDocx(bytes.buffer as ArrayBuffer, { preloadFonts: false });
    const request = buildResidentRegionLayoutRequest(parsed, 24, {
      themeColors: {},
      numericIds: {},
    });

    expect(parsed.package.document.sections?.length).toBe(1);
    expect(request.regions.sections).toHaveLength(1);
  });

  /** The section that was being repeated owned no page, and still owns none. */
  it('lays the document out the same as the repeated form did', async () => {
    const bytes = new Uint8Array(readFileSync(DOCX));
    const font = new Uint8Array(readFileSync(FONT));
    const parsed = await parseDocx(bytes.buffer as ArrayBuffer, { preloadFonts: false });
    const session = await createYrsSession({ clientId: 74001 });
    try {
      session.seedFromDocx(bytes);
      const base = buildResidentRegionLayoutRequest(parsed, 24, {
        themeColors: {},
        numericIds: {},
      });
      const measurement = {
        fontChains: { 'liberation sans|0|0': [session.registerFont(font)] },
        defaults: { fontSize: 24, fontFamily: 'Liberation Sans' },
        authoritativeShaping: true,
      };
      const layoutOf = (sections: (typeof base)['regions']['sections']): Layout =>
        (
          JSON.parse(
            session.layoutDocumentWithRegionsSlimJson(
              JSON.stringify({
                ...base,
                regions: { ...base.regions, sections },
                measurement,
              })
            )
          ) as { layout: Layout }
        ).layout;

      const repeated = [
        ...base.regions.sections,
        {
          sectionId: parsed.package.document.finalSectionProperties?.sectionId,
          properties: parsed.package.document.finalSectionProperties ?? {},
        },
      ];

      expect(layoutOf(base.regions.sections)).toEqual(layoutOf(repeated));
    } finally {
      session.destroy();
    }
  }, 30_000);
});
