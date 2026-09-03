import { beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { parseDocx } from '../docx';
import { preloadEditWasm } from '../wasm/edit';
import { createYrsSession } from './index';
import { documentToYrs } from './documentToYrs';
import type { Document } from '../types';

const WASM = resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm');
const FIXTURE = resolve(import.meta.dir, '../../../../apps/demo/public/betteroffice-demo.docx');

const BOOKMARK = { type: 'bookmarkStart', id: '7', name: 'chapter-two' } as const;
const EMU_PER_PIXEL = 9525;
const IMAGE_RUN = {
  type: 'run',
  content: [
    {
      type: 'drawing',
      image: {
        type: 'image',
        rId: 'rId99',
        src: 'data:image/png;base64,' + 'A'.repeat(4096),
        size: { width: 120 * EMU_PER_PIXEL, height: 80 * EMU_PER_PIXEL },
        wrap: { type: 'inline' },
      },
    },
  ],
} as const;

/** The fixture with a bookmark and an image spliced into its first paragraph. */
async function seededSession(clientId: number) {
  const document = (await parseDocx(
    new Uint8Array(readFileSync(FIXTURE)).buffer as ArrayBuffer,
    { preloadFonts: false }
  )) as Document;
  const paragraph = document.package.document.content.find(
    (block): block is Extract<(typeof block), { type: 'paragraph' }> => block.type === 'paragraph'
  );
  if (!paragraph) throw new Error('fixture has no paragraph to anchor to');
  paragraph.content = [BOOKMARK, IMAGE_RUN, ...paragraph.content] as typeof paragraph.content;

  const session = await createYrsSession({ clientId });
  documentToYrs(session, document);
  return session;
}

describe('storyOutline', () => {
  beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(WASM))));

  it('agrees with storySegments on every position it describes', async () => {
    const session = await seededSession(78001);
    const segments = session.storySegments('body');
    const outline = session.storyOutline('body');

    expect(outline.length).toBe(segments.length);
    outline.forEach((entry, index) => {
      const segment = segments[index];
      expect(entry.kind).toBe(segment.kind);
      if (entry.kind === 'text' && segment.kind === 'text') {
        expect(entry.len).toBe(segment.text.length);
      }
      if (entry.kind === 'pilcrow' && segment.kind === 'pilcrow') {
        expect(entry.paraId).toBe(segment.paraId);
      }
      if (entry.kind === 'embed' && segment.kind === 'embed') {
        expect(entry.embedKind).toBe(segment.embedKind);
      }
    });
    session.destroy();
  }, 120_000);

  it('keeps a bookmark, which is itself a position', async () => {
    const session = await seededSession(78002);
    const carried = session
      .storyOutline('body')
      .flatMap((entry) => (entry.kind === 'pilcrow' && entry.bookmarks ? [entry.bookmarks] : []));

    expect(JSON.stringify(carried)).toContain(BOOKMARK.name);
    session.destroy();
  }, 120_000);

  it('leaves the image bytes behind but keeps what places the image', async () => {
    const session = await seededSession(78003);
    const images = session
      .storyOutline('body')
      .flatMap((entry) => (entry.kind === 'embed' && entry.embedKind === 'image' ? [entry] : []));

    expect(images.length).toBeGreaterThan(0);
    for (const image of images) {
      expect(image.payload.src).toBeUndefined();
      expect(image.payload.rId).toBe('rId99');
      expect(image.payload.width).toBe(120);
      expect(image.payload.height).toBe(80);
    }
    expect(JSON.stringify(session.storyOutline('body'))).not.toContain('AAAAAAAA');
    session.destroy();
  }, 120_000);
});
