import { beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { parseDocx } from '../docx';
import { writeDocumentWithRust } from '../docx/rustSaveFacade';
import { preloadEditWasm } from '../wasm/edit';
import { createYrsSession, type YrsStorySegment } from './index';
import { documentToYrs } from './documentToYrs';
import { yrsToDocument } from './yrsToDocument';
import type { Document } from '../types';

const WASM = resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm');
const FIXTURE = resolve(import.meta.dir, '../../../../apps/demo/public/betteroffice-demo.docx');

/** The payload the editor's "insert page number" sends for a `field` embed. */
const PAGE_FIELD = {
  instruction: 'PAGE',
  fieldType: 'PAGE',
  displayText: '1',
  displayMode: 'result',
} as const;

/** End of the last paragraph's content — the same arithmetic the editor uses. */
function endOfContent(segments: YrsStorySegment[]): number {
  let end = 0;
  for (const segment of segments) {
    end += segment.kind === 'text' ? segment.text.length : 1;
  }
  if (segments.at(-1)?.kind === 'pilcrow') end -= 1;
  return Math.max(0, end);
}

/**
 * Depth-first search that follows Map values. `package.headers` and
 * `package.footers` are Maps, which both `JSON.stringify` and
 * `Object.values` silently skip — an earlier probe concluded from that
 * artifact that header edits were being dropped.
 */
function deepFind(value: unknown, pred: (n: any) => boolean, found: any[] = [], seen = new Set()): any[] {
  if (value === null || typeof value !== 'object' || seen.has(value)) return found;
  seen.add(value);
  if (pred(value as object)) found.push(value);
  const children =
    value instanceof Map ? [...value.values()] : Array.isArray(value) ? value : Object.values(value);
  for (const child of children) deepFind(child, pred, found, seen);
  return found;
}

// The serializer is free to write a PAGE field in either OOXML spelling —
// `fldSimple` or the `fldChar begin/separate/end` run sequence Word itself
// prefers — and the parser reads those back as simpleField and complexField
// respectively. The feature is the field, not the spelling.
const pageFields = (doc: Document) =>
  deepFind(
    doc.package.footers,
    (n) =>
      (n.type === 'simpleField' || n.type === 'complexField') &&
      typeof n.instruction === 'string' &&
      n.instruction.trim().startsWith('PAGE')
  );

describe('inserting a page number field', () => {
  beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(WASM))));

  async function withFieldInFooter(): Promise<{ original: ArrayBuffer; document: Document }> {
    const bytes = Uint8Array.from(readFileSync(FIXTURE));
    const original = bytes.buffer.slice(0) as ArrayBuffer;
    const parsed = await parseDocx(bytes.buffer);
    const session = await createYrsSession({ clientId: 81001 });
    documentToYrs(session, parsed);
    const story = 'hf:rId4';
    expect(session.storyIds()).toContain(story);
    session.applyRawOps(story, [
      {
        op: 'insertEmbed',
        index: endOfContent(session.storySegments(story)),
        kind: 'field',
        payload: { ...PAGE_FIELD },
      },
    ]);
    return { original, document: yrsToDocument(session, parsed) };
  }

  it('projects the field into the footer, after the existing text', async () => {
    const { document } = await withFieldInFooter();
    const fields = pageFields(document);
    expect(fields).toHaveLength(1);
    const texts = deepFind(
      document.package.footers,
      (n) => n.type === 'text' && typeof n.text === 'string'
    ).map((n) => n.text);
    expect(texts.at(-1)).toBe('1');
    expect(texts.join('')).toContain('Created with BetterOffice');
  });

  // A field parsed from a file is written back from its preserved source; a
  // field created here has no source and is built from the model. Those are
  // different write paths, so the only honest check is real bytes.
  it('survives a save and a reparse', async () => {
    const { original, document } = await withFieldInFooter();
    const saved = await writeDocumentWithRust(document, original);
    const reparsed = await parseDocx(saved.buffer);
    const fields = pageFields(reparsed);
    expect(fields).toHaveLength(1);
    expect(fields[0]?.fieldType).toBe('PAGE');
  });

  // The editor's other path: no footer exists, so one is materialised with
  // the field already inside. The document-model shape it builds must be one
  // the serializer can express and the parser reads back.
  it('writes a materialised footer holding the field', async () => {
    const bytes = Uint8Array.from(readFileSync(FIXTURE));
    const parsed = await parseDocx(bytes.buffer);
    const footers = new Map(parsed.package.footers ?? []);
    footers.set('rId4', {
      type: 'footer',
      hdrFtrType: 'default',
      content: [
        {
          type: 'paragraph',
          formatting: { alignment: 'center' },
          content: [
            {
              type: 'simpleField',
              instruction: 'PAGE',
              fieldType: 'PAGE',
              content: [{ type: 'run', content: [{ type: 'text', text: '1' }] }],
            },
          ],
        },
      ],
    } as never);
    const edited = { ...parsed, package: { ...parsed.package, footers } } as Document;

    const saved = await writeDocumentWithRust(edited, bytes.buffer.slice(0) as ArrayBuffer);
    const reparsed = await parseDocx(saved.buffer);
    expect(pageFields(reparsed)).toHaveLength(1);
  });

  it('leaves an untouched document without one', async () => {
    const parsed = await parseDocx(Uint8Array.from(readFileSync(FIXTURE)).buffer);
    expect(pageFields(parsed)).toHaveLength(0);
  });
});
