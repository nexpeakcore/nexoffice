import { describe, expect, test } from 'bun:test';
import { Window } from 'happy-dom';
import { buildMirrorPageOutline } from './mirrorDom';
import type { DisplayPage } from './displayList';

const page = {
  pageIndex: 3,
  width: 816,
  height: 1056,
  primitives: [
    { kind: 'text', blockId: 1, text: 'Hello ', x: 10, y: 10, w: 40, h: 12, font: '12px x', docStart: 100, docEnd: 106, logicalOrder: 0 },
    { kind: 'text', blockId: 1, text: 'world', x: 50, y: 10, w: 30, h: 12, font: '12px x', docStart: 106, docEnd: 111, logicalOrder: 1 },
    { kind: 'text', blockId: 2, text: '1.', x: 10, y: 30, w: 8, h: 12, font: '12px x', listMarker: true, docStart: 120, docEnd: 120 },
    { kind: 'text', blockId: 2, text: 'Second', x: 20, y: 30, w: 35, h: 12, font: '12px x', docStart: 120, docEnd: 126 },
    { kind: 'line', x: 0, y: 50, w: 100, h: 1 },
  ],
} as unknown as DisplayPage;

const richPage = {
  pageIndex: 4,
  width: 816,
  height: 1056,
  primitives: [
    { kind: 'text', blockId: 1, text: 'see ', x: 10, y: 10, w: 20, h: 12, font: '12px x', docStart: 200, docEnd: 204, logicalOrder: 0 },
    { kind: 'text', blockId: 1, text: 'the docs', href: 'https://example.com', x: 30, y: 10, w: 40, h: 12, font: '12px x', docStart: 204, docEnd: 212, logicalOrder: 1 },
    { kind: 'image', blockId: 2, relId: 'rId7', altText: 'Architecture diagram', x: 10, y: 40, w: 200, h: 100, docStart: 220, docEnd: 221 },
    { kind: 'image', blockId: 3, relId: 'rId8', decorative: true, x: 10, y: 160, w: 50, h: 20, docStart: 230, docEnd: 231 },
  ],
} as unknown as DisplayPage;

const regionsPage = {
  pageIndex: 5,
  width: 816,
  height: 1056,
  primitives: [
    { kind: 'text', blockId: 1, text: 'Body line', x: 10, y: 200, w: 50, h: 12, font: '12px x', docStart: 300, docEnd: 309 },
  ],
  header: { kind: 'header', rId: 'rId3', y: 40, height: 40, primitives: [
    { kind: 'text', blockId: 9, text: 'Chapter heading', x: 10, y: 45, w: 80, h: 12, font: '12px x', docStart: 1, docEnd: 16 },
  ] },
  footer: { kind: 'footer', rId: 'rId4', y: 980, height: 40, primitives: [
    { kind: 'text', blockId: 8, text: 'Page 6', x: 10, y: 985, w: 30, h: 12, font: '12px x', docStart: 1, docEnd: 7 },
  ] },
  noteAreas: [
    { kind: 'footnote', noteIds: [1], primitives: [
      { kind: 'text', blockId: 7, text: 'A cited footnote', x: 10, y: 900, w: 80, h: 12, font: '12px x', docStart: 1, docEnd: 17 },
    ] },
  ],
} as unknown as DisplayPage;

describe('buildMirrorPageOutline', () => {
  const window = new Window();
  const doc = window.document as unknown as Document;

  test('emits one paragraph per block with text in reading order', () => {
    const el = buildMirrorPageOutline(page, { document: doc });
    const paragraphs = [...el.querySelectorAll('.layout-paragraph')];

    expect(paragraphs.map((p) => p.textContent)).toEqual(['Hello world', 'Second']);
    expect(paragraphs[0]?.getAttribute('data-doc-start')).toBe('100');
    expect(paragraphs[0]?.getAttribute('data-doc-end')).toBe('111');
  });

  test('stays far smaller than the positioned mirror', () => {
    const el = buildMirrorPageOutline(page, { document: doc });

    expect(el.querySelectorAll('*').length).toBeLessThan(6);
    expect(el.getAttribute('data-mirror-outline')).toBe('true');
  });

  test('carries the page label so a screen reader can announce off-window pages', () => {
    const el = buildMirrorPageOutline(page, { document: doc, labels: { page: 'Page 4' } });

    expect(el.getAttribute('aria-label')).toBe('Page 4');
    expect(el.getAttribute('role')).toBe('document');
  });
  test('keeps images announceable on an image-only page', () => {
    const el = buildMirrorPageOutline(richPage, { document: doc });
    const img = el.querySelector('[data-rel-id="rId7"]');

    expect(img?.getAttribute('role')).toBe('img');
    expect(img?.getAttribute('aria-label')).toBe('Architecture diagram');
  });

  test('hides decorative images from the accessibility tree', () => {
    const el = buildMirrorPageOutline(richPage, { document: doc });
    const decorative = el.querySelector('[data-rel-id="rId8"]');

    expect(decorative?.getAttribute('aria-hidden')).toBe('true');
    expect(decorative?.hasAttribute('role')).toBe(false);
  });

  test('keeps hyperlinks activatable instead of flattening them to text', () => {
    const el = buildMirrorPageOutline(richPage, { document: doc });
    const link = el.querySelector('a');

    expect(link?.getAttribute('href')).toBe('https://example.com');
    expect(link?.textContent).toBe('the docs');
    expect(el.querySelector('.layout-paragraph')?.textContent).toBe('see the docs');
  });
  test('keeps footnote text in the outline', () => {
    const el = buildMirrorPageOutline(regionsPage, { document: doc });
    const notes = el.querySelector('.layout-page-notes');

    expect(notes?.getAttribute('data-note-kind')).toBe('footnote');
    expect(notes?.textContent).toBe('A cited footnote');
  });

  test('keeps header and footer content in the outline', () => {
    const el = buildMirrorPageOutline(regionsPage, {
      document: doc,
      labels: { header: 'Header', footer: 'Footer' },
    });
    const header = el.querySelector('.layout-page-header');
    const footer = el.querySelector('.layout-page-footer');

    expect(header?.textContent).toBe('Chapter heading');
    expect(header?.getAttribute('data-hf-rid')).toBe('rId3');
    expect(header?.getAttribute('aria-label')).toBe('Header');
    expect(footer?.textContent).toBe('Page 6');
  });

  test('omits regions a page does not have', () => {
    const el = buildMirrorPageOutline(page, { document: doc });

    expect(el.querySelector('.layout-page-header')).toBeNull();
    expect(el.querySelector('.layout-page-notes')).toBeNull();
  });
});
