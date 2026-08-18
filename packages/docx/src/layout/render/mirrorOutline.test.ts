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
});
