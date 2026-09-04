import { describe, expect, test } from 'bun:test';
import {
  completesAPartialOpen,
  openRequestFor,
  residentLayoutRequest,
} from './residentLayoutPasses';

const engineOpen = {
  state: new Uint8Array([1]),
  selection: null,
  fonts: [],
  fontsRevision: 0,
};

describe('openRequestFor', () => {
  test('asks for the document’s first pages, not all of them', () => {
    // Without this the worker measures and paginates every block before the
    // first one is painted: 2145ms of a 2.2s open on a 343-page document.
    const request = openRequestFor(engineOpen);
    expect(request.firstBlocks).toBeGreaterThan(0);
    expect(request.pageWindow?.count).toBeGreaterThan(0);
  });

  test('carries the state and fonts it was given', () => {
    expect(openRequestFor(engineOpen).state).toBe(engineOpen.state);
    expect(openRequestFor(engineOpen).fonts).toBe(engineOpen.fonts);
  });
});

describe('residentLayoutRequest', () => {
  const input = '{"regions":{}}';

  test('opens a document the worker does not hold yet', () => {
    expect(residentLayoutRequest(true, null, input)).toBe('open');
  });

  test('rebuilds the frame when the region request has not moved', () => {
    expect(residentLayoutRequest(false, input, input)).toBe('frame');
  });

  test('lays the rest of the document out after a partial open', () => {
    // The open cleared the remembered input, which is the only thing that
    // asks for the document past the prefix it painted.
    expect(residentLayoutRequest(false, null, input)).toBe('relayout');
    expect(completesAPartialOpen(null)).toBe(true);
  });

  test('repaginates when the host changed the region request', () => {
    expect(residentLayoutRequest(false, '{"regions":{"a":1}}', input)).toBe('relayout');
    expect(completesAPartialOpen('{"regions":{"a":1}}')).toBe(false);
  });
});
