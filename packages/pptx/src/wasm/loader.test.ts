import { afterAll, beforeAll, describe, expect, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import type {
  PresentationHandle,
  ShapePrimitive,
  StorySnapshot,
  TextBoxPrimitive,
  TextStyle,
  TextStyleSnapshot,
} from '../index';
import { initWasm, openPresentation } from '../index';

const root = resolve(import.meta.dir, '../../../..');
let handle: PresentationHandle;
let fixture: Uint8Array;

beforeAll(async () => {
  const [wasm, pptx, font] = await Promise.all([
    readFile(resolve(import.meta.dir, 'generated/pptx_wasm_bg.wasm')),
    readFile(resolve(root, 'apps/demo/public/betteroffice-demo.pptx')),
    readFile(resolve(root, 'crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf')),
  ]);
  await initWasm(wasm);
  fixture = pptx;
  handle = openPresentation(pptx, {
    clientId: 9001,
    fonts: [{ family: 'Liberation Sans', bytes: font }],
  });
});

afterAll(() => handle.dispose());

describe('PPTX wasm boundary', () => {
  test('opens shared updates without parsing the file bytes', () => {
    const source = openPresentation(fixture, { clientId: 9002 });
    const seed = source.encodeStateAsUpdate();
    const left = openPresentation(Uint8Array.of(0xff), {
      clientId: 9003,
      initialUpdate: seed,
    });
    const right = openPresentation(Uint8Array.of(0xff), {
      clientId: 9004,
      initialUpdate: seed,
    });

    expect(left.snapshot()).toEqual(source.snapshot());
    expect([...left.encodeStateAsUpdate()]).toEqual([...seed]);
    expect([...right.encodeStateAsUpdate()]).toEqual([...seed]);
    expect([...left.encodeStateVector()]).toEqual([...right.encodeStateVector()]);

    source.dispose();
    left.dispose();
    right.dispose();
  });

  test('opens, edits, reflows, hit-tests, and observes a local update', () => {
    const snapshot = handle.snapshot();
    expect(snapshot.slides.length).toBe(3);
    const story = firstStory(snapshot.slides.flatMap((slide) => slide.shapes));
    const insertion = story.length - 1;

    const events: Array<{ origin: string; update: Uint8Array }> = [];
    const unsubscribe = handle.onUpdate((update, origin) => events.push({ origin, update }));
    const receipt = handle.insertText(story.id, insertion, ' edited', {
      bold: true,
      fontSizePt: 28,
      color: '#325ee6',
    });
    expect(receipt.storyId).toBe(story.id);
    expect(handle.story(story.id).paragraphs.some((paragraph) =>
      paragraph.runs.some((run) => run.text.includes('edited'))
    )).toBe(true);

    const frame = handle.layoutSlide(0);
    const textBox = frame.primitives.find(
      (primitive): primitive is TextBoxPrimitive =>
        primitive.kind === 'textBox' && primitive.storyId === story.id
    );
    expect(textBox?.lines.length).toBeGreaterThan(0);
    const line = textBox!.lines[0];
    expect(handle.hitTest(line.x, line.y + line.height / 2)?.kind).toBe('text');

    expect(events[0]?.origin).toBe('local');
    expect(events[0]?.update.length).toBeGreaterThan(0);
    expect(handle.canUndo()).toBe(true);
    expect(handle.undo().applied).toBe(true);
    expect(handle.story(story.id).paragraphs.some((paragraph) =>
      paragraph.runs.some((run) => run.text.includes('edited'))
    )).toBe(false);
    unsubscribe();
  });

  test('saves a text edit into the file and refuses what it cannot write', () => {
    const editor = openPresentation(fixture, { clientId: 9101 });
    const snapshot = editor.snapshot();
    const story = firstStory(snapshot.slides.flatMap((slide) => slide.shapes));
    const paragraph = story.paragraphs[0];

    const untouched = openPresentation(editor.save(), { clientId: 9102 });
    expect(untouched.snapshot()).toEqual(snapshot);
    untouched.dispose();

    editor.insertText(story.id, 0, 'Saved: ', caretStyle(paragraph.runs[0].style));
    const saved = editor.save();
    const reopened = openPresentation(saved, { clientId: 9103 });
    expect(reopened.story(story.id).paragraphs[0].runs[0].text).toBe(
      `Saved: ${paragraph.runs[0].text}`
    );
    expect(reopened.snapshot().slides.length).toBe(snapshot.slides.length);
    reopened.dispose();

    editor.moveShape(snapshot.slides[0].id, snapshot.slides[0].shapes[0].id, 10, 20);
    expect(() => editor.save()).toThrow(/cannot save yet/);
    editor.dispose();
  });

  test('saves a paragraph split and a merge back into the file', () => {
    const editor = openPresentation(fixture, { clientId: 9201 });
    const story = firstStory(editor.snapshot().slides.flatMap((slide) => slide.shapes));
    const before = editor.story(story.id).paragraphs.length;
    const head = editor.story(story.id).paragraphs[0].runs[0].text.length;

    editor.insertParagraphBreak(story.id, head);
    const split = openPresentation(editor.save(), { clientId: 9202 });
    expect(split.story(story.id).paragraphs.length).toBe(before + 1);
    split.dispose();

    editor.deleteParagraphBreak(story.id, head);
    const merged = openPresentation(editor.save(), { clientId: 9203 });
    expect(merged.story(story.id).paragraphs.length).toBe(before);
    expect(plainText(merged.story(story.id))).toBe(plainText(editor.story(story.id)));
    merged.dispose();
    editor.dispose();
  });

  test('inserts and styles preset shapes with undo and redo', () => {
    const slide = handle.snapshot().slides[0];
    const receipt = handle.addShape(slide.id, {
      name: 'Styled rounded rectangle',
      geometry: 'roundRect',
      rect: { x: 900_000, y: 1_000_000, width: 3_100_000, height: 1_400_000 },
      fill: '#D9EAF7',
    });
    expect(shapeSnapshot(receipt.shapeId).geometry).toBe('roundRect');
    expect(handle.undo().snapshot.slides[0].shapes.some(
      (shape) => shape.id === receipt.shapeId
    )).toBe(false);
    expect(handle.redo().snapshot.slides[0].shapes.some(
      (shape) => shape.id === receipt.shapeId
    )).toBe(true);

    handle.setShapeFill(slide.id, receipt.shapeId, '#3367D6');
    expect(shapeSnapshot(receipt.shapeId).fill?.color?.rgb).toBe('3367D6');
    expect(handle.undo().applied).toBe(true);
    expect(shapeSnapshot(receipt.shapeId).fill?.color?.rgb).toBe('D9EAF7');
    expect(handle.redo().applied).toBe(true);

    handle.setShapeStroke(slide.id, receipt.shapeId, {
      color: '#EA4335',
      widthPt: 3,
    });
    expect(shapeSnapshot(receipt.shapeId).outline?.width).toBe(38_100);
    expect(handle.undo().applied).toBe(true);
    expect(shapeSnapshot(receipt.shapeId).outline).toBeNull();
    expect(handle.redo().applied).toBe(true);

    handle.setShapeAdjust(slide.id, receipt.shapeId, { adj: 0.32 });
    expect(shapeSnapshot(receipt.shapeId).adjustValues.adj).toBe(0.32);
    expect(handle.undo().applied).toBe(true);
    expect(shapeSnapshot(receipt.shapeId).adjustValues.adj).toBeCloseTo(0.16667);
    expect(handle.redo().applied).toBe(true);

    const primitive = handle.layoutSlide(0).primitives.find(
      (candidate): candidate is ShapePrimitive =>
        candidate.kind === 'shape' && candidate.shapeId === receipt.shapeId
    );
    expect(primitive?.fill).toEqual({ kind: 'solid', color: '#3367D6' });
    expect(primitive?.stroke?.color).toBe('#EA4335');
    expect(primitive?.adjustValues?.adj).toBeCloseTo(0.32);

    handle.setShapeFill(slide.id, receipt.shapeId, null);
    handle.setShapeStroke(slide.id, receipt.shapeId, {});
    const cleared = handle.layoutSlide(0).primitives.find(
      (candidate): candidate is ShapePrimitive =>
        candidate.kind === 'shape' && candidate.shapeId === receipt.shapeId
    );
    expect(cleared?.fill).toBeUndefined();
    expect(cleared?.stroke).toBeUndefined();
  });
});

/**
 * The style a caret carries inside a run. Typing with the run's own style is
 * what keeps an edit inside one run, which is all `save()` can project today.
 */
function caretStyle(style: TextStyleSnapshot): TextStyle {
  return {
    bold: style.bold ?? undefined,
    italic: style.italic ?? undefined,
    fontSizePt: style.fontSizePt ?? undefined,
    color: style.color ?? undefined,
    fontFamily: style.fontFamily ?? undefined,
    underline: style.underline ?? undefined,
  };
}

function plainText(story: StorySnapshot): string {
  return story.paragraphs
    .map((paragraph) => paragraph.runs.map((run) => run.text).join(''))
    .join('\n');
}

function shapeSnapshot(shapeId: string) {
  const shape = handle.snapshot().slides[0].shapes.find((candidate) => candidate.id === shapeId);
  if (!shape) throw new Error(`shape ${shapeId} was not found`);
  return shape;
}

function firstStory(shapes: Array<{ textStories: StorySnapshot[]; children: unknown[] }>): StorySnapshot {
  for (const shape of shapes) {
    if (shape.textStories[0]) return shape.textStories[0];
  }
  throw new Error('fixture has no text story');
}
