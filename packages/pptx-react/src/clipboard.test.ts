import { describe, expect, it } from 'bun:test';
import {
  copyTextSelection,
  createClipboardQueue,
  cutTextSelection,
  deleteTextSelection,
  limitPasteText,
  PASTE_CHARACTER_LIMIT,
  PASTE_LINE_LIMIT,
  pasteTextSelection,
  sameTextSelection,
  textRangeOf,
  type ClipboardHost,
  type ClipboardTextSelection,
} from './clipboard';

const STORY = 'story-1';

function range(anchor: number, focus: number): ClipboardTextSelection {
  return { shapeId: 'shape-1', storyId: STORY, anchor, focus };
}

interface Deferred<T> {
  promise: Promise<T>;
  resolve: (value: T) => void;
  reject: (reason?: unknown) => void;
}

function deferred<T>(): Deferred<T> {
  let resolve!: (value: T) => void;
  let reject!: (reason?: unknown) => void;
  const promise = new Promise<T>((res, rej) => {
    resolve = res;
    reject = rej;
  });
  return { promise, resolve, reject };
}

/**
 * A story whose index space matches the engine's: paragraph breaks are one
 * character, so a paste's ops can be replayed into a plain string.
 */
function createHost(initial: {
  text: string;
  selection: ClipboardTextSelection | null;
  read?: () => Promise<string>;
  write?: (text: string) => Promise<void>;
}) {
  const state = {
    text: initial.text,
    selection: initial.selection,
    written: [] as string[],
    commits: [] as (ClipboardTextSelection | null)[],
    errors: [] as unknown[],
    ops: [] as string[],
    pasteLimits: [] as number[],
  };
  const host: ClipboardHost = {
    selection: () => state.selection,
    storyText: (storyId) => {
      expect(storyId).toBe(STORY);
      return state.text;
    },
    deleteText: (_storyId, start, end) => {
      state.ops.push(`delete(${start},${end})`);
      state.text = state.text.slice(0, start) + state.text.slice(end);
    },
    insertText: (_storyId, index, text) => {
      state.ops.push(`insert(${index},${text})`);
      state.text = state.text.slice(0, index) + text + state.text.slice(index);
    },
    insertParagraphBreak: (_storyId, index) => {
      state.ops.push(`break(${index})`);
      state.text = `${state.text.slice(0, index)}\n${state.text.slice(index)}`;
    },
    commit: (selection) => {
      state.commits.push(selection);
      state.selection = selection;
    },
    readClipboard: () => {
      state.ops.push('read');
      return (initial.read ?? (() => Promise.resolve('')))();
    },
    writeClipboard: (text) => {
      state.ops.push('write');
      return (initial.write ?? ((value: string) => {
        state.written.push(value);
        return Promise.resolve();
      }))(text);
    },
    reportError: (value) => state.errors.push(value),
    reportPasteLimit: (dropped) => state.pasteLimits.push(dropped),
  };
  return { host, state };
}

/** The ops a paste actually sent to the engine, with the clipboard read dropped. */
function engineOps(ops: readonly string[]): string[] {
  return ops.filter((op) => op !== 'read' && op !== 'write');
}

describe('pptx clipboard ranges', () => {
  it('orders a selection into a range and compares selections by value', () => {
    expect(textRangeOf(range(7, 2))).toEqual({ start: 2, end: 7 });
    expect(textRangeOf(null)).toBeNull();
    expect(sameTextSelection(range(2, 7), range(2, 7))).toBe(true);
    expect(sameTextSelection(range(2, 7), range(2, 8))).toBe(false);
    expect(sameTextSelection(range(2, 7), null)).toBe(false);
    expect(sameTextSelection(null, null)).toBe(true);
  });
});

describe('pptx clipboard copy and cut', () => {
  it('copies the selected range', async () => {
    const { host, state } = createHost({ text: 'Hello world', selection: range(0, 5) });

    expect(await copyTextSelection(host)).toBe(true);
    expect(state.written).toEqual(['Hello']);
    expect(state.text).toBe('Hello world');
  });

  it('copies nothing for a collapsed caret', async () => {
    const { host, state } = createHost({ text: 'Hello world', selection: range(3, 3) });

    expect(await copyTextSelection(host)).toBe(false);
    expect(state.ops).toEqual([]);
  });

  it('deletes the range once the clipboard has taken the text', async () => {
    const { host, state } = createHost({ text: 'Hello world', selection: range(0, 6) });

    await cutTextSelection(host);

    expect(state.written).toEqual(['Hello ']);
    expect(state.text).toBe('world');
    expect(state.commits).toEqual([range(0, 0)]);
  });

  it('keeps the text when the clipboard write fails', async () => {
    const { host, state } = createHost({
      text: 'Hello world',
      selection: range(0, 6),
      write: () => Promise.reject(new Error('clipboard permission denied')),
    });

    await cutTextSelection(host);

    expect(state.text).toBe('Hello world');
    expect(state.ops).toEqual(['write']);
    expect(state.commits).toEqual([]);
    expect(state.errors).toEqual([]);
  });

  it('abandons a cut whose caret moved while the clipboard was writing', async () => {
    const pending = deferred<void>();
    const { host, state } = createHost({
      text: 'Hello world',
      selection: range(0, 6),
      write: () => pending.promise,
    });

    const cut = cutTextSelection(host);
    state.selection = range(6, 11);
    pending.resolve();
    await cut;

    expect(state.text).toBe('Hello world');
    expect(state.commits).toEqual([]);
  });

  it('deletes the live range and collapses the caret', () => {
    const { host, state } = createHost({ text: 'Hello world', selection: range(11, 5) });

    deleteTextSelection(host);

    expect(state.text).toBe('Hello');
    expect(state.commits).toEqual([range(5, 5)]);
  });
});

describe('pptx clipboard paste', () => {
  it('replaces the selection and breaks lines into paragraphs', async () => {
    const { host, state } = createHost({
      text: 'Hello world',
      selection: range(6, 11),
      read: () => Promise.resolve('one\ntwo'),
    });

    await pasteTextSelection(host);

    expect(state.text).toBe('Hello one\ntwo');
    expect(state.ops).toEqual(['read', 'delete(6,11)', 'insert(6,one)', 'break(9)', 'insert(10,two)']);
    expect(state.commits).toEqual([range(13, 13)]);
  });

  it('drops a paste whose selection moved while the clipboard was read', async () => {
    const pending = deferred<string>();
    const { host, state } = createHost({
      text: 'Hello world',
      selection: range(0, 5),
      read: () => pending.promise,
    });

    const paste = pasteTextSelection(host);
    state.selection = range(6, 11);
    pending.resolve('pasted');
    await paste;

    expect(state.text).toBe('Hello world');
    expect(state.ops).toEqual(['read']);
    expect(state.commits).toEqual([]);
  });

  it('drops a paste whose story went away while the clipboard was read', async () => {
    const pending = deferred<string>();
    const { host, state } = createHost({
      text: 'Hello world',
      selection: range(0, 5),
      read: () => pending.promise,
    });

    const paste = pasteTextSelection(host);
    state.selection = null;
    pending.resolve('pasted');
    await paste;

    expect(state.text).toBe('Hello world');
    expect(state.commits).toEqual([]);
  });

  it('ignores an empty clipboard', async () => {
    const { host, state } = createHost({
      text: 'Hello world',
      selection: range(0, 5),
      read: () => Promise.resolve(''),
    });

    await pasteTextSelection(host);

    expect(state.text).toBe('Hello world');
    expect(state.ops).toEqual(['read']);
  });
});

describe('pptx clipboard paste limits', () => {
  it('leaves a paste under the caps byte for byte alone', async () => {
    const { host, state } = createHost({
      text: '',
      selection: range(0, 0),
      read: () => Promise.resolve('Héllo, wörld! \t— ok\r\nsecond\tline'),
    });

    await pasteTextSelection(host);

    expect(state.ops).toEqual([
      'read',
      'insert(0,Héllo, wörld! \t— ok)',
      'break(19)',
      'insert(20,second\tline)',
    ]);
    expect(state.pasteLimits).toEqual([]);
    expect(state.errors).toEqual([]);
  });

  it('pastes text sitting exactly on the character cap whole and says nothing', async () => {
    const clipboard = 'a'.repeat(PASTE_CHARACTER_LIMIT);
    const { host, state } = createHost({
      text: '',
      selection: range(0, 0),
      read: () => Promise.resolve(clipboard),
    });

    await pasteTextSelection(host);

    expect(state.text).toBe(clipboard);
    expect(state.pasteLimits).toEqual([]);
    expect(state.commits).toEqual([range(PASTE_CHARACTER_LIMIT, PASTE_CHARACTER_LIMIT)]);
  });

  it('drops the one character past the cap and reports exactly that', async () => {
    const clipboard = 'a'.repeat(PASTE_CHARACTER_LIMIT + 1);
    const { host, state } = createHost({
      text: '',
      selection: range(0, 0),
      read: () => Promise.resolve(clipboard),
    });

    await pasteTextSelection(host);

    expect(state.text).toBe('a'.repeat(PASTE_CHARACTER_LIMIT));
    expect(state.pasteLimits).toEqual([1]);
    expect(state.commits).toEqual([range(PASTE_CHARACTER_LIMIT, PASTE_CHARACTER_LIMIT)]);
  });

  it('bounds the ops a multi-megabyte clipboard can turn into', async () => {
    const clipboard = Array.from({ length: 60_000 }, () => 'lorem ipsum dolor sit').join('\n');
    const { host, state } = createHost({
      text: '',
      selection: range(0, 0),
      read: () => Promise.resolve(clipboard),
    });

    await pasteTextSelection(host);

    expect(clipboard.length).toBeGreaterThan(1_000_000);
    expect(engineOps(state.ops).length).toBeLessThanOrEqual(PASTE_LINE_LIMIT * 2);
    expect(state.text.length).toBeLessThanOrEqual(PASTE_CHARACTER_LIMIT);
    expect(state.pasteLimits).toEqual([clipboard.length - state.text.length]);
    expect(state.errors).toEqual([]);
  });

  it('caps by line count even when the text is well under the character cap', async () => {
    const clipboard = '\n'.repeat(PASTE_LINE_LIMIT + 100);
    const { host, state } = createHost({
      text: '',
      selection: range(0, 0),
      read: () => Promise.resolve(clipboard),
    });

    await pasteTextSelection(host);

    expect(clipboard.length).toBeLessThan(PASTE_CHARACTER_LIMIT);
    expect(engineOps(state.ops).length).toBe(PASTE_LINE_LIMIT - 1);
    expect(state.text).toBe('\n'.repeat(PASTE_LINE_LIMIT - 1));
    expect(state.pasteLimits).toEqual([clipboard.length - (PASTE_LINE_LIMIT - 1)]);
  });

  it('takes the stranded carriage return when the cap lands inside a CRLF', async () => {
    const clipboard = `${'a'.repeat(PASTE_CHARACTER_LIMIT - 1)}\r\ntail`;
    const { host, state } = createHost({
      text: '',
      selection: range(0, 0),
      read: () => Promise.resolve(clipboard),
    });

    await pasteTextSelection(host);

    expect(state.text).toBe('a'.repeat(PASTE_CHARACTER_LIMIT - 1));
    expect(engineOps(state.ops)).toEqual([`insert(0,${'a'.repeat(PASTE_CHARACTER_LIMIT - 1)})`]);
    expect(state.pasteLimits).toEqual([6]);
  });

  it('replaces the selection before capping, so the cap cannot resurrect old text', async () => {
    const clipboard = 'b'.repeat(PASTE_CHARACTER_LIMIT + 50);
    const { host, state } = createHost({
      text: 'Hello world',
      selection: range(6, 11),
      read: () => Promise.resolve(clipboard),
    });

    await pasteTextSelection(host);

    expect(state.text).toBe(`Hello ${'b'.repeat(PASTE_CHARACTER_LIMIT)}`);
    expect(state.pasteLimits).toEqual([50]);
  });

  it('measures the cut against the original text', () => {
    expect(limitPasteText('short')).toEqual({ text: 'short', droppedCharacters: 0 });
    expect(limitPasteText('')).toEqual({ text: '', droppedCharacters: 0 });
    expect(limitPasteText('a\r\nb')).toEqual({ text: 'a\r\nb', droppedCharacters: 0 });

    const over = 'x'.repeat(PASTE_CHARACTER_LIMIT + 7);
    expect(limitPasteText(over)).toEqual({
      text: 'x'.repeat(PASTE_CHARACTER_LIMIT),
      droppedCharacters: 7,
    });
  });
});

describe('pptx clipboard queue', () => {
  it('runs queued pastes one at a time, each from the caret the last left', async () => {
    const first = deferred<string>();
    const second = deferred<string>();
    const reads = [first.promise, second.promise];
    const { host, state } = createHost({
      text: 'ab',
      selection: range(1, 1),
      read: () => reads.shift() ?? Promise.resolve(''),
    });
    const enqueue = createClipboardQueue();

    const pasteA = enqueue(() => pasteTextSelection(host));
    const pasteB = enqueue(() => pasteTextSelection(host));
    await Promise.resolve();

    expect(state.ops).toEqual(['read']);

    first.resolve('X');
    await pasteA;
    second.resolve('Y');
    await pasteB;

    expect(state.text).toBe('aXYb');
    expect(state.ops).toEqual(['read', 'insert(1,X)', 'read', 'insert(2,Y)']);
    expect(state.commits).toEqual([range(2, 2), range(3, 3)]);
  });

  it('serializes two oversized pastes, each capped on its own', async () => {
    const first = deferred<string>();
    const second = deferred<string>();
    const reads = [first.promise, second.promise];
    const { host, state } = createHost({
      text: '',
      selection: range(0, 0),
      read: () => reads.shift() ?? Promise.resolve(''),
    });
    const enqueue = createClipboardQueue();

    const pasteA = enqueue(() => pasteTextSelection(host));
    const pasteB = enqueue(() => pasteTextSelection(host));
    await Promise.resolve();

    expect(state.ops).toEqual(['read']);

    first.resolve('a'.repeat(PASTE_CHARACTER_LIMIT + 3));
    await pasteA;
    second.resolve('b'.repeat(PASTE_CHARACTER_LIMIT + 4));
    await pasteB;

    expect(state.text).toBe(
      `${'a'.repeat(PASTE_CHARACTER_LIMIT)}${'b'.repeat(PASTE_CHARACTER_LIMIT)}`
    );
    expect(state.pasteLimits).toEqual([3, 4]);
    expect(state.commits).toEqual([
      range(PASTE_CHARACTER_LIMIT, PASTE_CHARACTER_LIMIT),
      range(PASTE_CHARACTER_LIMIT * 2, PASTE_CHARACTER_LIMIT * 2),
    ]);
  });

  it('keeps running queued tasks after one rejects', async () => {
    const enqueue = createClipboardQueue();
    const seen: string[] = [];

    const failing = enqueue(async () => {
      seen.push('first');
      throw new Error('clipboard unavailable');
    });

    expect(
      await failing.then(
        () => 'resolved',
        () => 'rejected'
      )
    ).toBe('rejected');
    await enqueue(async () => {
      seen.push('second');
    });

    expect(seen).toEqual(['first', 'second']);
  });
});
