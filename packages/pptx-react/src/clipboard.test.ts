import { describe, expect, it } from 'bun:test';
import {
  copyTextSelection,
  createClipboardQueue,
  cutTextSelection,
  deleteTextSelection,
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
  };
  return { host, state };
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
