/**
 * Text clipboard actions, lifted out of the editor component so the awaits
 * around the system clipboard are testable without a real one.
 *
 * The system clipboard is slow and fallible: a permission prompt, a Wayland
 * bridge or a busy compositor can hold `read`/`write` for hundreds of
 * milliseconds, and the caret is free to move meanwhile. So every action
 * re-reads the live selection through {@link ClipboardHost.selection} after
 * each await instead of trusting the one it started with — a cut whose copy
 * failed keeps its text, and a paste whose selection moved is dropped rather
 * than applied to whatever now occupies that range.
 */

export interface ClipboardTextSelection {
  shapeId: string;
  storyId: string;
  anchor: number;
  focus: number;
}

export interface TextRange {
  start: number;
  end: number;
}

/**
 * The editor seam these actions drive. Every member reads or writes live state
 * at call time: `selection` must return the caret as it is now, never a value
 * captured when the action began.
 */
export interface ClipboardHost {
  selection: () => ClipboardTextSelection | null;
  storyText: (storyId: string) => string;
  deleteText: (storyId: string, start: number, end: number) => void;
  insertText: (storyId: string, index: number, text: string) => void;
  insertParagraphBreak: (storyId: string, index: number) => void;
  commit: (selection: ClipboardTextSelection) => void;
  readClipboard: () => Promise<string>;
  writeClipboard: (text: string) => Promise<void>;
  reportError: (value: unknown) => void;
}

export type ClipboardQueue = (task: () => Promise<void>) => Promise<void>;

export function textRangeOf(
  selection: ClipboardTextSelection | null
): TextRange | null {
  if (!selection) return null;
  return {
    start: Math.min(selection.anchor, selection.focus),
    end: Math.max(selection.anchor, selection.focus),
  };
}

export function sameTextSelection(
  a: ClipboardTextSelection | null,
  b: ClipboardTextSelection | null
): boolean {
  if (!a || !b) return a === b;
  return (
    a.shapeId === b.shapeId &&
    a.storyId === b.storyId &&
    a.anchor === b.anchor &&
    a.focus === b.focus
  );
}

/**
 * Runs clipboard actions one at a time. Two Cmd+V presses a frame apart would
 * otherwise both resolve their `readText` against the same pre-paste caret and
 * write over each other; queued, the second starts from the caret the first
 * left behind.
 */
export function createClipboardQueue(): ClipboardQueue {
  let tail = Promise.resolve();
  return (task) => {
    const run = tail.then(task);
    tail = run.then(
      () => undefined,
      () => undefined
    );
    return run;
  };
}

/** Copies the selected text, reporting whether the clipboard actually took it. */
export async function copyTextSelection(host: ClipboardHost): Promise<boolean> {
  return writeSelectionToClipboard(host, host.selection());
}

/**
 * Copies the selection and deletes it only once the clipboard has the text and
 * the caret has not moved — a cut that deleted after a failed or superseded
 * copy would destroy text no one could paste back.
 */
export async function cutTextSelection(host: ClipboardHost): Promise<void> {
  const target = host.selection();
  const copied = await writeSelectionToClipboard(host, target);
  if (!copied) return;
  if (!sameTextSelection(target, host.selection())) return;
  deleteTextSelection(host);
}

export function deleteTextSelection(host: ClipboardHost): void {
  const target = host.selection();
  const range = textRangeOf(target);
  if (!target || !range || range.start === range.end) return;
  try {
    host.deleteText(target.storyId, range.start, range.end);
    host.commit({ ...target, anchor: range.start, focus: range.start });
  } catch (value) {
    host.reportError(value);
  }
}

/**
 * Replaces the selection with the clipboard's plain text, one engine op per
 * line and paragraph break. The wasm engine exposes no way to batch ops into
 * one undo step, so each of those is its own step today; the loop at least
 * stays a single synchronous burst, with no await between the delete and the
 * inserts for another action to slip into.
 */
export async function pasteTextSelection(host: ClipboardHost): Promise<void> {
  const target = host.selection();
  const range = textRangeOf(target);
  if (!target || !range) return;
  let clipboard: string;
  try {
    clipboard = await host.readClipboard();
  } catch {
    return;
  }
  if (!clipboard) return;
  if (!sameTextSelection(target, host.selection())) return;
  try {
    if (range.start !== range.end) {
      host.deleteText(target.storyId, range.start, range.end);
    }
    let index = range.start;
    for (const [lineIndex, line] of clipboard.split(/\r\n|\r|\n/).entries()) {
      if (lineIndex > 0) {
        host.insertParagraphBreak(target.storyId, index);
        index += 1;
      }
      if (line) {
        host.insertText(target.storyId, index, line);
        index += line.length;
      }
    }
    host.commit({ ...target, anchor: index, focus: index });
  } catch (value) {
    host.reportError(value);
  }
}

async function writeSelectionToClipboard(
  host: ClipboardHost,
  target: ClipboardTextSelection | null
): Promise<boolean> {
  const range = textRangeOf(target);
  if (!target || !range || range.start === range.end) return false;
  let text: string;
  try {
    text = host.storyText(target.storyId).slice(range.start, range.end);
  } catch (value) {
    host.reportError(value);
    return false;
  }
  if (!text) return false;
  try {
    await host.writeClipboard(text);
  } catch {
    return false;
  }
  return true;
}
