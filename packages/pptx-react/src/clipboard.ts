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
  /**
   * Tells the user a paste was cut short, given how many characters were left
   * out. Called only when {@link limitPasteText} actually drops something, and
   * always before the caret commits, so the banner is up by the time the
   * shortened text appears.
   */
  reportPasteLimit: (droppedCharacters: number) => void;
}

export type ClipboardQueue = (task: () => Promise<void>) => Promise<void>;

/**
 * Ceiling on the characters one paste may insert.
 *
 * A paste costs one engine op per line plus one per break between them, and
 * every op is a synchronous wasm round-trip whose cost grows with the story it
 * is inserting into — so an unbounded clipboard freezes the renderer for as
 * long as it takes, with no frame in between to paint or cancel. The engine
 * exposes no way to batch those ops into one undo step either, so an unbounded
 * paste is also an unbounded number of undo presses to take back.
 *
 * Both caps are far past any real slide: a text-heavy slide runs a few hundred
 * characters over a dozen or so lines, so this bounds the damage without
 * standing in the way of anything a deck would plausibly hold.
 */
export const PASTE_CHARACTER_LIMIT = 10_000;

/**
 * Ceiling on the paragraphs one paste may create. Needed alongside the
 * character cap because ops scale with lines, not length: 10,000 newlines are
 * only 10,000 characters but would still be 20,000 ops.
 */
export const PASTE_LINE_LIMIT = 500;

const LINE_BREAK = /\r\n|\r|\n/;

export interface LimitedPasteText {
  text: string;
  droppedCharacters: number;
}

/**
 * Cuts clipboard text down to what one paste is allowed to insert, reporting
 * how much was left out so the caller can say so.
 *
 * The cut is an index into the original string rather than a re-join of parsed
 * lines, so text under the caps comes through byte for byte — including which
 * flavour of line ending it arrived with. A cut that lands between a `\r` and
 * its `\n` takes the stranded `\r` with it, which would otherwise read as one
 * more empty paragraph than the text actually had.
 */
export function limitPasteText(clipboard: string): LimitedPasteText {
  let cut = Math.min(clipboard.length, PASTE_CHARACTER_LIMIT, lineLimitIndex(clipboard));
  if (cut >= clipboard.length) return { text: clipboard, droppedCharacters: 0 };
  if (clipboard[cut - 1] === '\r') cut -= 1;
  return { text: clipboard.slice(0, cut), droppedCharacters: clipboard.length - cut };
}

/** The index at which keeping more text would exceed {@link PASTE_LINE_LIMIT}. */
function lineLimitIndex(clipboard: string): number {
  const breaks = new RegExp(LINE_BREAK.source, 'g');
  let lines = 1;
  for (let match = breaks.exec(clipboard); match; match = breaks.exec(clipboard)) {
    lines += 1;
    if (lines > PASTE_LINE_LIMIT) return match.index;
  }
  return clipboard.length;
}

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
 * line and paragraph break, up to {@link limitPasteText}'s caps.
 *
 * The insert loop stays one synchronous burst on purpose. Yielding between
 * chunks to let the renderer paint would hand the story to whatever else runs
 * in that gap — typing, a remote edit, another queued paste — and every index
 * after `range.start` is computed against the story as it was, so the rest of
 * the paste would land in the wrong place. Bounding the work is what keeps
 * that burst short; the queue only stops two pastes from interleaving.
 */
export async function pasteTextSelection(host: ClipboardHost): Promise<void> {
  const target = host.selection();
  const range = textRangeOf(target);
  if (!target || !range) return;
  let raw: string;
  try {
    raw = await host.readClipboard();
  } catch {
    return;
  }
  if (!raw) return;
  if (!sameTextSelection(target, host.selection())) return;
  const { text: clipboard, droppedCharacters } = limitPasteText(raw);
  if (droppedCharacters > 0) host.reportPasteLimit(droppedCharacters);
  try {
    if (range.start !== range.end) {
      host.deleteText(target.storyId, range.start, range.end);
    }
    let index = range.start;
    for (const [lineIndex, line] of clipboard.split(LINE_BREAK).entries()) {
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
