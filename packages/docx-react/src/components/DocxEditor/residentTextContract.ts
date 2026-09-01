/**
 * Which inserted text the resident engine path may carry.
 *
 * This used to be `/^[\x20-\x7e]+$/`, which sent every non-Latin script down
 * the compatibility relayout — Vietnamese, CJK and anything else typed through
 * an IME paid a full document pass per commit while ASCII did not. The engine
 * never had that rule: `apply_input` rejects only empty text and paragraph
 * breaks, and its fast path is byte-identical to a full pass on Vietnamese,
 * CJK, astral characters and combining marks alike.
 *
 * What the resident path genuinely cannot carry is a control character. A
 * paragraph break is a structural operation rather than text, and a tab is its
 * own kind of run, so both keep taking the compatibility path.
 */
const CONTROL_OR_LINE_SEPARATOR = /[\p{Cc}\p{Zl}\p{Zp}]/u;

export function isResidentInputText(text: string): boolean {
  return text.length > 0 && !CONTROL_OR_LINE_SEPARATOR.test(text);
}

/**
 * Whether a run of text may be COALESCED with later input into one operation.
 *
 * Deliberately narrower than [`isResidentInputText`]. Batching widens the blast
 * radius of anything that drops a pending run: one lost operation loses every
 * character accumulated into it, not one. `insertText` has such a path — it
 * returns without inserting when no selection or position map resolves — so
 * until that path cannot lose text, only the input that was already batched
 * before is batched.
 *
 * Text outside this set still takes the resident fast path; it is applied one
 * operation at a time instead of coalesced.
 */
const BATCHABLE = /^[\x20-\x7e]+$/u;

export function isBatchableInputText(text: string): boolean {
  return BATCHABLE.test(text);
}
