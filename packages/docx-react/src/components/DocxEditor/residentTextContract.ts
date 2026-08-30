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
