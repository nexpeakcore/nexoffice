/**
 * An IME commit has nowhere to live between the moment the composition box
 * clears and the moment the frame carrying the text is presented — a round trip
 * through the engine. Left alone the word blinks out and comes back.
 *
 * This holds the committed text on screen across that gap, and releases it the
 * instant the frame that contains it lands, so the reader never sees the hole.
 */

export interface CommittedComposition {
  text: string;
  /** Where the composition box was, so the preview does not appear to move. */
  left: number;
  top: number;
  height: number;
}

/** Longest a commit stays on screen when its frame never arrives. */
export const COMMITTED_COMPOSITION_TIMEOUT_MS = 2000;

export class CommittedCompositionHold {
  private held: CommittedComposition | null = null;
  private awaitedFrameEpoch: number | null = null;
  private heldSince = Number.NEGATIVE_INFINITY;

  /**
   * Hold `composition` until the frame at `frameEpoch` is presented. A null
   * epoch means the edit produced no frame to wait for, so only the timeout
   * releases it.
   */
  hold(composition: CommittedComposition, frameEpoch: number | null, now: number): void {
    if (!composition.text) {
      this.clear();
      return;
    }
    this.held = composition;
    this.awaitedFrameEpoch = frameEpoch;
    this.heldSince = now;
  }

  /** What to paint over the page, or null when the page is showing it itself. */
  visible(): CommittedComposition | null {
    return this.held;
  }

  /**
   * A presented frame. Returns whether this released the hold — the page is now
   * drawing the text, so the preview must go or it would be drawn twice.
   */
  onFramePresented(frameEpoch: number | null): boolean {
    if (!this.held) return false;
    if (this.awaitedFrameEpoch === null) return false;
    if (frameEpoch === null || frameEpoch < this.awaitedFrameEpoch) return false;
    this.clear();
    return true;
  }

  /**
   * Releases a commit whose frame never came — a lost worker, a failed layout.
   * Without it the text would stay painted over a page that does not have it.
   */
  tick(now: number): boolean {
    if (!this.held) return false;
    if (now - this.heldSince < COMMITTED_COMPOSITION_TIMEOUT_MS) return false;
    this.clear();
    return true;
  }

  clear(): void {
    this.held = null;
    this.awaitedFrameEpoch = null;
    this.heldSince = Number.NEGATIVE_INFINITY;
  }
}
