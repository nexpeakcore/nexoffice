/**
 * Mounts the accessibility mirror (core `buildMirrorPage`) 1:1 under one
 * canvas page: same origin, same page-local pixel space, so
 * `getBoundingClientRect` on mirror nodes returns the rects the canvas
 * painted. The mirror is invisible (opacity 0) and inert to the pointer
 * (pointer-events none) but deliberately NOT aria-hidden — it is the
 * accessible content of the canvas. Rebuilt whenever the page's display list
 * changes — the same trigger that re-rasters the canvas.
 *
 * Windowed alongside the canvas bitmap, but never emptied: a large document's
 * full mirror is hundreds of thousands of DOM nodes, so an off-window page
 * falls back to `buildMirrorPageOutline` — paragraph text in reading order,
 * no geometry — and is promoted back to the positioned mirror on re-entry.
 * Off-window pages therefore stay in the accessibility tree (a screen reader
 * still reaches the whole document); only geometry consumers require the
 * windowed full mirror, and those already read live canvas rects.
 *
 * Focus never lands here: the hidden input remains the editing surface.
 */

import { memo, useEffect, useRef } from 'react';
import {
  buildMirrorPage,
  buildMirrorPageOutline,
  type DisplayPage,
} from '@betteroffice/docx/layout/render';
import { useTranslation } from '../../i18n';

/**
 * Trailing delay before the mirror DOM is rebuilt after a page update. During
 * a typing burst a per-keystroke rebuild churns thousands of text nodes and
 * turns every later getBoundingClientRect into a full forced reflow; screen
 * readers are indifferent to a fraction-of-a-second refresh.
 */
const MIRROR_REBUILD_DELAY_MS = 200;

function CanvasPageMirrorComponent({
  page,
  zoom = 1,
  live = true,
}: {
  page: DisplayPage;
  zoom?: number;
  /**
   * Whether the page is inside the page window. An off-window page keeps a
   * text-only outline mirror instead of the positioned one.
   */
  live?: boolean;
}) {
  const hostRef = useRef<HTMLDivElement>(null);
  const { t } = useTranslation();
  const builtOnceRef = useRef(false);

  useEffect(() => {
    const host = hostRef.current;
    if (!host) return;
    const labels = {
      page: t('a11y.pageLabel', { number: page.pageIndex + 1 }),
      header: t('a11y.headerLabel'),
      footer: t('a11y.footerLabel'),
    };
    if (!live) {
      builtOnceRef.current = false;
      host.replaceChildren(buildMirrorPageOutline(page, { labels }));
      return;
    }
    const build = (): void => {
      const mirror = buildMirrorPage(page, { labels });
      // Keep the previous mirror connected until this replacement is ready.
      // Clearing in effect cleanup creates a detached-DOM window on every page
      // update; unmounting already removes the host and its complete subtree.
      host.replaceChildren(mirror);
    };
    if (!builtOnceRef.current) {
      builtOnceRef.current = true;
      build();
      return;
    }
    const timer = setTimeout(build, MIRROR_REBUILD_DELAY_MS);
    return () => clearTimeout(timer);
  }, [page, t, live]);

  return (
    <div
      ref={hostRef}
      className="canvas-page-mirror"
      // The mirror content is built in page-local px; when the canvas is
      // enlarged for zoom (CSS size = page * zoom), CSS-scale the mirror by the
      // same factor from its top-left origin so its nodes' `getBoundingClientRect`
      // still lands on the painted glyphs. At zoom = 1 this is an identity scale.
      style={{
        position: 'absolute',
        left: 0,
        top: 0,
        pointerEvents: 'none',
        transform: zoom !== 1 ? `scale(${zoom})` : undefined,
        transformOrigin: '0 0',
      }}
    />
  );
}

/**
 * A page the reader is not looking at renders identically on every keystroke.
 * With a page per canvas, an unmemoed mirror re-renders the whole document —
 * measured at 338 of these per render of the page list, and the page list
 * renders more than once a key.
 */
export const CanvasPageMirror = memo(CanvasPageMirrorComponent);
