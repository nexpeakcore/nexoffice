import {
  nextPageWindow,
  PAGE_WINDOW_MIN_PAGES,
  type PageWindowRange,
} from './pageWindow';
import {
  useCallback,
  useEffect,
  useLayoutEffect,
  useMemo,
  useRef,
  useState,
  type ReactNode,
  type Ref,
} from 'react';
import { findVerticalScrollParentOrRoot } from '@betteroffice/docx/utils/findVerticalScrollParent';
import {
  presentDisplayPageBackBuffer,
  rasterizeDisplayPageToBackBuffer,
  GlyphCache,
  loadGlyphOutlineProvider,
  type DisplayList,
  type GlyphOutlineProvider,
  type ImageResolver,
  type RetainedFrame,
} from '@betteroffice/docx/layout/render';
import type { UseCanvasRendererResult } from './hooks/useDisplayList';
import { CanvasPageMirror } from './CanvasPageMirror';
import { CanvasInteractiveOverlay } from './CanvasInteractiveOverlay';
import { CanvasA11yLiveRegion, type CanvasA11yLiveRegionProps } from './CanvasA11yLiveRegion';
import { CANVAS_PAGE_GAP_PX, CANVAS_PAGES_PADDING_PX } from '@betteroffice/docx/layout/render';
import { SIDEBAR_DOCUMENT_SHIFT } from '../sidebar/constants';
import { DefaultLoadingIndicator, ParseError } from '../DocxEditorHelpers';
import { displayListNeedsHostImages } from './canvasPresentation';
import { resolveCaretPaintColor } from './paintedCaret';
import { DEFAULT_CARET_WIDTH } from './overlays/SelectionOverlay';

// Canvas is the sole visible renderer. The editing/input subtree stays mounted
// independently so hidden input focus and IME state survive initial
// loading and renderer errors.
export function CanvasPagedArea({
  renderer,
  a11y,
  sidebarOpen = false,
  zoom = 1,
  interactive = false,
  children,
}: {
  renderer: UseCanvasRendererResult;
  /** live-region wiring (host notify ref + Yrs session getter) — see CanvasA11yLiveRegion */
  a11y?: Omit<CanvasA11yLiveRegionProps, 'active'>;
  /** shifts the canvas pages left to make room for the comments sidebar, mirroring the DOM painter's viewport transform */
  sidebarOpen?: boolean;
  /** zoom level (1 = 100%); the canvas re-rasters at `zoom * DPR` so text stays crisp */
  zoom?: number;
  /** mounts the focusable content-control (SDT) overlay above each page; off in read-only mode */
  interactive?: boolean;
  children: ReactNode;
}) {
  return (
    <>
      {renderer.status === 'ready' && renderer.displayList ? (
        <CanvasPagesView
          displayList={renderer.displayList}
          frame={renderer.frame}
          resolveImage={renderer.resolveImage}
          hostRef={renderer.canvasHostRef}
          sidebarOpen={sidebarOpen}
          zoom={zoom}
          interactive={interactive}
          glyphOutlineProvider={renderer.glyphOutlineProvider}
          offscreenReplay={renderer.offscreenReplay}
          onWorkerPresentationChange={renderer.setWorkerPresentationActive}
          onPageWindowChange={renderer.setPageWindow}
        />
      ) : renderer.status === 'error' ? (
        <div data-testid="canvas-renderer-error" role="alert" style={{ minHeight: 240 }}>
          <ParseError message={renderer.error?.message ?? 'Canvas renderer failed.'} />
        </div>
      ) : (
        <div data-testid="canvas-renderer-loading" role="status" style={{ minHeight: 240 }}>
          <DefaultLoadingIndicator />
        </div>
      )}
      {children}
      {a11y ? <CanvasA11yLiveRegion active={renderer.status === 'ready'} {...a11y} /> : null}
    </>
  );
}



/** Replays display-list pages to canvas with accessibility mirrors. */
export function CanvasPagesView({
  displayList,
  frame,
  resolveImage,
  hostRef,
  sidebarOpen = false,
  zoom = 1,
  interactive = false,
  glyphOutlineProvider,
  offscreenReplay,
  onWorkerPresentationChange,
  onPageWindowChange,
}: {
  displayList: DisplayList;
  /** Binary retained-frame metadata used to scope page replay. */
  frame?: RetainedFrame | null;
  resolveImage?: ImageResolver;
  /** pointer routing maps client coords → page-local through this host element */
  hostRef?: Ref<HTMLDivElement>;
  /** shift the page column left to reserve room for the comments sidebar */
  sidebarOpen?: boolean;
  /**
   * Zoom level (1 = 100%). Instead of CSS-scaling the page column (which would
   * blur the rastered text), each page canvas is re-sized and re-drawn at
   * `zoom * devicePixelRatio` so glyph outlines stay crisp — see
   * `sizeCanvasForPage`. The a11y mirror is CSS-scaled to keep its nodes 1:1
   * over the enlarged canvas.
   */
  zoom?: number;
  /**
   * Mounts the interactive content-control overlay (focusable SDT widgets)
   * above each page. The a11y mirror stays pointer-inert; this separate layer
   * owns the only clickable/focusable SDT controls on the canvas path.
   */
  interactive?: boolean;
  /** Outline source sharing the display engine's resident font store. */
  glyphOutlineProvider?: GlyphOutlineProvider | null;
  /** Dedicated worker replay surface; unsupported/media-heavy pages use DOM canvas. */
  offscreenReplay?: UseCanvasRendererResult['offscreenReplay'];
  onWorkerPresentationChange?: (active: boolean) => void;
  /** Reports the pages near the viewport, so the rest need not be carried. */
  onPageWindowChange?: (window: PageWindowRange | null) => void;
}) {
  const canvasesRef = useRef(new Map<string, HTMLCanvasElement>());
  const transferredCanvasesRef = useRef(new WeakSet<HTMLCanvasElement>());
  const presentedCanvasesRef = useRef(new WeakSet<HTMLCanvasElement>());
  const offscreenSignatureRef = useRef('');
  const replayGenerationRef = useRef(0);
  const [offscreenFailed, setOffscreenFailed] = useState(false);
  const offscreenFailedRef = useRef(false);
  const offscreenAttachedRef = useRef(false);
  const workerPresentationRef = useRef(false);
  const publishWorkerPresentation = useCallback(
    (active: boolean) => {
      if (workerPresentationRef.current === active) return;
      workerPresentationRef.current = active;
      onWorkerPresentationChange?.(active);
    },
    [onWorkerPresentationChange]
  );
  const offscreenEligible = useMemo(
    () => Boolean(offscreenReplay && frame && !displayListNeedsHostImages(displayList)),
    [displayList, frame, offscreenReplay]
  );
  useEffect(() => {
    if (!offscreenEligible || offscreenFailed) publishWorkerPresentation(false);
  }, [offscreenEligible, offscreenFailed, publishWorkerPresentation]);
  useEffect(
    () => () => {
      publishWorkerPresentation(false);
    },
    [publishWorkerPresentation]
  );
  const rasterEnvironmentRef = useRef<{
    dpr: number;
    zoom: number;
    glyphCacheReady: boolean;
    resolveImage?: ImageResolver;
  } | null>(null);

  // ===========================================================================
  // Page windowing: only pages near the viewport hold rastered bitmaps and the
  // positioned a11y mirror. Every page keeps its canvas element (stable keys
  // and CSS-sized boxes, so scroll geometry, pointer routing, and canvas-rect
  // overlays are untouched); an off-window page's backing store is released
  // (attributes zeroed on the DOM path, offscreen buffer zeroed by the worker)
  // and its mirror drops to a text-only outline, both restored on re-entry.
  // The window moves only with scrolling/resize/zoom, never with document
  // invalidation.
  // ===========================================================================
  const innerHostRef = useRef<HTMLDivElement | null>(null);
  const setHostRef = useMemo(
    () =>
      (element: HTMLDivElement | null): void => {
        innerHostRef.current = element;
        if (typeof hostRef === 'function') hostRef(element);
        else if (hostRef) (hostRef as { current: HTMLDivElement | null }).current = element;
      },
    [hostRef]
  );
  const pageWindowAllowed = useMemo(() => {
    if (typeof window === 'undefined') return false;
    // diagnostic escape hatch, mirroring `offscreenReplay=0`
    return new URLSearchParams(window.location.search).get('pageWindow') !== '0';
  }, []);
  const windowingEnabled = pageWindowAllowed && displayList.pages.length > PAGE_WINDOW_MIN_PAGES;
  const [pageWindow, setPageWindow] = useState<PageWindowRange | null>(null);
  // Column-space page tops/bottoms from display-list geometry alone (no DOM
  // reads): padding, then each page height at the current zoom plus the gap.
  const pageOffsets = useMemo(() => {
    const tops = new Array<number>(displayList.pages.length);
    const bottoms = new Array<number>(displayList.pages.length);
    let y = CANVAS_PAGES_PADDING_PX;
    displayList.pages.forEach((page, index) => {
      tops[index] = y;
      bottoms[index] = y + page.height * zoom;
      y = bottoms[index] + CANVAS_PAGE_GAP_PX;
    });
    return { tops, bottoms };
  }, [displayList, zoom]);
  useLayoutEffect(() => {
    if (!windowingEnabled) {
      setPageWindow(null);
      return;
    }
    const host = innerHostRef.current;
    if (!host) return;
    let rafId: number | null = null;
    const recompute = (): void => {
      rafId = null;
      const column = host.firstElementChild as HTMLElement | null;
      if (!column || !host.isConnected) {
        // unmeasurable — fail open (all pages live) so replay is never
        // deferred forever
        setPageWindow(
          (previous) => previous ?? { start: 0, end: displayList.pages.length - 1 }
        );
        return;
      }
      // Resolved per-recompute: at first mount the editor's scroll container
      // is not yet scrollable (the walk requires overflow height), so a
      // one-time resolution here permanently falls back to the root and the
      // window never follows the real scroller.
      const scrollParent = findVerticalScrollParentOrRoot(host);
      const isRoot =
        scrollParent === document.scrollingElement || scrollParent === document.documentElement;
      // client rects are viewport-relative: the visible band starts at the
      // scroller's client top for an element scroller, at 0 for the root
      const viewportTop = isRoot ? 0 : scrollParent.getBoundingClientRect().top;
      const viewportHeight = isRoot ? window.innerHeight : scrollParent.clientHeight;
      const columnRect = column.getBoundingClientRect();
      const viewTop = viewportTop - columnRect.top;
      const viewBottom = viewTop + viewportHeight;
      const { tops, bottoms } = pageOffsets;
      let first = tops.length - 1;
      for (let index = 0; index < tops.length; index += 1) {
        if (bottoms[index] >= viewTop) {
          first = index;
          break;
        }
      }
      let last = first;
      for (let index = tops.length - 1; index >= first; index -= 1) {
        if (tops[index] <= viewBottom) {
          last = index;
          break;
        }
      }
      setPageWindow((previous) => nextPageWindow(previous, first, last, tops.length));
    };
    const schedule = (): void => {
      if (rafId === null) rafId = requestAnimationFrame(recompute);
    };
    recompute();

    // Capture-phase on window: scroll events do not bubble off an element
    // scroller, and which element scrolls the editor can change after mount.
    window.addEventListener('scroll', schedule, { capture: true, passive: true });
    window.addEventListener('resize', schedule);
    return () => {
      if (rafId !== null) cancelAnimationFrame(rafId);
      window.removeEventListener('scroll', schedule, { capture: true });
      window.removeEventListener('resize', schedule);
    };
  }, [windowingEnabled, pageOffsets]);

  // The raster window is also the data window: a page whose bitmap is released
  // has no reader for its primitives either.
  const onPageWindowChangeRef = useRef(onPageWindowChange);
  onPageWindowChangeRef.current = onPageWindowChange;
  useEffect(() => {
    onPageWindowChangeRef.current?.(windowingEnabled ? pageWindow : null);
  }, [pageWindow, windowingEnabled]);
  // Until the first measurement lands (set pre-paint by the layout effect
  // above), the replay effect is deferred entirely — never guess a window
  // that could blank a visible page, and never raster every page of a large
  // document just because the window is not measured yet.
  const windowPending = windowingEnabled && pageWindow === null;
  const effectiveWindow: PageWindowRange | null = windowingEnabled ? pageWindow : null;
  const pageInWindow = (index: number): boolean =>
    effectiveWindow === null || (index >= effectiveWindow.start && index <= effectiveWindow.end);

  // One glyph-outline cache for the canvas lifetime (task contract: not
  // per-render). The wasm-backed outline provider loads lazily through the
  // SAME module the display-list builder already resolved — no extra fetch.
  // `glyphCacheReady` re-runs the draw effect once the provider lands so the
  // first shaped frame repaints as real glyph outlines (until then a glyphRun
  // falls back to fillText inside the backend, so text is never blank).
  const glyphCacheRef = useRef<GlyphCache | null>(null);
  const [glyphCacheReady, setGlyphCacheReady] = useState(false);
  useEffect(() => {
    setOffscreenFailed(false);
    offscreenFailedRef.current = false;
    offscreenAttachedRef.current = false;
    offscreenSignatureRef.current = '';
  }, [offscreenReplay]);
  useEffect(() => {
    let cancelled = false;
    glyphCacheRef.current = null;
    setGlyphCacheReady(false);
    const provider = glyphOutlineProvider
      ? Promise.resolve(glyphOutlineProvider)
      : loadGlyphOutlineProvider();
    void provider
      .then((provider) => {
        if (cancelled) return;
        glyphCacheRef.current = new GlyphCache({ provider });
        setGlyphCacheReady(true);
      })
      .catch(() => {
        // outline export absent → the backend keeps painting glyph runs with
        // fillText; nothing to do here.
      });
    return () => {
      cancelled = true;
    };
  }, [glyphOutlineProvider]);

  const windowStart = effectiveWindow?.start ?? -1;
  const windowEnd = effectiveWindow?.end ?? -1;
  useEffect(() => {
    // The window measurement lands pre-paint (layout effect) and re-runs this
    // effect; rastering before it exists would process every page.
    if (windowPending) return;
    const replayGeneration = ++replayGenerationRef.current;
    const dpr = window.devicePixelRatio || 1;
    if (offscreenEligible && !offscreenFailed && frame && offscreenReplay) {
      const pages: Array<{ pageId: string; canvas: OffscreenCanvas }> = [];
      const activePageIds: string[] = [];
      for (let index = 0; index < frame.pages.length; index += 1) {
        if (!pageInWindow(index)) continue;
        const retainedPage = frame.pages[index];
        const page = displayList.pages[index];
        const pageId = retainedPage.pageId.toString();
        activePageIds.push(pageId);
        const canvas = canvasesRef.current.get(pageId);
        if (!canvas || !page) continue;
        canvas.style.width = `${page.width * zoom}px`;
        canvas.style.height = `${page.height * zoom}px`;
        if (transferredCanvasesRef.current.has(canvas)) continue;
        try {
          pages.push({ pageId, canvas: canvas.transferControlToOffscreen() });
          transferredCanvasesRef.current.add(canvas);
        } catch {
          offscreenFailedRef.current = true;
          publishWorkerPresentation(false);
          setOffscreenFailed(true);
          return;
        }
      }
      const caretColor = resolveCaretPaintColor(innerHostRef.current);
      const caretStyle = { color: caretColor, width: DEFAULT_CARET_WIDTH };
      const signature = `${activePageIds.join(',')}|${dpr}|${zoom}|${caretColor}`;
      if (pages.length > 0 || signature !== offscreenSignatureRef.current) {
        offscreenSignatureRef.current = signature;
        void offscreenReplay.attach(pages, activePageIds, dpr, zoom, caretStyle).then((attached) => {
          // Publish on resolution regardless of replay generation: attach
          // resolutions are FIFO, so the last one reflects the worker's real
          // attachment state. Gating on the generation dropped the publish
          // whenever a frame landed while the first attach was still
          // rastering — i.e. on every document load — leaving presentation
          // permanently unpublished while the worker was in fact presenting.
          offscreenAttachedRef.current = attached;
          if (!offscreenFailedRef.current) publishWorkerPresentation(attached);
          if (!attached) {
            // transient (no worker client yet) — clear the signature so the
            // next pass retries instead of permanently flipping surfaces
            offscreenSignatureRef.current = '';
          }
        }, () => {
          offscreenFailedRef.current = true;
          publishWorkerPresentation(false);
          setOffscreenFailed(true);
        });
      } else if (offscreenAttachedRef.current) {
        // Heal any publish lost to ordering (StrictMode remount, late
        // resolution): the worker is attached and this pass kept it active.
        publishWorkerPresentation(true);
      }
      return;
    }
    const glyphCache = glyphCacheRef.current ?? undefined;
    const previousEnvironment = rasterEnvironmentRef.current;
    const rasterEnvironmentChanged =
      !previousEnvironment ||
      previousEnvironment.dpr !== dpr ||
      previousEnvironment.zoom !== zoom ||
      previousEnvironment.glyphCacheReady !== glyphCacheReady ||
      previousEnvironment.resolveImage !== resolveImage;
    rasterEnvironmentRef.current = { dpr, zoom, glyphCacheReady, resolveImage };
    const preparations: Array<
      Promise<{
        canvas: HTMLCanvasElement;
        buffer: HTMLCanvasElement;
        page: DisplayList['pages'][number];
      }>
    > = [];
    for (const [i, page] of displayList.pages.entries()) {
      const retainedPage = frame?.pages[i];
      const pageKey = retainedPage ? retainedPage.pageId.toString() : `index:${page.pageIndex}`;
      const canvas = canvasesRef.current.get(pageKey);
      const ctx = canvas?.getContext('2d');
      if (!canvas || !ctx) continue;
      if (!pageInWindow(i)) {
        // release the off-window bitmap; the CSS-sized element stays for
        // geometry consumers. Dropping the presented mark makes re-entry
        // repaint through the ordinary remount rule below.
        if (canvas.width !== 0 || canvas.height !== 0) {
          canvas.width = 0;
          canvas.height = 0;
          presentedCanvasesRef.current.delete(canvas);
        }
        continue;
      }
      // A remounted canvas (surface-mode flip) has no pixels regardless of
      // the retained frame's damage set — always paint it.
      const damaged =
        !frame ||
        !retainedPage ||
        rasterEnvironmentChanged ||
        !presentedCanvasesRef.current.has(canvas) ||
        frame.damagedPageIds.has(retainedPage.pageId);
      if (!damaged) continue;
      // Raster off-DOM first. The connected canvas keeps its previous pixels
      // until every damaged page has finished all async image/glyph work.
      const buffer = document.createElement('canvas');
      preparations.push(
        rasterizeDisplayPageToBackBuffer(
          buffer,
          page,
          { resolveImage, glyphCache },
          dpr,
          zoom
        ).then(() => ({ canvas, buffer, page }))
      );
    }
    const present = (prepared: Awaited<(typeof preparations)[number]>[]) => {
      if (replayGeneration !== replayGenerationRef.current) return;
      for (const { canvas, buffer, page } of prepared) {
        presentDisplayPageBackBuffer(canvas, buffer, page, zoom);
        presentedCanvasesRef.current.add(canvas);
      }
    };
    void Promise.all(preparations).then(present, (error) => {
      if (replayGeneration === replayGenerationRef.current) {
        console.error('[CanvasRenderer] Atomic canvas replay failed', error);
      }
    });
    // glyphCacheReady is a redraw trigger (the cache itself is read via ref);
    // zoom re-runs the raster so the enlarged canvas paints at full resolution;
    // windowStart/windowEnd re-run it so pages entering the window paint and
    // the offscreen active set prunes pages that left it
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [
    displayList,
    frame,
    resolveImage,
    glyphCacheReady,
    offscreenEligible,
    offscreenFailed,
    offscreenReplay,
    zoom,
    windowPending,
    windowStart,
    windowEnd,
    publishWorkerPresentation,
  ]);

  // The host stays a full-width, un-transformed positioned box so the
  // interactive comment overlays (portalled in by DocxEditorPagedArea) anchor
  // their `50%`-centered X / host-relative Y to the page's un-shifted center.
  // Only the inner page column shifts left when the sidebar opens, mirroring
  // the DOM painter's viewport `translateX(-SIDEBAR_DOCUMENT_SHIFT)`. Pointer
  // routing reads each canvas's live `getBoundingClientRect`, so the transform
  // is factored out for free.
  return (
    <div ref={setHostRef} className="canvas-pages" style={{ position: 'relative' }}>
      <div
        className="canvas-pages__column"
        style={{
          display: 'flex',
          flexDirection: 'column',
          alignItems: 'center',
          gap: CANVAS_PAGE_GAP_PX,
          padding: CANVAS_PAGES_PADDING_PX,
          transform: sidebarOpen ? `translateX(-${SIDEBAR_DOCUMENT_SHIFT}px)` : undefined,
          transition: 'transform 0.2s ease',
        }}
      >
        {displayList.pages.map((page, i) => {
          const retainedPage = frame?.pages[i];
          const pageKey = retainedPage ? retainedPage.pageId.toString() : `index:${page.pageIndex}`;
          const surfaceKey = `${pageKey}:${offscreenEligible && !offscreenFailed ? 'offscreen' : 'dom'}`;
          return (
            // per-page wrapper so the mirror positions 1:1 over its canvas.
            // Every page keeps its wrapper and canvas element (stable page
            // geometry); the page window releases the bitmap backing store and
            // demotes the mirror to a text-only outline off-window.
            <div key={surfaceKey} className="canvas-page" style={{ position: 'relative' }}>
              <canvas
                ref={(el) => {
                  if (el) canvasesRef.current.set(pageKey, el);
                  else canvasesRef.current.delete(pageKey);
                }}
                data-page-index={page.pageIndex}
                style={{
                  display: 'block',
                  width: page.width * zoom,
                  height: page.height * zoom,
                  background: '#ffffff',
                  boxShadow: '0 1px 3px var(--doc-shadow)',
                }}
              />
              <CanvasPageMirror page={page} zoom={zoom} live={!windowPending && pageInWindow(i)} />
              {interactive ? <CanvasInteractiveOverlay page={page} zoom={zoom} /> : null}
            </div>
          );
        })}
      </div>
    </div>
  );
}
