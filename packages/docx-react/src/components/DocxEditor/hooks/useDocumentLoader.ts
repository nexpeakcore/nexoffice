import { useCallback, useEffect, useRef, useState } from 'react';
import type { Document } from '@betteroffice/docx/types/document';
import type { Comment } from '@betteroffice/docx/types/content';
import type { YrsDocxHost } from '@betteroffice/docx/yrs';
import {
  loadEmbeddedFonts,
  releaseBufferFontFaces,
  createBufferFontOwner,
  loadDocumentFonts,
  loadFontsWithMapping,
  getRenderableDocumentFonts,
  getEmbeddedFontFamilies,
  selectRenderableFonts,
  toArrayBuffer,
  type BufferFontOwner,
  type DocxInput,
} from '@betteroffice/docx/utils';
import type { FontOption } from '@betteroffice/docx/utils/fontOptions';
import type { UseHistoryReturn } from '../../../hooks/useHistory';
import type { PagedEditorRef } from '../PagedEditor';
import type { CommentIdAllocator } from '../commentFactories';
import { DocumentLoadGeneration } from './documentLoadGeneration';

/**
 * Document lifecycle: load buffer / pre-parsed doc, react to
 * `documentBuffer` / `document` prop changes, and extract any baked-in
 * comments from the document model on initial load.
 *
 * State reset across the editor on a fresh load is heavy (~10 distinct
 * state setters across multiple hooks), so the parent assembles a
 * single `resetForNewDocument` callback and threads it in.
 */
export function useDocumentLoader({
  documentBuffer,
  initialDocument,
  externalContent,
  history,
  pagedEditorRef,
  setLoadingState,
  setComments,
  setShowCommentsSidebar,
  onError,
  resetForNewDocument,
  commentsLoadedRef,
  commentIdAllocator,
  setDocumentFonts,
}: {
  documentBuffer: DocxInput | null | undefined;
  initialDocument: Document | null | undefined;
  externalContent: boolean | undefined;
  history: UseHistoryReturn<Document | null>;
  pagedEditorRef: React.RefObject<PagedEditorRef | null>;
  // The full EditorState shape lives in the parent; we only need to flip
  // `isLoading` and `parseError`, so the parent exposes a focused callback.
  setLoadingState: (state: { isLoading: boolean; parseError: string | null }) => void;
  setComments: React.Dispatch<React.SetStateAction<Comment[]>>;
  setShowCommentsSidebar: React.Dispatch<React.SetStateAction<boolean>>;
  onError: ((error: Error) => void) | undefined;
  resetForNewDocument: () => void;
  // `resetForNewDocument` (declared earlier in the parent) needs to clear
  // this ref on every load. Lifted out of the hook for that reason.
  commentsLoadedRef: React.RefObject<boolean>;
  // Per-editor-instance ID allocator; seeded above the loaded doc's max ID.
  commentIdAllocator: CommentIdAllocator;
  // Fonts the document references that the browser can actually render
  // (embedded or system-resolved), surfaced in the picker's "Document fonts"
  // group.
  setDocumentFonts: (fonts: FontOption[]) => void;
}) {
  // The live history document changes after every edit, but yrs must only be
  // reseeded when a new source document is loaded. Keep that load boundary
  // separate so PagedEditor can replace its session without treating normal
  // edits as fresh documents.
  const [yrsSeedDocument, setYrsSeedDocument] = useState<Document | null>(
    initialDocument ?? null
  );
  const [yrsSeedBytes, setYrsSeedBytes] = useState<Uint8Array | null>(null);
  const [yrsSeedGeneration, setYrsSeedGeneration] = useState(0);
  const [loadGeneration] = useState(() => new DocumentLoadGeneration());
  // The open document's claim on the faces it embedded. Each face holds an
  // object URL over its bytes plus an `@font-face` rule, neither of which the
  // browser reclaims on its own — so the editor hands the claim back when the
  // document is replaced or the editor goes away. A face the next document
  // also embeds survives on that document's own claim.
  const embeddedFacesRef = useRef<BufferFontOwner | null>(null);
  // Hand back whatever the ref holds. Every path that ends a document — a
  // replacement that parses, one that fails to, a pre-parsed document, the
  // editor unmounting — goes through here, so a document's faces never
  // outlive it waiting for the next successful open.
  const releaseEmbeddedFaces = useCallback(() => {
    const owner = embeddedFacesRef.current;
    embeddedFacesRef.current = null;
    if (owner) releaseBufferFontFaces(owner);
  }, []);
  useEffect(() => () => releaseEmbeddedFaces(), [releaseEmbeddedFaces]);

  const loadParsedDocument = useCallback(
    (doc: Document, seedBytes?: Uint8Array) => {
      const generation = loadGeneration.begin();
      resetForNewDocument();
      setYrsSeedDocument(doc);
      setYrsSeedBytes(seedBytes?.slice() ?? null);
      setYrsSeedGeneration(generation);
      history.reset(doc);
      releaseEmbeddedFaces();
      setLoadingState({ isLoading: false, parseError: null });
      loadDocumentFonts(doc).catch((err) => {
        console.warn('Failed to load document fonts:', err);
      });
      // Offer the document's own renderable fonts (embedded faces are loaded by
      // parseDocx; system fonts are probed) in the picker.
      setDocumentFonts(
        getRenderableDocumentFonts(doc, {
          embeddedFamilies: getEmbeddedFontFamilies(doc.package?.fontTable),
        })
      );
    },
    [
      loadGeneration,
      resetForNewDocument,
      history,
      releaseEmbeddedFaces,
      setLoadingState,
      setDocumentFonts,
    ]
  );

  const loadBuffer = useCallback(
    async (buffer: DocxInput) => {
      const generation = loadGeneration.begin();
      resetForNewDocument();
      setLoadingState({ isLoading: true, parseError: null });
      setYrsSeedDocument(null);
      setYrsSeedBytes(null);
      setYrsSeedGeneration(generation);
      try {
        const source = buffer instanceof ArrayBuffer ? buffer : await toArrayBuffer(buffer);
        if (!loadGeneration.isCurrent(generation)) return;
        setYrsSeedBytes(new Uint8Array(source));
        history.reset(null);
        await loadGeneration.waitForCompletion(generation);
      } catch (error) {
        if (!loadGeneration.complete(generation)) return;
        releaseEmbeddedFaces();
        const message = error instanceof Error ? error.message : 'Failed to parse document';
        setLoadingState({ isLoading: false, parseError: message });
        onError?.(error instanceof Error ? error : new Error(message));
      }
    },
    [loadGeneration, resetForNewDocument, history, onError, releaseEmbeddedFaces, setLoadingState]
  );

  const acceptHostDocument = useCallback(
    (host: YrsDocxHost, generation: number) => {
      if (!loadGeneration.complete(generation)) return;
      const doc = host.document;
      history.reset(doc);
      setLoadingState({ isLoading: false, parseError: null });
      const embeddedFamilies = getEmbeddedFontFamilies(doc.package.fontTable);
      const documentFonts = [
        ...getRenderableDocumentFonts(doc, { embeddedFamilies }),
        ...selectRenderableFonts(host.referencedFonts, { embeddedFamilies }),
      ];
      setDocumentFonts(
        [...new Map(documentFonts.map((font) => [font.name.toLowerCase(), font])).values()]
      );
      // This document's own claim, taken before the previous one is given
      // back: a face both documents embed is registered once, and releasing
      // the old claim first would take it away from under this load.
      const owner = createBufferFontOwner();
      const previousOwner = embeddedFacesRef.current;
      embeddedFacesRef.current = owner;
      void loadEmbeddedFonts(
        doc.package.fontTable,
        host.embeddedFonts,
        host.fontTableRelationshipsXml,
        owner
      )
        .catch((error) => {
          console.warn('Failed to load embedded document fonts:', error);
        })
        .then(() => {
          // A load that was superseded mid-flight holds faces nothing will
          // ever draw; give its claim back rather than keep it.
          if (!loadGeneration.isCurrent(generation) && embeddedFacesRef.current === owner) {
            embeddedFacesRef.current = null;
          }
          if (embeddedFacesRef.current !== owner) releaseBufferFontFaces(owner);
          // The previous document's claim goes back only now, so a face both
          // documents embed stays registered throughout, on this one's claim.
          if (previousOwner) releaseBufferFontFaces(previousOwner);
        })
        .then(() =>
          Promise.all([loadFontsWithMapping(host.referencedFonts), loadDocumentFonts(doc)])
        )
        .catch((error) => {
          console.warn('Failed to load document fonts:', error);
        });
    },
    [loadGeneration, history, setDocumentFonts, setLoadingState]
  );

  const failHostDocument = useCallback(
    (error: Error, generation: number) => {
      if (!loadGeneration.complete(generation)) return;
      // The document this replaced is gone from the screen either way, so its
      // faces go back now rather than waiting for the next one that parses.
      releaseEmbeddedFaces();
      setYrsSeedDocument(null);
      setYrsSeedBytes(null);
      setLoadingState({ isLoading: false, parseError: error.message });
      onError?.(error);
    },
    [loadGeneration, onError, releaseEmbeddedFaces, setLoadingState]
  );

  const isCurrentLoad = useCallback(
    (generation: number) => loadGeneration.isCurrent(generation),
    [loadGeneration]
  );

  // React to documentBuffer / document prop changes.
  useEffect(() => {
    // External-content mode: the caller populates the document directly —
    // skip the load.
    if (externalContent) return;

    if (!documentBuffer) {
      if (initialDocument) {
        loadParsedDocument(initialDocument);
      }
      return;
    }

    loadBuffer(documentBuffer);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [documentBuffer, initialDocument, externalContent]);

  // Extract any baked-in comments from the document model on first load.
  // Bumps the shared comment/revision ID counter above all loaded IDs so new
  // comments and tracked changes don't collide with existing ones (they
  // share the OOXML ID space).
  useEffect(() => {
    if (commentsLoadedRef.current) return;
    const doc = history.state;
    if (!doc) return;
    commentsLoadedRef.current = true;
    const bodyComments = doc.package?.document?.comments;
    if (bodyComments && bodyComments.length > 0) {
      setComments(bodyComments);
      setShowCommentsSidebar(true);
    }
    // New Yrs revisions have replica-stable string IDs; the numeric OOXML
    // comment allocator only needs to stay above loaded comment/reply IDs.
    commentIdAllocator.seedAbove(
      (bodyComments ?? []).reduce((max, comment) => Math.max(max, comment.id), 0)
    );
  }, [
    history.state,
    pagedEditorRef,
    setComments,
    setShowCommentsSidebar,
    commentsLoadedRef,
    commentIdAllocator,
  ]);

  useEffect(
    () => () => {
      loadGeneration.invalidate();
    },
    [loadGeneration]
  );

  return {
    loadParsedDocument,
    loadBuffer,
    yrsSeedDocument,
    yrsSeedBytes,
    yrsSeedGeneration,
    isCurrentLoad,
    acceptHostDocument,
    failHostDocument,
  };
}
