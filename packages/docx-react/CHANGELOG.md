# @betteroffice/docx-react

## 0.1.0

### Patch Changes

- 80e3d4b: Decoded document images no longer accumulate for the life of the editor.

  The canvas image resolver cached every decoded bitmap forever: the cache had no size limit, and it was created once per editor mount, so switching documents carried the previous document's decoded images along — each one costing width × height × 4 bytes of memory that nothing would ever draw again. The cache is now bounded to a 64MB working set that drops the least recently used bitmaps (an evicted image simply decodes again on its next repaint — the sources are local blob/data URLs, so nothing is fetched), and the editor keys the cache to the document session, so loading a document starts it empty. An image whose own decoded size is larger than the whole budget is not retained at all, rather than sitting above the bound until some other image happens to settle.

- a3b4043: Fonts a DOCX embeds in itself are given back when the document closes.

  Each embedded face was registered by minting an object URL over its bytes and appending an `@font-face` rule to the document head. Neither was ever undone, so every embedded font of every document opened in a session stayed resident — the bytes behind the object URL, the style element, and the browser's parsed face — even after the document was replaced. A face that failed to load leaked the same way, having created both before the failure.

  `loadFontFromBuffer` and `loadEmbeddedFonts` now take an owner from `createBufferFontOwner()`, and `releaseBufferFontFaces(owner)` removes the rules and revokes the bytes for the faces that owner alone still holds. The editor takes an owner per load and releases it when the document is replaced or the editor unmounts. A face two documents both embed is registered once and claimed by both, so replacing a document while its fonts are still loading cannot leave the open document without the face it was told it already had. Only faces registered from raw buffers are affected: consumer-hosted faces (the `fonts` prop, `loadFontFromUrl`) outlive any one document and are left in place, as are Google-fetched families. Releasing a family's last buffer face also clears its loaded/registered markers, so re-opening a document that embeds it registers the bytes again rather than trusting a face that is gone.

  Buffer-registered and URL-registered faces no longer share one dedupe key. A consumer calling `loadFontFromUrl` for a family, weight and style a document already embedded now installs its own rule instead of being deduped against the document's, so releasing the document leaves the consumer's font in place; the reverse holds too, and an embedded face registers even when a consumer face already claims that key. Family-wide provenance also survives a release that lands while another face of the same family is still loading, so the subsetted-face glyph-coverage fallback is not silently switched off for it.

- be971f3: The bundled faces a document's families resolve to are fetched as soon as the document is parsed, rather than in the middle of the first layout pass. Measurement meets a family only when it reaches text using it, so the bytes were fetched on discovery and the pass waited on them.
- Updated dependencies [80e3d4b]
- Updated dependencies [f8c1123]
- Updated dependencies [f30eabb]
- Updated dependencies [ee39619]
- Updated dependencies [a3b4043]
- Updated dependencies [7831aa6]
- Updated dependencies [b51086c]
  - @betteroffice/docx@0.1.0
  - @betteroffice/docx-i18n@0.1.0

## 0.0.4

### Patch Changes

- 5c9a482: ArrowUp/ArrowDown move the caret by visual line with persistent goal-X (including across paragraphs, pages, columns, and into tables), and content below tables is clickable and editable again.
- 5c9a482: Collaborative presence: remote collaborator carets and selections render as colored overlays with name flags, anchored by yrs sticky indices so they rebase exactly under concurrent edits; carets follow remote typing instantly by inferring position from document updates.
- 5c9a482: Opening a document now seeds the collaborative session directly in the Rust engine instead of materializing the full TypeScript document model and projecting it; the TS model is built lazily only where the public API still exposes it, and the internal DrawingML host package is dissolved.
- 5c9a482: Remote collaborators' edits no longer move the local viewport: relayouts triggered by remote updates anchor to the topmost visible line via yrs sticky positions and compensate the scroll offset, while caret scrolling fires only for local actions. Anchoring holds across page boundaries too, so text overflowing onto a new page (or pulling back off one) no longer jumps the viewport for either the author or a viewer.
- Updated dependencies [5c9a482]
- Updated dependencies [5c9a482]
- Updated dependencies [5c9a482]
- Updated dependencies [5c9a482]
  - @betteroffice/docx@0.0.4
  - @betteroffice/docx-i18n@0.0.4

## 0.0.3

### Patch Changes

- Updated dependencies [b34bb01]
  - @betteroffice/docx@0.0.3
  - @betteroffice/docx-i18n@0.0.3

## 0.0.2

### Patch Changes

- eed05a6: Fix the published dependency ranges: 0.0.1 shipped the unresolved `workspace:*` protocol for `@betteroffice/docx` and `@betteroffice/docx-i18n`, which made `npm install @betteroffice/docx-react` fail. Ranges are now pinned to concrete versions at publish time.
  - @betteroffice/docx@0.0.2
  - @betteroffice/docx-i18n@0.0.2
