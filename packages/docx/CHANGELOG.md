# @betteroffice/docx

## 0.1.0

### Minor Changes

- 7831aa6: The renderer can now say which of its parts is holding a document's memory.

  Chromium runs Web Workers as threads inside the renderer process, so a host asking the platform how much memory the editor uses gets one number covering the resident layout engine's wasm heap, the four wasm modules on the main thread, the decoded-image cache and the JS heap together — with no way to tell which of them grew. A new `@betteroffice/docx/diagnostics` entry collects those readings: the wasm loaders publish each module's linear memory (read through the `WebAssembly.Memory`, so it tracks every grow), the canvas image resolver publishes its cached bytes, and the resident engine worker reports its own heap alongside the frames it already returns — the only way across a worker boundary. `memoryReport()` returns what each is holding, largest first, and `registerMemoryReader` lets a host add its own.

### Patch Changes

- 80e3d4b: Decoded document images no longer accumulate for the life of the editor.

  The canvas image resolver cached every decoded bitmap forever: the cache had no size limit, and it was created once per editor mount, so switching documents carried the previous document's decoded images along — each one costing width × height × 4 bytes of memory that nothing would ever draw again. The cache is now bounded to a 64MB working set that drops the least recently used bitmaps (an evicted image simply decodes again on its next repaint — the sources are local blob/data URLs, so nothing is fetched), and the editor keys the cache to the document session, so loading a document starts it empty. An image whose own decoded size is larger than the whole budget is not retained at all, rather than sitting above the bound until some other image happens to settle.

- f8c1123: Authoritative text measurement now reports one paint slice per run instead of one per character, which is what the display list, the DOM mirror and the layout engine all size themselves against.

  `advance_metadata` emitted a `bidiSlices` entry for every shaped cluster. The display list turned each entry into its own `glyphRun` primitive carrying the full attribute payload — colour, block key, document offsets, line indices and the resolved CSS face — so a plain paragraph of prose cost roughly 370 bytes of JSON per character. Consecutive clusters that share a run, a bidi level and a contiguous character span now fold into one slice, and the per-glyph advances they carried survive inside the primitive's `glyphs` array, so paint and hit geometry are unchanged. Letter-spaced runs keep their clusters apart: tracking lives in the per-cluster advance and would not survive re-shaping a merged span.

  Justification counts its space pool in characters now rather than in space-only paint items, which is the same number before coalescing and the right one after.

  Each placed glyph also carried a `logicalOrder` and a `bidiLevel` that only ever repeated the values on its own glyph run, and nothing read them. They are gone from the glyph contract.

  On a 250-page document of plain prose the display list drops from 229MB to 39MB, the renderer's wasm memory from 2174MB to 440-556MB, its JS heap from 690MB to 154-196MB, and the layout pass runs in 2.3s instead of 3.9s.

- f30eabb: Opening a long DOCX no longer takes minutes.

  Seeding the editing document reached every insertion index by walking yrs' block list from the start of the story — yrs 0.27 keeps no position cache — so filling a story with N chunks cost 0 + 1 + ... + N block visits. Seeding was quadratic in document length: 3.3s for 400 pages, 13.6s for 800, and 109.5s for 2000. Every other phase of the open (unzip, parse, lowering, op building) was already linear and together under 2% of the time.

  A seed batch is one long run of ascending inserts, so it now carries a cursor: the position each write already leaves behind is the one the next write needs. The same blocks are produced in the same order — the collaboration seed regenerates byte-identically — and the walk happens once instead of N times. The same documents now seed in 0.11s, 0.21s and 3.6s.

  The cursor API is not in upstream yrs yet; `vendor/yrs` carries it as 73 added lines over 0.27.3 while that change is proposed upstream, and `crates/docx-edit` keeps its original per-insert path behind a default-off `yrs-cursor` feature so the published crate still builds against upstream yrs.

- ee39619: The editor now holds a lot less memory, without changing what it renders.

  The DOCX editing core is compiled once per document instead of twice. The resident engine runs in a worker, and that worker used to fetch and compile the same 11MB wasm binary the main thread had already compiled — two independent compilations, and two sets of machine code, of identical bytes. The main thread now hands the worker its compiled `WebAssembly.Module` before the first request; a worker that cannot be given one still loads the asset itself, so nothing depends on the handshake succeeding. Opening a document is also faster by however long that second compilation took.

  Bundled font bytes are no longer cached forever. Every consumer copies out of these buffers — the browser into its font system, the Rust font store into wasm — so the cache was holding a third copy of everything the document had ever touched, up to ~35MB once CJK faces were involved. It now keeps a 24MB working set and drops the least recently used faces; registration is memoized by face, so an eviction never causes a face to be registered with the engine twice.

  Workbook undo history is bounded and no longer copies per-cell state on every edit. Each committed transaction recorded two full copies of the per-cell shared-string provenance, on a stack with no depth limit — so a long editing session grew without bound, and every keystroke paid a copy proportional to the number of string cells in the workbook. The two sides of a history entry are now shared handles, copied only when a structural op (sheet add/remove, row/column insert/delete) actually changes them, which cell edits never do. History is capped at 200 transactions; undo past that depth is no longer offered, as in Excel and Word.

  The hidden textarea that carries IME composition no longer requests spell checking. It is empty except mid-composition, so the check had nothing to judge.

- a3b4043: Fonts a DOCX embeds in itself are given back when the document closes.

  Each embedded face was registered by minting an object URL over its bytes and appending an `@font-face` rule to the document head. Neither was ever undone, so every embedded font of every document opened in a session stayed resident — the bytes behind the object URL, the style element, and the browser's parsed face — even after the document was replaced. A face that failed to load leaked the same way, having created both before the failure.

  `loadFontFromBuffer` and `loadEmbeddedFonts` now take an owner from `createBufferFontOwner()`, and `releaseBufferFontFaces(owner)` removes the rules and revokes the bytes for the faces that owner alone still holds. The editor takes an owner per load and releases it when the document is replaced or the editor unmounts. A face two documents both embed is registered once and claimed by both, so replacing a document while its fonts are still loading cannot leave the open document without the face it was told it already had. Only faces registered from raw buffers are affected: consumer-hosted faces (the `fonts` prop, `loadFontFromUrl`) outlive any one document and are left in place, as are Google-fetched families. Releasing a family's last buffer face also clears its loaded/registered markers, so re-opening a document that embeds it registers the bytes again rather than trusting a face that is gone.

  Buffer-registered and URL-registered faces no longer share one dedupe key. A consumer calling `loadFontFromUrl` for a family, weight and style a document already embedded now installs its own rule instead of being deduped against the document's, so releasing the document leaves the consumer's font in place; the reverse holds too, and an embedded face registers even when a consumer face already claims that key. Family-wide provenance also survives a release that lands while another face of the same family is still loading, so the subsetted-face glyph-coverage fallback is not silently switched off for it.

- b51086c: Undo in the browser now takes back a word rather than a letter. The undo manager groups edits that land inside a 500 ms window, but the wasm build had no clock to measure that window with and used a counter that advanced 501 ms on every reading — so every keystroke fell outside the previous one's window and became its own undo step. Taking back a five-letter word cost five presses. The docx and pptx engines now read `Date.now` through the JS boundary, and one press takes back the run as it was typed.

  Workbooks keep one press to one cell, deliberately. A spreadsheet commits an edit per cell on Enter or Tab, and grouping by time would take back every cell filled in the same half second; callers that want several edits undone together still say so through the batch entry point.

## 0.0.4

### Patch Changes

- 5c9a482: ArrowUp/ArrowDown move the caret by visual line with persistent goal-X (including across paragraphs, pages, columns, and into tables), and content below tables is clickable and editable again.
- 5c9a482: Collaborative presence: remote collaborator carets and selections render as colored overlays with name flags, anchored by yrs sticky indices so they rebase exactly under concurrent edits; carets follow remote typing instantly by inferring position from document updates.
- 5c9a482: Opening a document now seeds the collaborative session directly in the Rust engine instead of materializing the full TypeScript document model and projecting it; the TS model is built lazily only where the public API still exposes it, and the internal DrawingML host package is dissolved.
- 5c9a482: Remote collaborators' edits no longer move the local viewport: relayouts triggered by remote updates anchor to the topmost visible line via yrs sticky positions and compensate the scroll offset, while caret scrolling fires only for local actions. Anchoring holds across page boundaries too, so text overflowing onto a new page (or pulling back off one) no longer jumps the viewport for either the author or a viewer.

## 0.0.3

### Patch Changes

- b34bb01: Docx typing hot path is 7x faster (resident region fast path, memoized font parsing, direct frame-delta encoding, incremental worker sync); pages no longer remount and flash on remote or structural edits; page bitmaps are windowed to the viewport on long documents; the caret is painted by the renderer in the same frame as the glyphs while typing and blinks in the DOM at idle.

## 0.0.2
