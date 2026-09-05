# @betteroffice/rust-crates

## 0.1.0

### Minor Changes

- c3713b9: A save that does not write now says which kind of thing went wrong, rather than reporting every one of them as a change the writer cannot express. `EditError::Unprojectable` kept only the cases it names — a formatting patch, a moved shape, an added slide — and four new variants took the rest: `Unsavable` for a replica opened from a collaborative update, which never carried the source bytes and can never be written to a file; `WriteLimit` for an edit larger than one save may write; `WriteFailed` for a write that broke or a writer that reached a state it does not hold; and `VerificationFailed` for bytes that did not read back as the deck they were planned from.

  The distinction is what lets a host answer correctly. `EditError::save_fault` returns a `SaveFault` whose `undoing_helps` is true for exactly one of them, so an offer to abandon edits is only ever made when undoing the named change would let the same save through. A broken write and a failed verification are the writer's problem, not the edit's, and the work has to survive them.

  `saveBytes` across the wasm boundary now rejects with an `Error` carrying `code`, `reason` and `undoingHelps` properties, and `@betteroffice/pptx` exports `saveFault` to read them. Classify on the code: `saveFault` returns `null` for anything the writer did not classify — a disposed handle, a panic, an error that merely quotes the writer's wording — so a host cannot be talked into discarding work by a message that reads like a refusal.

  Breaking for Rust callers that match on `EditError::Unprojectable` to detect a failed save: the cases above now arrive as their own variants. Match on `save_fault` instead.

### Patch Changes

- d9841a7: A function given a range, or an array function's result, where it takes one value now runs once per cell and spills: `=LEN(B1:B9)`, `=IF(A1:A9>2,"big","small")`, `=VLOOKUP(A1:A9,Table,2,FALSE)`, `=COUNTIF(B1:B9,UNIQUE(B1:B9))`, `=LARGE(A1:A9,SEQUENCE(3))`, and the conditions written with one — `=FILTER(A1:B9,ISNUMBER(SEARCH("x",B1:B9)))` — all answer instead of `#VALUE!`. Parameters that take a range, like `SUM`'s, are read whole as before.
- 66758d9: A spilled result now looks like one. Selecting any cell it fills outlines the whole result, the formula bar shows the formula that wrote it rather than the bare number, and typing into one of those cells says which cell to edit instead of failing silently.
- 5c0d5db: The formula engine can produce a rectangle of values, and `SEQUENCE` produces one. Nothing spills it into the sheet yet, so a cell holding such a formula reports `#VALUE!` — the function exists, but what it makes needs more room than one cell.
- f8c1123: Authoritative text measurement now reports one paint slice per run instead of one per character, which is what the display list, the DOM mirror and the layout engine all size themselves against.

  `advance_metadata` emitted a `bidiSlices` entry for every shaped cluster. The display list turned each entry into its own `glyphRun` primitive carrying the full attribute payload — colour, block key, document offsets, line indices and the resolved CSS face — so a plain paragraph of prose cost roughly 370 bytes of JSON per character. Consecutive clusters that share a run, a bidi level and a contiguous character span now fold into one slice, and the per-glyph advances they carried survive inside the primitive's `glyphs` array, so paint and hit geometry are unchanged. Letter-spaced runs keep their clusters apart: tracking lives in the per-cluster advance and would not survive re-shaping a merged span.

  Justification counts its space pool in characters now rather than in space-only paint items, which is the same number before coalescing and the right one after.

  Each placed glyph also carried a `logicalOrder` and a `bidiLevel` that only ever repeated the values on its own glyph run, and nothing read them. They are gone from the glyph contract.

  On a 250-page document of plain prose the display list drops from 229MB to 39MB, the renderer's wasm memory from 2174MB to 440-556MB, its JS heap from 690MB to 154-196MB, and the layout pass runs in 2.3s instead of 3.9s.

- 5d2afd1: Saving a workbook keeps every defined name it holds. The writer walked the source entries and the model's names in lockstep, so the first pair that disagreed dropped its entry and every later one with it — reordering three names left one. Names are now matched by name rather than by position, and a name the model gained is written instead of left out.
- 66c7312: Chart parsing and plot geometry now live in `betteroffice-drawingml` behind a `chart` cargo feature, shared by the docx, xlsx and pptx engines instead of being duplicated per format.

  Breaking, for Rust consumers of `betteroffice-drawingml` only. `ChartAxes`, `ChartAxis`, `ChartLegend`, `ChartMarker`, `ChartPoint`, `ChartPlotGroup` and `ChartSeries` were re-exported at the crate root in 0.0.4 and are no longer. Enable the `chart` feature and import them from `ooxml_drawingml::chart`. The feature is off by default, so a crate that does not draw charts no longer compiles the geometry engine.

- eba6194: A row whose cells run past the last column is refused when the workbook is read, the way an out-of-range reference spelled out in full always was. Counting on from the previous cell used to place a cell in column 16384, which SpreadsheetML cannot express: the workbook opened, and saving it wrote an `r="XFE1"` that neither Excel nor this parser could read back.
- ee39619: The editor now holds a lot less memory, without changing what it renders.

  The DOCX editing core is compiled once per document instead of twice. The resident engine runs in a worker, and that worker used to fetch and compile the same 11MB wasm binary the main thread had already compiled — two independent compilations, and two sets of machine code, of identical bytes. The main thread now hands the worker its compiled `WebAssembly.Module` before the first request; a worker that cannot be given one still loads the asset itself, so nothing depends on the handshake succeeding. Opening a document is also faster by however long that second compilation took.

  Bundled font bytes are no longer cached forever. Every consumer copies out of these buffers — the browser into its font system, the Rust font store into wasm — so the cache was holding a third copy of everything the document had ever touched, up to ~35MB once CJK faces were involved. It now keeps a 24MB working set and drops the least recently used faces; registration is memoized by face, so an eviction never causes a face to be registered with the engine twice.

  Workbook undo history is bounded and no longer copies per-cell state on every edit. Each committed transaction recorded two full copies of the per-cell shared-string provenance, on a stack with no depth limit — so a long editing session grew without bound, and every keystroke paid a copy proportional to the number of string cells in the workbook. The two sides of a history entry are now shared handles, copied only when a structural op (sheet add/remove, row/column insert/delete) actually changes them, which cell edits never do. History is capped at 200 transactions; undo past that depth is no longer offered, as in Excel and Word.

  The hidden textarea that carries IME composition no longer requests spell checking. It is empty except mid-composition, so the check had nothing to judge.

- 1617da6: An argument left out between two commas now parses. `=SORT(A1:B9,,,TRUE)` sorts by column and `=VLOOKUP(A1,B:C,2,)` matches exactly, both of which were a syntax error before — and a gap survives being rewritten when rows or columns move, rather than closing up and changing what the formula asks.
- 08a3f7c: An operator applied to a range or to an array function's result now runs cell by cell — `=A1:A4*2` spills four values and `B1:B4="x"` is the mask `FILTER` is written with — and any function that reads a range reads such a result the same way, so `=SUM(FILTER(…))`, `=COUNTA(UNIQUE(…))`, `=ROWS(FILTER(…))` and `=TEXTJOIN(",",TRUE,UNIQUE(…))` answer instead of refusing. A bare range in a cell spills.
- 74a329d: SORT, SORTBY, FILTER and UNIQUE now answer with a rectangle that spills into the sheet, and compose with each other and with SEQUENCE. Blank cells sort after everything in either direction, a result with nothing left in it reports `#CALC!` rather than a stale rectangle, and a cell holding `#CALC!` in a file Excel wrote is read back as that error instead of as the text of it.
- 69f35ba: An array formula spills its result into the sheet. `=SEQUENCE(3)` in A1 fills A1:A3, a formula reading one of those cells is calculated after the one that wrote it, anything standing in the way makes the result `#SPILL!` rather than overwriting it, and a result that shrinks gives back the cells it no longer reaches. The region is written as `<f t="array" ref="…">`, so a saved workbook reopens with its spilled cells still belonging to the formula.
- ab6e094: The cells an array formula filled belong to it. Typing into one is refused rather than torn out from under the formula, and inserting, deleting or sorting across a spilled result is refused the same way. Replacing the formula itself is allowed and takes the result with it.
- 692f2c7: Upgrade the XLSX raster backend to tiny-skia 0.12 with deterministic PNG encoding.
- b51086c: Undo in the browser now takes back a word rather than a letter. The undo manager groups edits that land inside a 500 ms window, but the wasm build had no clock to measure that window with and used a counter that advanced 501 ms on every reading — so every keystroke fell outside the previous one's window and became its own undo step. Taking back a five-letter word cost five presses. The docx and pptx engines now read `Date.now` through the JS boundary, and one press takes back the run as it was typed.

  Workbooks keep one press to one cell, deliberately. A spreadsheet commits an edit per cell on Enter or Tab, and grouping by time would take back every cell filled in the same half second; callers that want several edits undone together still say so through the batch entry point.

- f5d1b03: Saving a workbook now preserves the parts the model does not represent — charts, drawings, pivot tables, comments, macros, custom XML and their relationships — instead of rebuilding the package and dropping them. Sheets you did not touch are copied through byte for byte, so an edit on one sheet no longer strips hidden rows, outline levels, rich inline strings or shared-formula attributes from the rest of the workbook. The stylesheet is left alone unless styles actually change, and adding a format now patches one pool entry instead of rewriting every pool. Chartsheets keep their type, freeze-pane and hyperlink edits reach both the worksheet and its relationship part, and an edited save drops the stale calculation chain so Excel recalculates on open. Cells keep the shared-string entry they were authored against, and keep it through row and column edits that move them. A new occurrence of the same text is written without borrowing an existing rich-text entry. A save is refused when removing only some duplicate entries would leave too few entries for the distinct formatting still used by cells. Print areas and print titles are rewritten by row and column edits, not left pointing at the old ranges.

  Known limits. A sheet you edit is still reserialized from the model, so unmodeled row, column and cell markup on that one sheet is lost. Autofilter, data-validation, conditional-formatting, table and sparkline ranges on an edited sheet remain at their source coordinates, as do internal hyperlink locations outside the formula parser. Editing an existing style pool entry regenerates that entry from the modeled subset. Sheet rename, sheet removal and row or column edits are refused while a chart or pivot table is preserved, because their references cannot be rewritten; formulas naming a removed sheet keep that name instead of collapsing to `#REF!`. A row or column edit is also refused when a defined name aimed at that sheet is beyond the reference rewriter — a whole-row or whole-column reference nested inside a function, a structured table reference, or anything else the formula parser rejects. In a workbook with multiple sheets, the edit is refused for a workbook-scoped name with unqualified cell references because those references bind to whichever sheet uses the name. Collaborative sessions compare only the modeled workbook, so two peers holding the same cells but different charts or macros still accept each other as the same base. Replaying collaborative history drops which shared-string entry each cell was authored against; when an edited sheet is saved afterward, those cells are written as plain inline text rather than assigned another entry's formatting.

- 0ac3c46: Workbooks gain freeze panes, whole-row sort, auto filters with hidden rows, and classic cell comments across the model, the operation log, the reader and writer, and the renderer. Sorting carries a row's comments, hyperlinks and hidden state with it and rewrites relative formula references the way copying a row does, refusing rather than corrupting when a merged range or a multi-row hyperlink would be torn apart. Filters never hide the header row that owns their dropdown, and rows you hid by hand stay hidden when a filter changes or is cleared. Comments round trip through `xl/comments*.xml` alongside the legacy VML drawing Excel needs to draw their indicators, and a worksheet's header and footer artwork survives editing them. Filter criteria match a formula cell by its calculated value, and cell reads expose that same text so a filter dropdown can offer it.

  Filter criteria the engine cannot evaluate — custom comparisons, top ten, dynamic, colour and date groupings — are carried through verbatim instead of being dropped, constrain nothing while the workbook is open, and survive edits to the other columns of the same filter. A note attached to an otherwise empty cell now extends the sheet's extent, so its indicator is reachable.

  The collaboration schema advances to version 9 and upgrades version 3 through 8 snapshots when read, so a client on this release cannot share a collaboration room with an older one: upgrade every peer together.

  Known limits. SpreadsheetML records only that a row is hidden, never why, so a manual hide that also fails the active filter is re-attributed to the filter when the file is reopened. The criteria a filter stores are matched against raw cell text rather than the formatted text Excel writes, and rewriting a filter drops a column's own `hiddenButton` and `showButton` attributes. Comment text is kept as plain text, so editing a note collapses its own rich runs and phonetic properties, and a note you add is plain — the notes you leave alone keep their original markup, including when a sort or a row edit moves them. Two notes that share an author and text but differ in formatting can trade markup if they swap places, because nothing in the model tells them apart. The drawing that positions the notes keeps each untouched note's box, colours and visibility too, following a note that moves by its anchor alone — so a relocated note's box may sit over its old cells until Excel lays it out again, and a note you add gets the default yellow hidden box. Threaded comments are preserved as unmodeled parts but are not represented, so editing the classic comments a workbook falls back to can leave the threaded part stale. Sorting leaves row heights and outline levels at their original positions.

## 0.0.4

### Patch Changes

- 5c9a482: ArrowUp/ArrowDown move the caret by visual line with persistent goal-X (including across paragraphs, pages, columns, and into tables), and content below tables is clickable and editable again.
- 5c9a482: Collaborative presence: remote collaborator carets and selections render as colored overlays with name flags, anchored by yrs sticky indices so they rebase exactly under concurrent edits; carets follow remote typing instantly by inferring position from document updates.
- 5c9a482: Opening a document now seeds the collaborative session directly in the Rust engine instead of materializing the full TypeScript document model and projecting it; the TS model is built lazily only where the public API still exposes it, and the internal DrawingML host package is dissolved.
- 5c9a482: Remote collaborators' edits no longer move the local viewport: relayouts triggered by remote updates anchor to the topmost visible line via yrs sticky positions and compensate the scroll offset, while caret scrolling fires only for local actions. Anchoring holds across page boundaries too, so text overflowing onto a new page (or pulling back off one) no longer jumps the viewport for either the author or a viewer.
- 5212690: Google Slides-style editor toolbar for the PPTX editor: new-slide split button
  with layout picker, undo/redo, zoom, select and text-box tools, and contextual
  text formatting that also applies to whole shapes on selection. Text formatting
  now spans paragraph boundaries as a single undoable operation, double/triple
  click select word/paragraph, and roundRect corners render circular per the
  OOXML adj value instead of stretching with the shape.
- b87185f: Shape insertion and styling: a Slides-style shape picker inserts preset
  geometries (rectangles, ellipse, polygons, stars, arrows, chevron) by click
  or drag, and selected shapes get contextual fill, border color, border width,
  and corner-radius controls backed by new undoable, collaboration-native
  addShape/setShapeFill/setShapeStroke/setShapeAdjust engine operations.
- 4e04087: Formulas referencing defined names now resolve correctly, frozen panes render, and hyperlinks survive the round trip. The collaboration schema advances to version 5 and upgrades version 3 and 4 snapshots when read, so a client on this release cannot share a collaboration room with an older one: upgrade every peer together.
- 0d3baa1: Collaborative presence: remote collaborators' cell and range selections render as colored outlines with name flags, plus toolbar avatar chips; worksheets expose stable collaborative ids so presence survives sheet renames.

## 0.0.3

### Patch Changes

- b34bb01: Docx typing hot path is 7x faster (resident region fast path, memoized font parsing, direct frame-delta encoding, incremental worker sync); pages no longer remount and flash on remote or structural edits; page bitmaps are windowed to the viewport on long documents; the caret is painted by the renderer in the same frame as the glyphs while typing and blinks in the DOM at idle.
- 793b761: Render pending proposals as Word-style tracked changes: struck-through old
  values with a red run highlight, new values in green with a dashed underline
  and green run highlight, laid out side by side or new-over-old and following
  cell alignment. Proposal staging recalculates the formula graph and ghosts
  downstream dependents whose computed values change, proposal edits can carry
  a number format, and no-op proposals render unmarked.
- c6ad184: Add a Google Sheets-style toolbar to the XLSX editor backed by new engine
  APIs for range styling, number formats, selection-format aggregation, format
  painting, merge queries, and history state. Formatting is fully collaborative
  through a content-addressed style catalog (collaboration schema v3; v2 state
  does not migrate). Merging replaces intersecting ranges like Excel, parsing
  repairs overlapping merges in third-party files, and display-list font fields
  now serialize correctly so styled text renders with its real font, size, and
  weight.
- 793b761: Pending agent proposals render as in-cell tracked-change ghosts painted by the engine: the new value in green above the old value struck through in red, repainting immediately on propose, accept, and reject. Display-list text commands now serialize camelCase so cell fonts, sizes, and strike/underline offsets reach the canvas, and uninstalled workbook fonts fall back to sans-serif instead of the browser serif default.

## 0.0.2

### Patch Changes

- a34e721: Add deterministic Yrs replicas, bounded and validated sync-v1 exchange, a
  transport-agnostic npm collaboration provider, and React peer-update repainting.
  Collaborative sessions support nonstructural cell and dimension edits; inverse-op
  undo and redo remain disabled until a Yrs-aware undo manager can preserve
  concurrent edits.

## 0.0.1

### Patch Changes

- c92706e: Publish the Rust XLSX engine crates under BetterOffice names through the shared release workflow.
