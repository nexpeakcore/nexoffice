# @betteroffice/pptx

## 0.1.0

### Minor Changes

- c3713b9: A save that does not write now says which kind of thing went wrong, rather than reporting every one of them as a change the writer cannot express. `EditError::Unprojectable` kept only the cases it names — a formatting patch, a moved shape, an added slide — and four new variants took the rest: `Unsavable` for a replica opened from a collaborative update, which never carried the source bytes and can never be written to a file; `WriteLimit` for an edit larger than one save may write; `WriteFailed` for a write that broke or a writer that reached a state it does not hold; and `VerificationFailed` for bytes that did not read back as the deck they were planned from.

  The distinction is what lets a host answer correctly. `EditError::save_fault` returns a `SaveFault` whose `undoing_helps` is true for exactly one of them, so an offer to abandon edits is only ever made when undoing the named change would let the same save through. A broken write and a failed verification are the writer's problem, not the edit's, and the work has to survive them.

  `saveBytes` across the wasm boundary now rejects with an `Error` carrying `code`, `reason` and `undoingHelps` properties, and `@betteroffice/pptx` exports `saveFault` to read them. Classify on the code: `saveFault` returns `null` for anything the writer did not classify — a disposed handle, a panic, an error that merely quotes the writer's wording — so a host cannot be talked into discarding work by a message that reads like a refusal.

  Breaking for Rust callers that match on `EditError::Unprojectable` to detect a failed save: the cases above now arrive as their own variants. Match on `save_fault` instead.

### Patch Changes

- b51086c: Undo in the browser now takes back a word rather than a letter. The undo manager groups edits that land inside a 500 ms window, but the wasm build had no clock to measure that window with and used a counter that advanced 501 ms on every reading — so every keystroke fell outside the previous one's window and became its own undo step. Taking back a five-letter word cost five presses. The docx and pptx engines now read `Date.now` through the JS boundary, and one press takes back the run as it was typed.

  Workbooks keep one press to one cell, deliberately. A spreadsheet commits an edit per cell on Enter or Tab, and grouping by time would take back every cell filled in the same half second; callers that want several edits undone together still say so through the batch entry point.

## 0.0.3

### Patch Changes

- 5212690: Google Slides-style editor toolbar for the PPTX editor: new-slide split button
  with layout picker, undo/redo, zoom, select and text-box tools, and contextual
  text formatting that also applies to whole shapes on selection. Text formatting
  now spans paragraph boundaries as a single undoable operation, double/triple
  click select word/paragraph, and roundRect corners render circular per the
  OOXML adj value instead of stretching with the shape.
- c134b2f: Collaborative presence: remote collaborators' shape selections render as colored outlines with name flags, with toolbar avatar chips and filmstrip dots showing which slide each peer is viewing.
- b87185f: Shape insertion and styling: a Slides-style shape picker inserts preset
  geometries (rectangles, ellipse, polygons, stars, arrows, chevron) by click
  or drag, and selected shapes get contextual fill, border color, border width,
  and corner-radius controls backed by new undoable, collaboration-native
  addShape/setShapeFill/setShapeStroke/setShapeAdjust engine operations.

## 0.0.2
