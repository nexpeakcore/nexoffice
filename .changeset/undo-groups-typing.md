---
"@betteroffice/docx": patch
"@betteroffice/pptx": patch
"@betteroffice/rust-crates": patch
---

Undo in the browser now takes back a word rather than a letter. The undo manager groups edits that land inside a 500 ms window, but the wasm build had no clock to measure that window with and used a counter that advanced 501 ms on every reading — so every keystroke fell outside the previous one's window and became its own undo step. Taking back a five-letter word cost five presses. The docx and pptx engines now read `Date.now` through the JS boundary, and one press takes back the run as it was typed.

Workbooks keep one press to one cell, deliberately. A spreadsheet commits an edit per cell on Enter or Tab, and grouping by time would take back every cell filled in the same half second; callers that want several edits undone together still say so through the batch entry point.
