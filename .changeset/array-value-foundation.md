---
"@betteroffice/xlsx": patch
"@betteroffice/rust-crates": patch
---

The formula engine can produce a rectangle of values, and `SEQUENCE` produces one. Nothing spills it into the sheet yet, so a cell holding such a formula reports `#VALUE!` — the function exists, but what it makes needs more room than one cell.
