---
"@betteroffice/xlsx": patch
"@betteroffice/rust-crates": patch
---

An array formula spills its result into the sheet. `=SEQUENCE(3)` in A1 fills A1:A3, a formula reading one of those cells is calculated after the one that wrote it, anything standing in the way makes the result `#SPILL!` rather than overwriting it, and a result that shrinks gives back the cells it no longer reaches. The region is written as `<f t="array" ref="…">`, so a saved workbook reopens with its spilled cells still belonging to the formula.
