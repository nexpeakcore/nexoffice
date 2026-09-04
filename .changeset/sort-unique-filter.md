---
"@betteroffice/xlsx": patch
"@betteroffice/rust-crates": patch
---

SORT, SORTBY, FILTER and UNIQUE now answer with a rectangle that spills into the sheet, and compose with each other and with SEQUENCE. Blank cells sort after everything in either direction, a result with nothing left in it reports `#CALC!` rather than a stale rectangle, and a cell holding `#CALC!` in a file Excel wrote is read back as that error instead of as the text of it.
