---
"@betteroffice/xlsx": patch
"@betteroffice/rust-crates": patch
---

An operator applied to a range or to an array function's result now runs cell by cell — `=A1:A4*2` spills four values and `B1:B4="x"` is the mask `FILTER` is written with — and any function that reads a range reads such a result the same way, so `=SUM(FILTER(…))`, `=COUNTA(UNIQUE(…))`, `=ROWS(FILTER(…))` and `=TEXTJOIN(",",TRUE,UNIQUE(…))` answer instead of refusing. A bare range in a cell spills.
