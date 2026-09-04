---
"@betteroffice/xlsx": patch
"@betteroffice/rust-crates": patch
---

A row whose cells run past the last column is refused when the workbook is read, the way an out-of-range reference spelled out in full always was. Counting on from the previous cell used to place a cell in column 16384, which SpreadsheetML cannot express: the workbook opened, and saving it wrote an `r="XFE1"` that neither Excel nor this parser could read back.
