---
"@betteroffice/xlsx": patch
"@betteroffice/rust-crates": patch
---

A function given a range, or an array function's result, where it takes one value now runs once per cell and spills: `=LEN(B1:B9)`, `=IF(A1:A9>2,"big","small")`, `=VLOOKUP(A1:A9,Table,2,FALSE)`, `=COUNTIF(B1:B9,UNIQUE(B1:B9))`, `=LARGE(A1:A9,SEQUENCE(3))`, and the conditions written with one — `=FILTER(A1:B9,ISNUMBER(SEARCH("x",B1:B9)))` — all answer instead of `#VALUE!`. Parameters that take a range, like `SUM`'s, are read whole as before.
