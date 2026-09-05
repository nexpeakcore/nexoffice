---
"@betteroffice/xlsx": patch
"@betteroffice/rust-crates": patch
---

An argument left out between two commas now parses. `=SORT(A1:B9,,,TRUE)` sorts by column and `=VLOOKUP(A1,B:C,2,)` matches exactly, both of which were a syntax error before — and a gap survives being rewritten when rows or columns move, rather than closing up and changing what the formula asks.
