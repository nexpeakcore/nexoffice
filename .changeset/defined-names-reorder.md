---
"@betteroffice/xlsx": patch
"@betteroffice/rust-crates": patch
---

Saving a workbook keeps every defined name it holds. The writer walked the source entries and the model's names in lockstep, so the first pair that disagreed dropped its entry and every later one with it — reordering three names left one. Names are now matched by name rather than by position, and a name the model gained is written instead of left out.
