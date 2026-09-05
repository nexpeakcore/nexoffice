---
"@betteroffice/docx": patch
"@betteroffice/rust-crates": patch
---

Lowering asked the document whether an embed ends a paragraph twice for every paragraph, and built a `String` each time to compare it with one word. It asks once and compares in place.
