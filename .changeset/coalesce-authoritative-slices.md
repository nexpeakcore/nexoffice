---
"@betteroffice/rust-crates": patch
"@betteroffice/docx": patch
---

Authoritative text measurement now reports one paint slice per run instead of one per character, which is what the display list, the DOM mirror and the layout engine all size themselves against.

`advance_metadata` emitted a `bidiSlices` entry for every shaped cluster. The display list turned each entry into its own `glyphRun` primitive carrying the full attribute payload — colour, block key, document offsets, line indices and the resolved CSS face — so a plain paragraph of prose cost roughly 370 bytes of JSON per character. Consecutive clusters that share a run, a bidi level and a contiguous character span now fold into one slice, and the per-glyph advances they carried survive inside the primitive's `glyphs` array, so paint and hit geometry are unchanged. Letter-spaced runs keep their clusters apart: tracking lives in the per-cluster advance and would not survive re-shaping a merged span.

Justification counts its space pool in characters now rather than in space-only paint items, which is the same number before coalescing and the right one after.

On a 250-page document of plain prose the display list drops from 229MB to 56MB, the renderer's wasm memory from 2.2GB to 0.5GB, its JS heap from 690MB to 230MB, and the layout pass runs in 2.3s instead of 3.9s.
