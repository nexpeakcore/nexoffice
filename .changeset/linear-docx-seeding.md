---
"@betteroffice/docx": patch
---

Opening a long DOCX no longer takes minutes.

Seeding the editing document reached every insertion index by walking yrs' block list from the start of the story — yrs 0.27 keeps no position cache — so filling a story with N chunks cost 0 + 1 + ... + N block visits. Seeding was quadratic in document length: 3.3s for 400 pages, 13.6s for 800, and 109.5s for 2000. Every other phase of the open (unzip, parse, lowering, op building) was already linear and together under 2% of the time.

A seed batch is one long run of ascending inserts, so it now carries a cursor: the position each write already leaves behind is the one the next write needs. The same blocks are produced in the same order — the collaboration seed regenerates byte-identically — and the walk happens once instead of N times. The same documents now seed in 0.11s, 0.21s and 3.6s.

The cursor API is not in upstream yrs yet; `vendor/yrs` carries it as 73 added lines over 0.27.3 while that change is proposed upstream, and `crates/docx-edit` keeps its original per-insert path behind a default-off `yrs-cursor` feature so the published crate still builds against upstream yrs.
