# @betteroffice/docx-fonts

## 0.0.2

### Patch Changes

- ee39619: The editor now holds a lot less memory, without changing what it renders.

  The DOCX editing core is compiled once per document instead of twice. The resident engine runs in a worker, and that worker used to fetch and compile the same 11MB wasm binary the main thread had already compiled — two independent compilations, and two sets of machine code, of identical bytes. The main thread now hands the worker its compiled `WebAssembly.Module` before the first request; a worker that cannot be given one still loads the asset itself, so nothing depends on the handshake succeeding. Opening a document is also faster by however long that second compilation took.

  Bundled font bytes are no longer cached forever. Every consumer copies out of these buffers — the browser into its font system, the Rust font store into wasm — so the cache was holding a third copy of everything the document had ever touched, up to ~35MB once CJK faces were involved. It now keeps a 24MB working set and drops the least recently used faces; registration is memoized by face, so an eviction never causes a face to be registered with the engine twice.

  Workbook undo history is bounded and no longer copies per-cell state on every edit. Each committed transaction recorded two full copies of the per-cell shared-string provenance, on a stack with no depth limit — so a long editing session grew without bound, and every keystroke paid a copy proportional to the number of string cells in the workbook. The two sides of a history entry are now shared handles, copied only when a structural op (sheet add/remove, row/column insert/delete) actually changes them, which cell edits never do. History is capped at 200 transactions; undo past that depth is no longer offered, as in Excel and Word.

  The hidden textarea that carries IME composition no longer requests spell checking. It is empty except mid-composition, so the check had nothing to judge.

- ea35453: Bundled faces registered with the DOM can now be given back.

  `registerBundledFontFace` installs a face under a caller-chosen CSS family name, and each registration parses its own copy of the font bytes — the same face registered under twenty family names is twenty parsed fonts. Nothing removed them, so a host registering per-document families accumulated every face of every document it had ever opened, and for CJK documents each of those copies is several megabytes.

  A caller now passes an owner from `createFontFaceOwner()`, and `releaseBundledFontFaces(owner)` gives back that owner's claim: a face no other owner still holds leaves `document.fonts` and loses its memo, so a later registration loads it again rather than resolving against a face that is gone. Two documents naming the same family share one registration and each hold it, so the one that closes first cannot take the face away from the one still open. Registering without an owner holds the face for the session, which is what the shared metric aliases want. A release that lands while a registration is still loading prevents that face being installed at all, instead of adding one nothing tracks.
