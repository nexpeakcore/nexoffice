# Vendored dependencies

## yrs

A patched copy of [yrs](https://github.com/y-crdt/y-crdt) 0.27.3 (MIT), the
Rust implementation of the Yjs CRDT. Three crates build on it —
`docx-edit`, `betteroffice-xlsx`, `pptx-edit` — so it carries every format's
editing and collaboration.

### Why

`Text::insert*` reaches its index through `find_position`, which walks the
block list from the start of the text and caches nothing. Filling a story with
N chunks therefore costs 0 + 1 + ... + N block visits. Seeding a 2000-page
DOCX took 109.5s; every other phase of the open was linear and together under
2% of it.

### What changed

`vendor/yrs-cursor.patch` is the whole diff: **73 lines added, none removed or
edited**. It adds an opaque `TextCursor` plus three methods on the `Text`
trait — `cursor`, `insert_with_attributes_at`, `insert_embed_with_attributes_at`
— that reuse the position `insert` already advances internally rather than
rediscovering it. Same blocks, same order, one walk instead of N.

Seeding the same 2000-page document now takes 3.6s.

### Callers

Only `crates/docx-edit/src/raw.rs`, behind the crate's `yrs-cursor` feature.
The feature is off by default so `betteroffice-docx-edit` still compiles
against upstream yrs when published to crates.io — `[patch.crates-io]` applies
to this workspace only, never to a published crate. With the feature off the
crate keeps its original per-insert behaviour.

### Refreshing against a new upstream release

```
cargo package --list -p yrs   # or download the .crate from crates.io
# unpack over vendor/yrs, then:
git apply vendor/yrs-cursor.patch
```

### Exit plan

This exists to avoid waiting on someone else's release schedule, not to own a
CRDT library. When the cursor API lands upstream:

1. bump `yrs` in `docx-edit`, `betteroffice-xlsx` and `pptx-edit`;
2. delete `[patch.crates-io]` from the root `Cargo.toml`;
3. delete the `yrs-cursor` feature and its `cfg` in `raw.rs`;
4. delete `vendor/`.

Until then, re-apply the patch on every yrs upgrade and re-run
`bun run --filter './apps/demo' check:seeds`, which is what catches a change in
the blocks seeding produces.
