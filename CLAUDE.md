@AGENTS.md

# NexOffice — Electron Desktop Office Suite

Fork of BetterOffice (Apache-2.0). Building an Electron + Rust office suite (Word + Excel clone).

## Architecture
- **Rust engines** in `crates/` (176K LOC, 25 crates) → compiled to WASM
- **TypeScript packages** in `packages/` — WASM wrappers + React components
- **Electron desktop shell** in `desktop/` — electron-vite + electron-builder
- Build tool: bun + wasm-pack
- Package manager: bun

## Rust Crate Structure
```
crates/
├── docx-parse/     — Parse OOXML DOCX (31K LOC)
├── docx-layout/    — Page layout engine (30K LOC)
├── docx-edit/      — CRDT editing engine (29K LOC)
├── betteroffice-docx/ — Public DOCX API
├── xlsx-model/     — XLSX data model
├── xlsx-parse/     — XLSX parser
├── xlsx-calc/      — Formula engine (110+ functions, dynamic arrays)
├── xlsx-ops/       — CRDT operations
├── xlsx-render/    — Canvas renderer
├── xlsx-raster/    — PNG export
├── betteroffice-xlsx/ — Public XLSX API
├── pptx-parse|edit|render/
├── ooxml-opc/      — Open Packaging Conventions
├── ooxml-drawingml/ — Charts, shapes, images
├── ooxml-text/     — Text shaping (rustybuzz)
```

## TypeScript Packages
```
packages/
├── docx/           — @betteroffice/docx WASM wrapper
├── docx-react/     — @betteroffice/docx-react (88 components)
├── docx-i18n/      — 10 languages
├── xlsx/           — @betteroffice/xlsx WASM wrapper
├── xlsx-react/     — @betteroffice/xlsx-react
├── xlsx-i18n/      — 10 languages
├── pptx/           — @betteroffice/pptx
├── pptx-react/     — @betteroffice/pptx-react
├── pptx-i18n/      — 10 languages
└── fonts/
```

## Key Commands
```bash
bun install                    # Install deps
bun run build:xlsx-wasm        # Build XLSX WASM (3.6MB)
bun run build:docx-wasm        # Build DOCX WASM (~19MB, 4 modules)
bun run build:pptx-wasm        # Build PPTX WASM
bun run rust:check             # cargo fmt + clippy + test
```

## Current State
- Rust engines: production-quality, 193 files in `docx-react`, 17 in `xlsx-react`
- Web demo app at `apps/demo` and `apps/web`
- Desktop app ships: `desktop-v1.0.0` through `v1.1.2` are tagged; **v1.1.2 has
  no installers attached** and `desktop/package.json` already reads 1.1.4, so
  the tree is two versions ahead of anything a user can install
- Phases 1-3 are done (Electron shell, XLSX sheet features, PPTX editing).
  Phase 4 (editor depth) is where new feature work goes

## What is blocked, and on what
These are the reasons work stops, not a backlog. Check them before promising a
release or a published package.

1. **Actions cannot open a release PR.** `default_workflow_permissions` is
   `read` and `can_approve_pull_request_reviews` is false, so `release.yml`
   fails at `changesets/action` on every push to main. **24 changesets are
   queued unpublished.** The fix is a repo setting the owner must flip:
   Settings → Actions → General → Workflow permissions → read+write, plus
   "Allow GitHub Actions to create and approve pull requests".
2. **Windows installers are unsigned.** `desktop/electron-builder.yml` has no
   `publisherName`, so electron-updater's signature check is a no-op and
   SmartScreen warns on install.
3. **`vendor/yrs` is a fork the repo calls TEMPORARY.** 1.5MB, 58 files, kept
   for a text-cursor API that makes DOCX seeding linear. Its exit condition is
   written in the root `Cargo.toml`. Until upstream carries the change, the
   `yrs-cursor` feature stays off by default — which means a plain
   `cargo test` measures the OLD quadratic seed. Pass `--features yrs-cursor`
   to any measurement of seeding or opening.
4. **The worker-authoritative refactor is 2 stages of 3.** The worker owns the
   document and the display list is page-windowed; the main-thread `bootstrap`
   path it replaced is still there alongside it.

## Phase 1-3, shipped
Electron + React shell with WASM bundles · file open/save via native dialogs ·
menu bar and shortcuts · spell check · word count · PDF export · page numbers ·
XLSX sort, filter, freeze panes, merge, comments · charts (render, author,
agent tool) · AI assistant with approval-gated writes · auto-update over Azure.

## Next, in order
1. **Typing on long documents.** ~160ms a keystroke on a 338-page file, of
   which lowering is ~53ms. Measured: reading the document out of yrs costs
   1ms of that; the other 30ms rebuilds all 2753 blocks when one paragraph
   changed. `lower_timing` in `crates/docx-edit/src/engine.rs` prints the
   split. The fix is to reuse lowered paragraphs by the content key
   `lower_story` already computes.
2. **Phase 4 editor depth**: vector PDF, spell check beyond English, threaded
   comments, incremental projection.
3. **Dynamic-array leftovers**: `CHOOSE` does not lift, `{1,2,3}` does not lex,
   `@` is not understood.

## Tech Stack for Electron
- Electron latest
- React 19 + TypeScript
- Tailwind CSS
- bun for package management
- electron-builder for packaging
- WASM loaded in renderer process

## Code Standards
- NO inline comments unless code is incomprehensible
- Conventional commits: type(scope): summary
- Rust edition 2024
- TypeScript strict mode
