@AGENTS.md

# NexOffice — Electron Desktop Office Suite

Fork of BetterOffice (Apache-2.0). Building an Electron + Rust office suite (Word + Excel clone).

## Architecture
- **Rust engines** in `crates/` (149K LOC, 24 crates) → compiled to WASM
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
├── xlsx-calc/      — Formula engine (110+ functions)
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
├── docx-i18n/      — 12 languages
├── xlsx/           — @betteroffice/xlsx WASM wrapper
├── xlsx-react/     — @betteroffice/xlsx-react
├── xlsx-i18n/      — 12 languages
├── pptx/           — @betteroffice/pptx
├── pptx-react/     — @betteroffice/pptx-react
├── pptx-i18n/      — 12 languages
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
- Rust engines: Production-quality, build successfully
- React components: 180+ files for DOCX editor, 17 for XLSX
- Web demo app exists at apps/demo and apps/web
- Electron shell in `desktop/` — Phase 1, in progress

## Phase 1 Goals (Current)
Build Electron desktop wrapper + critical missing features:
1. Electron + React shell with WASM bundles
2. File Open/Save via native dialogs
3. Menu bar + keyboard shortcuts
4. Spell check (hunspell)
5. Word count
6. Export PDF
7. Page numbers
8. Sort & Filter (XLSX)
9. Freeze panes (XLSX)
10. Cell merge + comments (XLSX)

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
