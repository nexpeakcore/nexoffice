# @betteroffice/xlsx-react

## 0.0.9

### Patch Changes

- 66758d9: A spilled result now looks like one. Selecting any cell it fills outlines the whole result, the formula bar shows the formula that wrote it rather than the bare number, and typing into one of those cells says which cell to edit instead of failing silently.
- Updated dependencies [d9841a7]
- Updated dependencies [66758d9]
- Updated dependencies [5c0d5db]
- Updated dependencies [5d2afd1]
- Updated dependencies [eba6194]
- Updated dependencies [1617da6]
- Updated dependencies [08a3f7c]
- Updated dependencies [74a329d]
- Updated dependencies [69f35ba]
- Updated dependencies [ab6e094]
- Updated dependencies [692f2c7]
- Updated dependencies [f5d1b03]
- Updated dependencies [0ac3c46]
  - @betteroffice/xlsx@0.0.9
  - @betteroffice/xlsx-i18n@0.0.9

## 0.0.8

### Patch Changes

- 4e04087: Formulas referencing defined names now resolve correctly, frozen panes render, and hyperlinks survive the round trip. The collaboration schema advances to version 5 and upgrades version 3 and 4 snapshots when read, so a client on this release cannot share a collaboration room with an older one: upgrade every peer together.
- 0d3baa1: Collaborative presence: remote collaborators' cell and range selections render as colored outlines with name flags, plus toolbar avatar chips; worksheets expose stable collaborative ids so presence survives sheet renames.
- Updated dependencies [4e04087]
- Updated dependencies [47c37b0]
- Updated dependencies [0d3baa1]
  - @betteroffice/xlsx@0.0.8
  - @betteroffice/xlsx-i18n@0.0.8

## 0.0.7

### Patch Changes

- 793b761: Render pending proposals as Word-style tracked changes: struck-through old
  values with a red run highlight, new values in green with a dashed underline
  and green run highlight, laid out side by side or new-over-old and following
  cell alignment. Proposal staging recalculates the formula graph and ghosts
  downstream dependents whose computed values change, proposal edits can carry
  a number format, and no-op proposals render unmarked.
- c6ad184: Add a Google Sheets-style toolbar to the XLSX editor backed by new engine
  APIs for range styling, number formats, selection-format aggregation, format
  painting, merge queries, and history state. Formatting is fully collaborative
  through a content-addressed style catalog (collaboration schema v3; v2 state
  does not migrate). Merging replaces intersecting ranges like Excel, parsing
  repairs overlapping merges in third-party files, and display-list font fields
  now serialize correctly so styled text renders with its real font, size, and
  weight.
- Updated dependencies [793b761]
- Updated dependencies [c6ad184]
- Updated dependencies [793b761]
  - @betteroffice/xlsx@0.0.7
  - @betteroffice/xlsx-i18n@0.0.7

## 0.0.6

### Patch Changes

- a34e721: Add deterministic Yrs replicas, bounded and validated sync-v1 exchange, a
  transport-agnostic npm collaboration provider, and React peer-update repainting.
  Collaborative sessions support nonstructural cell and dimension edits; inverse-op
  undo and redo remain disabled until a Yrs-aware undo manager can preserve
  concurrent edits.
- 69d62f1: Refine the XLSX and PPTX editor toolbars with compact DOCX-style control rails,
  grouped icon actions, and responsive value fields.
- Updated dependencies [a34e721]
  - @betteroffice/xlsx@0.0.6
  - @betteroffice/xlsx-i18n@0.0.6

## 0.0.5

### Patch Changes

- Updated dependencies [e8678aa]
  - @betteroffice/xlsx@0.0.5

## 0.0.4

### Patch Changes

- 6a1ab98: Publish the spreadsheet packages as ESM-only and load the WebAssembly core as a separate asset.
- Updated dependencies [6a1ab98]
  - @betteroffice/xlsx@0.0.4

## 0.0.3

### Patch Changes

- 68d15b8: Fix `@betteroffice/xlsx-react` so its dependency on `@betteroffice/xlsx` resolves to the matching published version.
- Updated dependencies [68d15b8]
  - @betteroffice/xlsx@0.0.3
