---
"@betteroffice/xlsx": patch
"@betteroffice/rust-crates": patch
---

Workbooks gain freeze panes, whole-row sort, auto filters with hidden rows, and classic cell comments across the model, the operation log, the reader and writer, and the renderer. Sorting carries a row's comments, hyperlinks and hidden state with it and rewrites relative formula references the way copying a row does, refusing rather than corrupting when a merged range or a multi-row hyperlink would be torn apart. Filters never hide the header row that owns their dropdown, and rows you hid by hand stay hidden when a filter changes or is cleared. Comments round trip through `xl/comments*.xml` alongside the legacy VML drawing Excel needs to draw their indicators, and a worksheet's header and footer artwork survives editing them. Filter criteria match a formula cell by its calculated value, and cell reads expose that same text so a filter dropdown can offer it.

The collaboration schema advances to version 8 and upgrades version 3 through 7 snapshots when read, so a client on this release cannot share a collaboration room with an older one: upgrade every peer together.

Known limits. SpreadsheetML records only that a row is hidden, never why, so a manual hide that also fails the active filter is re-attributed to the filter when the file is reopened. Filter criteria beyond a value list — custom comparisons, top ten, dynamic, colour, icon and date groupings — parse as no criteria and are dropped if that filter is later edited, and the criteria a filter stores are matched against raw cell text rather than the formatted text Excel writes. Comment text is kept as plain text, so rich formatting inside a note collapses when the note is edited. Threaded comments are preserved as unmodeled parts but are not represented, so editing the classic comments a workbook falls back to can leave the threaded part stale. Sorting leaves row heights and outline levels at their original positions.
