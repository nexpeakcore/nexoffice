---
"@betteroffice/docx": patch
"@betteroffice/docx-react": patch
---

Fonts a DOCX embeds in itself are given back when the document closes.

Each embedded face was registered by minting an object URL over its bytes and appending an `@font-face` rule to the document head. Neither was ever undone, so every embedded font of every document opened in a session stayed resident — the bytes behind the object URL, the style element, and the browser's parsed face — even after the document was replaced. A face that failed to load leaked the same way, having created both before the failure.

`releaseBufferFontFaces(families)` removes the rules and revokes the bytes for the named families, and the editor calls it when a document is replaced or the editor unmounts. Only faces registered from raw buffers are affected: consumer-hosted faces (the `fonts` prop, `loadFontFromUrl`) outlive any one document and are left in place, as are Google-fetched families. Releasing a family's last buffer face also clears its loaded/registered markers, so re-opening a document that embeds it registers the bytes again rather than trusting a face that is gone.
