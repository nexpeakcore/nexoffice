---
"@betteroffice/docx": patch
"@betteroffice/docx-react": patch
---

Fonts a DOCX embeds in itself are given back when the document closes.

Each embedded face was registered by minting an object URL over its bytes and appending an `@font-face` rule to the document head. Neither was ever undone, so every embedded font of every document opened in a session stayed resident — the bytes behind the object URL, the style element, and the browser's parsed face — even after the document was replaced. A face that failed to load leaked the same way, having created both before the failure.

`loadFontFromBuffer` and `loadEmbeddedFonts` now take an owner from `createBufferFontOwner()`, and `releaseBufferFontFaces(owner)` removes the rules and revokes the bytes for the faces that owner alone still holds. The editor takes an owner per load and releases it when the document is replaced or the editor unmounts. A face two documents both embed is registered once and claimed by both, so replacing a document while its fonts are still loading cannot leave the open document without the face it was told it already had. Only faces registered from raw buffers are affected: consumer-hosted faces (the `fonts` prop, `loadFontFromUrl`) outlive any one document and are left in place, as are Google-fetched families. Releasing a family's last buffer face also clears its loaded/registered markers, so re-opening a document that embeds it registers the bytes again rather than trusting a face that is gone.

Buffer-registered and URL-registered faces no longer share one dedupe key. A consumer calling `loadFontFromUrl` for a family, weight and style a document already embedded now installs its own rule instead of being deduped against the document's, so releasing the document leaves the consumer's font in place; the reverse holds too, and an embedded face registers even when a consumer face already claims that key.
