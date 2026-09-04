---
"@betteroffice/docx-react": patch
---

The bundled faces a document's families resolve to are fetched as soon as the document is parsed, rather than in the middle of the first layout pass. Measurement meets a family only when it reaches text using it, so the bytes were fetched on discovery and the pass waited on them.
