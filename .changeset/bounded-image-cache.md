---
"@betteroffice/docx": patch
"@betteroffice/docx-react": patch
---

Decoded document images no longer accumulate for the life of the editor.

The canvas image resolver cached every decoded bitmap forever: the cache had no size limit, and it was created once per editor mount, so switching documents carried the previous document's decoded images along — each one costing width × height × 4 bytes of memory that nothing would ever draw again. The cache is now bounded to a 64MB working set that drops the least recently used bitmaps (an evicted image simply decodes again on its next repaint — the sources are local blob/data URLs, so nothing is fetched), and the editor keys the cache to the document session, so loading a document starts it empty. An image whose own decoded size is larger than the whole budget is not retained at all, rather than sitting above the bound until some other image happens to settle.
