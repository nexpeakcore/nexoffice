---
"@betteroffice/docx": minor
---

The renderer can now say which of its parts is holding a document's memory.

Chromium runs Web Workers as threads inside the renderer process, so a host asking the platform how much memory the editor uses gets one number covering the resident layout engine's wasm heap, the four wasm modules on the main thread, the decoded-image cache and the JS heap together — with no way to tell which of them grew. A new `@betteroffice/docx/diagnostics` entry collects those readings: the wasm loaders publish each module's linear memory (read through the `WebAssembly.Memory`, so it tracks every grow), the canvas image resolver publishes its cached bytes, and the resident engine worker reports its own heap alongside the frames it already returns — the only way across a worker boundary. `memoryReport()` returns what each is holding, largest first, and `registerMemoryReader` lets a host add its own.
