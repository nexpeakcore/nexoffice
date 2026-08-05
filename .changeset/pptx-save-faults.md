---
"@betteroffice/pptx": minor
"@betteroffice/rust-crates": minor
---

A save that does not write now says which kind of thing went wrong, rather than reporting every one of them as a change the writer cannot express. `EditError::Unprojectable` kept only the cases it names — a formatting patch, a moved shape, an added slide — and four new variants took the rest: `Unsavable` for a replica opened from a collaborative update, which never carried the source bytes and can never be written to a file; `WriteLimit` for an edit larger than one save may write; `WriteFailed` for a write that broke or a writer that reached a state it does not hold; and `VerificationFailed` for bytes that did not read back as the deck they were planned from.

The distinction is what lets a host answer correctly. `EditError::save_fault` returns a `SaveFault` whose `undoing_helps` is true for exactly one of them, so an offer to abandon edits is only ever made when undoing the named change would let the same save through. A broken write and a failed verification are the writer's problem, not the edit's, and the work has to survive them.

`saveBytes` across the wasm boundary now rejects with an `Error` carrying `code`, `reason` and `undoingHelps` properties, and `@betteroffice/pptx` exports `saveFault` to read them. Classify on the code: `saveFault` returns `null` for anything the writer did not classify — a disposed handle, a panic, an error that merely quotes the writer's wording — so a host cannot be talked into discarding work by a message that reads like a refusal.

Breaking for Rust callers that match on `EditError::Unprojectable` to detect a failed save: the cases above now arrive as their own variants. Match on `save_fault` instead.
