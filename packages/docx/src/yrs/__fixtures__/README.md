# Legacy CRDT seed

`legacy-seed-pre-cursor.bin` is the collaboration seed as it stood before
cursor seeding changed it (58,261 bytes; the current one is 59,841). It is a
compatibility record, not a generated artifact: **never regenerate it.**

It exists because the change that replaced it also regenerated the fixture the
compatibility test reads, so that test compared the new seeding path against
its own output and could not have failed. The two histories are not
interchangeable — same client id, clock 3814 against 3854, and merging them in
opposite orders disagrees about `hf:rId4`.

Nothing in the app merges them today: joining an existing room loads its state
instead of seeding (`useYrsCoreSession.ts`, `openDocx(bytes, !initialUpdate)`).
This file is what a real compatibility test needs when that changes, or when
seeding is rewritten again.
