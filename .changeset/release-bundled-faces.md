---
"@betteroffice/docx-fonts": patch
---

Bundled faces registered with the DOM can now be given back.

`registerBundledFontFace` installs a face under a caller-chosen CSS family name, and each registration parses its own copy of the font bytes — the same face registered under twenty family names is twenty parsed fonts. Nothing removed them, so a host registering per-document families accumulated every face of every document it had ever opened, and for CJK documents each of those copies is several megabytes.

A caller now passes an owner from `createFontFaceOwner()`, and `releaseBundledFontFaces(owner)` gives back that owner's claim: a face no other owner still holds leaves `document.fonts` and loses its memo, so a later registration loads it again rather than resolving against a face that is gone. Two documents naming the same family share one registration and each hold it, so the one that closes first cannot take the face away from the one still open. Registering without an owner holds the face for the session, which is what the shared metric aliases want. A release that lands while a registration is still loading prevents that face being installed at all, instead of adding one nothing tracks.
