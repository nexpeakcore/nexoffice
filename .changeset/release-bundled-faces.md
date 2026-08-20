---
"@betteroffice/docx-fonts": patch
---

Bundled faces registered with the DOM can now be given back.

`registerBundledFontFace` installs a face under a caller-chosen CSS family name, and each registration parses its own copy of the font bytes — the same face registered under twenty family names is twenty parsed fonts. Nothing removed them, so a host registering per-document families accumulated every face of every document it had ever opened, and for CJK documents each of those copies is several megabytes.

`releaseBundledFontFaces(families)` removes the faces registered under the named families from `document.fonts` and clears their memo, so a later registration of the same family loads it again rather than resolving against a face that is gone. Families not named are untouched, which keeps shared aliases in place while per-document registrations come and go. A release that lands while a registration is still loading now prevents that face being installed at all, instead of adding one nothing tracks.
