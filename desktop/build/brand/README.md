# Nex Office Brand Assets — Option 1

This is the **reference-matched Option 1 system**: three dimensional-feeling documents stacked as the Suite master mark; a folded top-right corner; blue Docs, green Sheets, orange Slides, purple AI; restrained highlights and soft shadows.

## Approved palette

| Product / role | Hex |
|---|---|
| Suite foreground / Indigo | `#4F46E5` |
| Nex Docs / Blue | `#3B82F6` |
| Nex Sheets / Green | `#22C55E` |
| Nex Slides / Orange | `#F59E0B` |
| Nex AI / Purple | `#8B5CF6` |
| Wordmark / Ink | `#10142F` |
| Dark surface | `#10143C` |

## Contents

- **svg/** — clean transparent master vectors for Suite, Docs, Sheets, Slides, AI, wordmarks, and light/dark lockups.
- **png/** — transparent icon PNGs at 1024, 512, 256, 128, 64, and 32 px for every product.
- **app-icons/** — transparent app icon SVG/PNG sets for every product, plus Suite Android maskable variants.
- **favicon/** — Suite favicon as SVG, PNG (16/32/48/64), and a multi-resolution ICO.
- **preview/** — a quick Option 1 reference board for review; open the SVG in a modern browser.

## Light / dark use

- Use **nex-office-wordmark-dark.svg** and **nex-office-suite-lockup-light-background.svg** on white or pale backgrounds.
- Use **nex-office-wordmark-light.svg** and **nex-office-suite-lockup-dark-background.svg** on navy or other dark backgrounds.
- The transparent product icons are built for both backgrounds. Preserve the folds, gradient, and soft shadow; do not add outlines or recolor individual products.
- Prefer SVG in UI. Use the transparent PNGs where SVG is unavailable. At 16–24 px use the favicon rather than the full Suite stack.

## Production note

The wordmark uses common Avenir/Helvetica/Arial fallbacks to stay editable in SVG. Before a print-only or vendor release, outline the approved wordmark in the final layout file if the exact display font must be locked.
