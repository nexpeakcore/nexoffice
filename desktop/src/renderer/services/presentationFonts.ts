import {
  loadBundledFontBytes,
  registerBundledFontFace,
  resolveLastResortFace,
  resolveMetricCompatFace,
  type BundledFontFace,
} from '@betteroffice/docx-fonts'
import type { DeckSnapshot, PptxFontFace, ShapeSnapshot } from '@betteroffice/pptx'

// The pptx renderer resolves a run's family against its registered faces and
// falls back to the first one registered, so Arial leads and the families a
// deck is most likely to name follow, each as its own metric-compatible face.
// The editor and the PDF exporter share this list: if they drift, the same deck
// resolves different faces on screen and on paper.
const PRESENTATION_FAMILIES = [
  'Arial',
  'Calibri',
  'Cambria',
  'Times New Roman',
  'Courier New',
] as const

const PRESENTATION_FONT_STYLES = [
  { bold: false, italic: false },
  { bold: true, italic: false },
  { bold: false, italic: true },
  { bold: true, italic: true },
] as const

// Mirrors crates/pptx-render/src/layout.rs: `MAX_FONTS` refuses registrations
// past 256 faces, and `resolve_style` ignores a family longer than 256 bytes.
export const MAX_PRESENTATION_FONT_FACES = 256
const MAX_FONT_FAMILY_BYTES = 256

export interface PresentationFontRequest {
  family: string
  bold: boolean
  italic: boolean
}

const familyEncoder = new TextEncoder()

/**
 * The registerable form of a deck-supplied family, or null when the engine
 * would never match it: blank, a theme reference the engine resolves itself,
 * or longer than the engine's family-length limit.
 */
export function normalizeFontFamily(family: string): string | null {
  const trimmed = family.trim()
  if (trimmed === '' || trimmed.startsWith('+')) return null
  if (familyEncoder.encode(trimmed).length > MAX_FONT_FAMILY_BYTES) return null
  return trimmed
}

export function resolvePresentationFace(
  family: string,
  bold: boolean,
  italic: boolean,
): BundledFontFace {
  return (
    resolveMetricCompatFace(family, bold, italic) ?? resolveLastResortFace(family, bold, italic)
  )
}

const faceBytes = new Map<string, Promise<Uint8Array>>()

/**
 * One copy of a face's bytes, shared by every family alias that resolves to it.
 * Unknown families all land on a handful of Liberation faces, so a deck naming
 * hundreds of them would otherwise hold hundreds of copies of the same bytes.
 */
export function loadPresentationFaceBytes(face: BundledFontFace): Promise<Uint8Array> {
  const cached = faceBytes.get(face.file)
  if (cached) return cached
  const pending = loadBundledFontBytes(face).then((bytes) => new Uint8Array(bytes.slice(0)))
  pending.catch(() => {
    if (faceBytes.get(face.file) === pending) faceBytes.delete(face.file)
  })
  faceBytes.set(face.file, pending)
  return pending
}

/**
 * A DOM family carrying one face and nothing else, so text pinned to it paints
 * from that face whatever weight or style is asked for: the face's own weight
 * matches exactly, and anything else the browser synthesizes. Chrome's synthetic
 * bold and oblique keep the real face's advances, which is what lets text the
 * engine measured with this face paint at the width layout gave it while still
 * looking bold. Registered lazily and shared: aliases are keyed by face file,
 * not by the family a deck happened to ask for.
 */
export function metricAliasFamily(face: BundledFontFace): string {
  return `NexOffice Metric ${face.file}`
}

export function registerMetricAlias(face: BundledFontFace): Promise<void> {
  return registerBundledFontFace(face, metricAliasFamily(face))
}

export function baseFontRequests(): PresentationFontRequest[] {
  return PRESENTATION_FAMILIES.flatMap((family) =>
    PRESENTATION_FONT_STYLES.map(({ bold, italic }) => ({ family, bold, italic })),
  )
}

let baseFaces: Promise<PptxFontFace[]> | null = null

/**
 * The faces every presentation starts from. Cached across editors, but a
 * rejection is evicted so a transient font failure does not outlive itself.
 */
export function loadBaseFontFaces(): Promise<PptxFontFace[]> {
  const cached = baseFaces
  if (cached) return cached
  const pending = Promise.all(
    baseFontRequests().map(async (request) => ({
      ...request,
      bytes: await loadPresentationFaceBytes(
        resolvePresentationFace(request.family, request.bold, request.italic),
      ),
    })),
  )
  pending.catch(() => {
    if (baseFaces === pending) baseFaces = null
  })
  baseFaces = pending
  return pending
}

/**
 * Every family the deck's own runs name, on top of the base set.
 *
 * The base set is five Latin families. A deck that names anything else — a
 * CJK family, an alias like Helvetica — has no face for it unless it is
 * collected from the deck itself, and the engine then measures that text with
 * a font that has none of its glyphs. Both the editor and the PDF exporter
 * read this, so the screen cannot resolve a run to one family while the export
 * resolves it to another.
 */
export function collectFontRequests(snapshot: DeckSnapshot): PresentationFontRequest[] {
  const requests = new Map<string, PresentationFontRequest>()
  for (const request of baseFontRequests()) {
    addFontRequest(requests, request.family, request.bold, request.italic)
  }
  for (const slide of snapshot.slides) {
    for (const shape of slide.shapes) collectShapeFonts(shape, requests)
    if (requests.size >= MAX_PRESENTATION_FONT_FACES) break
  }
  return [...requests.values()]
}

function addFontRequest(
  requests: Map<string, PresentationFontRequest>,
  family: string,
  bold: boolean,
  italic: boolean,
): void {
  if (requests.size >= MAX_PRESENTATION_FONT_FACES) return
  const normalized = normalizeFontFamily(family)
  if (normalized === null) return
  const key = `${normalized.toLowerCase()}|${bold ? 1 : 0}|${italic ? 1 : 0}`
  if (!requests.has(key)) requests.set(key, { family: normalized, bold, italic })
}

function collectShapeFonts(
  shape: ShapeSnapshot,
  requests: Map<string, PresentationFontRequest>,
): void {
  for (const story of shape.textStories) {
    for (const paragraph of story.paragraphs) {
      for (const run of paragraph.runs) {
        const family = run.style.fontFamily
        if (family === null) continue
        addFontRequest(requests, family, run.style.bold ?? false, run.style.italic ?? false)
      }
    }
  }
  for (const child of shape.children) collectShapeFonts(child, requests)
}

/**
 * Registers the faces this deck names but the base set does not carry, and
 * reports whether anything new arrived.
 *
 * The editor opens on the base set, because the families a deck names are only
 * knowable once it is parsed. Registering the rest afterwards is what stops the
 * screen measuring Chinese text with a Latin font while the PDF measures it
 * with the right one.
 *
 * Both registrations are needed and they are not the same one. The engine's is
 * what the layout measures with; the document's is what the canvas paints with,
 * because the canvas asks CSS for the family the display list names. Doing only
 * the first lays the text out at the right widths and then draws it in whatever
 * the browser substitutes.
 *
 * The editor cannot be handed these through its `fonts` prop instead: changing
 * it reopens the presentation, which would throw away everything typed since.
 */
export async function registerDeckFonts(
  handle: { registerFont: (face: PptxFontFace) => number },
  snapshot: DeckSnapshot,
  already: ReadonlyArray<PresentationFontRequest>,
): Promise<number> {
  const seen = new Set(
    already.map((request) => requestKey(request.family, request.bold, request.italic)),
  )
  let added = 0
  for (const request of collectFontRequests(snapshot)) {
    const key = requestKey(request.family, request.bold, request.italic)
    if (seen.has(key)) continue
    seen.add(key)
    try {
      const face = resolvePresentationFace(request.family, request.bold, request.italic)
      handle.registerFont({ ...request, bytes: await loadPresentationFaceBytes(face) })
      await registerBundledFontFace(face, request.family)
      await registerMetricAlias(face)
      added += 1
    } catch {
      // One family failing to load must not cost the deck the rest of them.
      continue
    }
  }
  return added
}

function requestKey(family: string, bold: boolean, italic: boolean): string {
  return `${family.toLowerCase()}|${bold ? 1 : 0}|${italic ? 1 : 0}`
}
