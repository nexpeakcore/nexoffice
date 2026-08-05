import {
  loadBundledFontBytes,
  registerBundledFontFace,
  resolveLastResortFace,
  resolveMetricCompatFace,
  type BundledFontFace,
} from '@betteroffice/docx-fonts'
import type { PptxFontFace } from '@betteroffice/pptx'

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
