import { registerBundledFontFace } from '@betteroffice/docx-fonts'
import {
  initWasm,
  openPresentation,
  paintSlide,
  type CanvasImageResolver,
  type DeckSnapshot,
  type PresentationHandle,
  type ShapeSnapshot,
  type SlideDisplayList,
} from '@betteroffice/pptx'
import { PRINT_PAGE_CAP } from '../../shared/ipc.js'
import {
  baseFontRequests,
  loadPresentationFaceBytes,
  MAX_PRESENTATION_FONT_FACES,
  metricAliasFamily,
  normalizeFontFamily,
  registerMetricAlias,
  resolvePresentationFace,
  type PresentationFontRequest,
} from '../services/presentationFonts.js'
import type { PageImage, PageSet } from './types.js'

const RASTER_SCALE = 2

// Slides carry their own size, so unlike a fixed paper page a deck can ask for
// a raster no canvas can back: past Chrome's limit the context silently stops
// drawing and `toDataURL` hands back a blank page. Every page is also held as a
// decoded bitmap until `printToPDF` runs, so the budget is shared across the
// export rather than granted per page.
const MAX_RASTER_EDGE = 8192
const MAX_PAGE_PIXELS = 4_194_304
const RASTER_PIXEL_BUDGET = 100_000_000

// The parser accepts a slide up to 1e15 EMU, which is a paper size no print
// engine will paginate: past roughly 3000px Chrome abandons the `@page` size and
// tiles the deck across default paper, and at the extreme `printToPDF` never
// returns. A slide past this prints scaled to fit, aspect ratio intact.
const MAX_PAGE_EDGE = 2_560

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

function collectShapeFonts(shape: ShapeSnapshot, requests: Map<string, PresentationFontRequest>): void {
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

function collectFontRequests(snapshot: DeckSnapshot): PresentationFontRequest[] {
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

interface LoadedFont {
  request: PresentationFontRequest
  bytes: Uint8Array
  metricFamily: string
}

interface RegisteredFace extends PresentationFontRequest {
  metricFamily: string
}

// Layout picks glyph advances from the faces registered with the engine while
// the canvas paints with `ctx.font`, so every face is installed on both sides
// under the family the deck asked for or the two disagree. A face that will not
// load is dropped instead of failing the export: the engine falls back for it,
// and `alignRunFaces` keeps the canvas on the same fallback.
async function loadFontFaces(requests: PresentationFontRequest[]): Promise<LoadedFont[]> {
  const loaded = await Promise.all(
    requests.map(async (request): Promise<LoadedFont | null> => {
      try {
        const face = resolvePresentationFace(request.family, request.bold, request.italic)
        const bytes = await loadPresentationFaceBytes(face)
        await registerBundledFontFace(face, request.family)
        await registerMetricAlias(face)
        return { request, bytes, metricFamily: metricAliasFamily(face) }
      } catch {
        return null
      }
    }),
  )
  return loaded.filter((entry): entry is LoadedFont => entry !== null)
}

function registerFonts(
  handle: PresentationHandle,
  loaded: LoadedFont[],
): Map<number, RegisteredFace> {
  const faces = new Map<number, RegisteredFace>()
  for (const { request, bytes, metricFamily } of loaded) {
    try {
      const id = handle.registerFont({
        family: request.family,
        bold: request.bold,
        italic: request.italic,
        bytes,
      })
      if (!faces.has(id)) faces.set(id, { ...request, metricFamily })
    } catch {
      continue
    }
  }
  return faces
}

// A run whose family the engine never received — a theme default, a family past
// the face cap, a face that failed to load — is measured with the fallback face
// but still reports its own bold and italic, so the canvas would paint a heavier
// face than the advances were taken from and overflow its box. Repointing those
// runs at the alias of the face the engine really used restores the advances,
// and the weight the run asked for survives as a synthesized one.
function alignRunFaces(list: SlideDisplayList, faces: Map<number, RegisteredFace>): void {
  for (const primitive of list.primitives) {
    if (primitive.kind !== 'textBox') continue
    for (const line of primitive.lines) {
      for (const run of line.runs) {
        const face = faces.get(run.fontId)
        if (!face) continue
        if (run.bold === face.bold && run.italic === face.italic) continue
        run.fontFamily = face.metricFamily
      }
    }
  }
}

function rasterScale(width: number, height: number, pageCount: number): number {
  const budget = Math.min(MAX_PAGE_PIXELS, RASTER_PIXEL_BUDGET / pageCount)
  const area = width * height
  const scale = area * RASTER_SCALE * RASTER_SCALE > budget
    ? Math.sqrt(budget / area)
    : RASTER_SCALE
  return Math.min(scale, MAX_RASTER_EDGE / width, MAX_RASTER_EDGE / height)
}

async function decodeMedia(
  handle: PresentationHandle,
  assetId: string,
): Promise<CanvasImageSource | null> {
  let bytes: Uint8Array
  try {
    bytes = handle.mediaBytes(assetId)
  } catch {
    return null
  }
  if (bytes.byteLength === 0) return null
  try {
    return await createImageBitmap(new Blob([bytes.slice(0)]))
  } catch {
    return null
  }
}

function createImageResolver(handle: PresentationHandle): CanvasImageResolver {
  const cache = new Map<string, Promise<CanvasImageSource | null>>()
  return (assetId: string) => {
    const cached = cache.get(assetId)
    if (cached) return cached
    const pending = decodeMedia(handle, assetId)
    cache.set(assetId, pending)
    return pending
  }
}

async function renderSlide(
  handle: PresentationHandle,
  faces: Map<number, RegisteredFace>,
  canvas: HTMLCanvasElement,
  context: CanvasRenderingContext2D,
  resolveImage: CanvasImageResolver,
  index: number,
  pageCount: number,
): Promise<PageImage> {
  const list = handle.layoutSlide(index)
  const { width, height } = list
  if (!Number.isFinite(width) || !Number.isFinite(height) || width <= 0 || height <= 0) {
    throw new Error(`Slide ${index + 1} has no usable size`)
  }
  alignRunFaces(list, faces)

  const pageScale = Math.min(1, MAX_PAGE_EDGE / width, MAX_PAGE_EDGE / height)
  const pageWidth = width * pageScale
  const pageHeight = height * pageScale
  const scale = rasterScale(pageWidth, pageHeight, pageCount)
  canvas.width = Math.max(1, Math.round(pageWidth * scale))
  canvas.height = Math.max(1, Math.round(pageHeight * scale))
  await paintSlide(context, list, pageScale * scale, 1, { resolveImage })
  context.save()
  context.setTransform(1, 0, 0, 1, 0, 0)
  // A deck without a resolved background paints onto transparency, which a
  // PDF viewer is free to render as anything; paper is white.
  context.globalCompositeOperation = 'destination-over'
  context.fillStyle = '#ffffff'
  context.fillRect(0, 0, canvas.width, canvas.height)
  context.restore()

  return { dataUrl: canvas.toDataURL('image/png'), width: pageWidth, height: pageHeight }
}

export async function renderPptxPages(data: Uint8Array): Promise<PageSet> {
  await initWasm()
  const handle = openPresentation(data)
  try {
    const snapshot = handle.snapshot()
    const total = snapshot.slides.length
    if (total === 0) throw new Error('The presentation has no slides')
    const faces = registerFonts(handle, await loadFontFaces(collectFontRequests(snapshot)))
    if (faces.size === 0) throw new Error('No font could be loaded for the presentation')

    const canvas = document.createElement('canvas')
    const context = canvas.getContext('2d')
    if (!context) throw new Error('Canvas 2D context unavailable')
    const resolveImage = createImageResolver(handle)

    const count = Math.min(total, PRINT_PAGE_CAP)
    const pages: PageImage[] = []
    let skipped = 0
    for (let index = 0; index < count; index += 1) {
      try {
        pages.push(
          await renderSlide(handle, faces, canvas, context, resolveImage, index, count),
        )
      } catch (error) {
        skipped += 1
        console.warn(`Skipped slide ${index + 1} of ${total}`, error)
      }
    }
    if (pages.length === 0) throw new Error('No slide could be rendered')

    return { pages, padding: 0, truncated: total > PRINT_PAGE_CAP || skipped > 0 }
  } finally {
    handle.dispose()
  }
}
