import {
  loadBundledFontBytes,
  registerBundledFontFace,
  resolveLastResortFace,
  resolveMetricCompatFace,
} from '@betteroffice/docx-fonts'
import {
  initWasm,
  openPresentation,
  paintSlide,
  sizeCanvasForSlide,
  type CanvasImageResolver,
  type DeckSnapshot,
  type PresentationHandle,
  type ShapeSnapshot,
} from '@betteroffice/pptx'
import { PRINT_PAGE_CAP } from '../../shared/ipc.js'
import type { PageImage, PageSet } from './types.js'

const RASTER_SCALE = 2

// The pptx renderer resolves a run's family against its registered faces and
// falls back to the first one registered, so Arial leads and the families a
// deck is most likely to name follow, each as its own metric-compatible face.
const BASE_FAMILIES = ['Arial', 'Calibri', 'Cambria', 'Times New Roman', 'Courier New']
const FONT_STYLES = [
  { bold: false, italic: false },
  { bold: true, italic: false },
  { bold: false, italic: true },
  { bold: true, italic: true },
]

interface FontRequest {
  family: string
  bold: boolean
  italic: boolean
}

function addFontRequest(
  requests: Map<string, FontRequest>,
  family: string,
  bold: boolean,
  italic: boolean,
): void {
  const trimmed = family.trim()
  if (trimmed === '') return
  const key = `${trimmed.toLowerCase()}|${bold ? 1 : 0}|${italic ? 1 : 0}`
  if (!requests.has(key)) requests.set(key, { family: trimmed, bold, italic })
}

function collectShapeFonts(shape: ShapeSnapshot, requests: Map<string, FontRequest>): void {
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

function collectFontRequests(snapshot: DeckSnapshot): FontRequest[] {
  const requests = new Map<string, FontRequest>()
  for (const family of BASE_FAMILIES) {
    for (const style of FONT_STYLES) addFontRequest(requests, family, style.bold, style.italic)
  }
  for (const slide of snapshot.slides) {
    for (const shape of slide.shapes) collectShapeFonts(shape, requests)
  }
  return [...requests.values()]
}

// Layout picks glyph advances from the faces registered with the engine while
// the canvas paints with `ctx.font`, so every face is installed on both sides
// under the family the deck asked for or the two disagree.
async function registerFonts(handle: PresentationHandle, requests: FontRequest[]): Promise<void> {
  const loaded = await Promise.all(
    requests.map(async (request) => {
      const face =
        resolveMetricCompatFace(request.family, request.bold, request.italic) ??
        resolveLastResortFace(request.family, request.bold, request.italic)
      const bytes = await loadBundledFontBytes(face)
      await registerBundledFontFace(face, request.family)
      return { request, bytes: new Uint8Array(bytes.slice(0)) }
    }),
  )
  for (const { request, bytes } of loaded) {
    handle.registerFont({
      family: request.family,
      bold: request.bold,
      italic: request.italic,
      bytes,
    })
  }
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

export async function renderPptxPages(data: Uint8Array): Promise<PageSet> {
  await initWasm()
  const handle = openPresentation(data)
  try {
    const snapshot = handle.snapshot()
    const total = snapshot.slides.length
    if (total === 0) throw new Error('The presentation has no slides')
    await registerFonts(handle, collectFontRequests(snapshot))

    const canvas = document.createElement('canvas')
    const context = canvas.getContext('2d')
    if (!context) throw new Error('Canvas 2D context unavailable')
    const resolveImage = createImageResolver(handle)

    const count = Math.min(total, PRINT_PAGE_CAP)
    const pages: PageImage[] = []
    for (let index = 0; index < count; index += 1) {
      const list = handle.layoutSlide(index)
      sizeCanvasForSlide(canvas, list, RASTER_SCALE)
      await paintSlide(context, list, RASTER_SCALE, 1, { resolveImage })
      context.save()
      context.setTransform(1, 0, 0, 1, 0, 0)
      // A deck without a resolved background paints onto transparency, which a
      // PDF viewer is free to render as anything; paper is white.
      context.globalCompositeOperation = 'destination-over'
      context.fillStyle = '#ffffff'
      context.fillRect(0, 0, canvas.width, canvas.height)
      context.restore()
      pages.push({
        dataUrl: canvas.toDataURL('image/png'),
        width: list.width,
        height: list.height,
      })
    }

    return { pages, padding: 0, truncated: total > PRINT_PAGE_CAP }
  } finally {
    handle.dispose()
  }
}
