import { initWasm, openWorkbook, paintDisplayList } from '@betteroffice/xlsx'
import { PRINT_PAGE_CAP } from '../../shared/ipc.js'
import type { PageImage, PageSet } from './types.js'

const PAGE_WIDTH = 794
const PAGE_HEIGHT = 1123
const PAGE_PADDING = 38
const RASTER_SCALE = 2

function frozenExtent(
  indices: number[] | undefined,
  offsets: number[] | undefined,
  frozenCount: number,
): number {
  if (frozenCount <= 0 || !indices || !offsets) return 0
  let visibleFrozen = 0
  for (const index of indices) {
    if (index < frozenCount) visibleFrozen += 1
  }
  return offsets[visibleFrozen] ?? 0
}

export async function renderXlsxPages(data: Uint8Array): Promise<PageSet> {
  await initWasm()
  const handle = openWorkbook(data)
  try {
    const info = handle.sheetInfo()
    const innerWidth = PAGE_WIDTH - PAGE_PADDING * 2
    const innerHeight = PAGE_HEIGHT - PAGE_PADDING * 2

    const probe = handle.displayList({ x: 0, y: 0, width: innerWidth, height: innerHeight })
    const frozenWidth = frozenExtent(probe.grid?.colIndices, probe.grid?.colOffsets, info.frozenCols)
    const frozenHeight = frozenExtent(probe.grid?.rowIndices, probe.grid?.rowOffsets, info.frozenRows)

    const stepX = Math.max(innerWidth - frozenWidth, innerWidth / 2)
    const stepY = Math.max(innerHeight - frozenHeight, innerHeight / 2)
    const bodyWidth = Math.max(info.contentWidth - frozenWidth, 1)
    const bodyHeight = Math.max(info.contentHeight - frozenHeight, 1)
    const columns = Math.max(1, Math.ceil(bodyWidth / stepX))
    const rows = Math.max(1, Math.ceil(bodyHeight / stepY))
    const truncated = columns * rows > PRINT_PAGE_CAP

    const canvas = document.createElement('canvas')
    canvas.width = innerWidth * RASTER_SCALE
    canvas.height = innerHeight * RASTER_SCALE
    const context = canvas.getContext('2d')
    if (!context) throw new Error('Canvas 2D context unavailable')

    const pages: PageImage[] = []
    outer: for (let row = 0; row < rows; row += 1) {
      for (let column = 0; column < columns; column += 1) {
        if (pages.length >= PRINT_PAGE_CAP) break outer
        const list = handle.displayList({
          x: column * stepX,
          y: row * stepY,
          width: innerWidth,
          height: innerHeight,
        })
        paintDisplayList(context, list, RASTER_SCALE)
        pages.push({
          dataUrl: canvas.toDataURL('image/png'),
          width: innerWidth,
          height: innerHeight,
        })
      }
    }

    return { pages, pageWidth: PAGE_WIDTH, pageHeight: PAGE_HEIGHT, padding: PAGE_PADDING, truncated }
  } finally {
    handle.dispose()
  }
}
