import { useCallback, useEffect, useRef, useState } from 'react'
import { DocxEditor, type DocxEditorRef } from '@betteroffice/docx-react'
import '@betteroffice/docx-react/styles.css'
import { PRINT_PAGE_CAP, type PrintJob } from '../../shared/ipc.js'
import { renderPptxPages } from './renderPptxPages.js'
import { renderXlsxPages } from './renderXlsxPages.js'
import type { PageImage, PageSet } from './types.js'

const POLL_INTERVAL_MS = 300
const STABLE_POLLS = 6
const CAPTURE_TIMEOUT_MS = 90_000

type Phase =
  | { mode: 'idle' }
  | { mode: 'docx'; data: Uint8Array }
  | { mode: 'pages'; set: PageSet }

function reportFailure(error: unknown): void {
  const message = error instanceof Error ? error.message : String(error)
  window.nexoffice.printRendered({ ok: false, error: message })
}

function DocxCapture({
  data,
  onCaptured,
  onError,
}: {
  data: Uint8Array
  onCaptured: (set: PageSet) => void
  onError: (error: unknown) => void
}) {
  const editorRef = useRef<DocxEditorRef>(null)
  const hostRef = useRef<HTMLDivElement | null>(null)
  const doneRef = useRef(false)

  useEffect(() => {
    const started = Date.now()
    let stablePolls = 0
    let lastSignature = ''

    const finish = (canvases: HTMLCanvasElement[]): void => {
      doneRef.current = true
      clearInterval(timer)
      try {
        const truncated = canvases.length > PRINT_PAGE_CAP
        const pages: PageImage[] = canvases.slice(0, PRINT_PAGE_CAP).map((canvas) => ({
          dataUrl: canvas.toDataURL('image/png'),
          width: parseFloat(canvas.style.width) || canvas.width,
          height: parseFloat(canvas.style.height) || canvas.height,
        }))
        if (pages.length === 0) throw new Error('No document pages were rendered')
        onCaptured({ pages, padding: 0, truncated, skippedPages: [] })
      } catch (error) {
        onError(error)
      }
    }

    const timer = setInterval(() => {
      if (doneRef.current) return
      if (Date.now() - started > CAPTURE_TIMEOUT_MS) {
        doneRef.current = true
        clearInterval(timer)
        onError(new Error('Timed out preparing document pages'))
        return
      }
      const host = hostRef.current
      const total = editorRef.current?.getTotalPages() ?? 0
      if (!host || total < 1) return
      const canvases = Array.from(host.querySelectorAll('canvas')).filter(
        (canvas) => canvas.width > 0 && canvas.height > 0 && canvas.style.width !== '',
      )
      if (canvases.length < total) return
      const signature = canvases
        .map((canvas) => `${canvas.width}x${canvas.height}`)
        .join(',')
      if (signature !== lastSignature) {
        lastSignature = signature
        stablePolls = 0
        return
      }
      stablePolls += 1
      if (stablePolls >= STABLE_POLLS) finish(canvases)
    }, POLL_INTERVAL_MS)

    return () => clearInterval(timer)
  }, [data, onCaptured, onError])

  return (
    <div ref={hostRef} style={{ position: 'fixed', inset: 0, overflow: 'auto', background: '#fff' }}>
      <DocxEditor
        ref={editorRef}
        documentBuffer={data}
        readOnly
        showToolbar={false}
        showOutlineButton={false}
        style={{ width: '100%', height: '100%' }}
        onError={(error) => {
          if (doneRef.current) return
          doneRef.current = true
          onError(error)
        }}
      />
    </div>
  )
}

interface PageBlock {
  page: PageImage
  boxWidth: number
  boxHeight: number
  pageName: string
}

function buildPageBlocks(set: PageSet): { blocks: PageBlock[]; pageRules: string } {
  const namesBySize = new Map<string, string>()
  const blocks = set.pages.map((page) => {
    const boxWidth = Math.round(page.width + set.padding * 2)
    const boxHeight = Math.round(page.height + set.padding * 2)
    const sizeKey = `${boxWidth}x${boxHeight}`
    let pageName = namesBySize.get(sizeKey)
    if (!pageName) {
      pageName = `sheet-${namesBySize.size}`
      namesBySize.set(sizeKey, pageName)
    }
    return { page, boxWidth, boxHeight, pageName }
  })
  const pageRules = Array.from(namesBySize.entries())
    .map(([sizeKey, pageName]) => {
      const [width, height] = sizeKey.split('x')
      return `@page ${pageName} { size: ${width}px ${height}px; margin: 0; }`
    })
    .join('\n')
  return { blocks, pageRules }
}

function PrintPages({ set }: { set: PageSet }) {
  const rootRef = useRef<HTMLDivElement | null>(null)
  const { blocks, pageRules } = buildPageBlocks(set)

  useEffect(() => {
    let cancelled = false
    const images = Array.from(rootRef.current?.querySelectorAll('img') ?? [])
    void Promise.all(images.map((image) => image.decode().catch(() => undefined))).then(() => {
      if (cancelled) return
      requestAnimationFrame(() =>
        requestAnimationFrame(() => {
          if (cancelled) return
          window.nexoffice.printRendered({
            ok: true,
            pages: set.pages.length,
            truncated: set.truncated,
            skippedPages: set.skippedPages,
          })
        }),
      )
    })
    return () => {
      cancelled = true
    }
  }, [set])

  return (
    <div ref={rootRef}>
      <style>{`
        ${pageRules}
        html, body { margin: 0; padding: 0; background: #ffffff; }
      `}</style>
      {blocks.map((block, index) => (
        <div
          key={index}
          style={{
            width: block.boxWidth,
            height: block.boxHeight,
            padding: set.padding,
            overflow: 'hidden',
            boxSizing: 'border-box',
            background: '#ffffff',
            page: block.pageName,
            breakAfter: index < blocks.length - 1 ? 'page' : 'auto',
          }}
        >
          <img
            src={block.page.dataUrl}
            alt=""
            style={{ width: block.page.width, height: block.page.height, display: 'block' }}
          />
        </div>
      ))}
    </div>
  )
}

export function PrintView() {
  const [phase, setPhase] = useState<Phase>({ mode: 'idle' })
  const startedRef = useRef(false)

  const handleCaptured = useCallback((set: PageSet) => {
    setPhase({ mode: 'pages', set })
  }, [])

  useEffect(() => {
    const unsubscribe = window.nexoffice.onPrintRender((job: PrintJob) => {
      if (startedRef.current) return
      startedRef.current = true
      if (job.kind === 'docx') {
        setPhase({ mode: 'docx', data: job.data })
      } else if (job.kind === 'xlsx') {
        void renderXlsxPages(job.data).then(handleCaptured).catch(reportFailure)
      } else if (job.kind === 'pptx') {
        void renderPptxPages(job.data).then(handleCaptured).catch(reportFailure)
      } else {
        window.nexoffice.printRendered({
          ok: false,
          error: `PDF export is not supported for ${String(job.kind)} files yet`,
        })
      }
    })
    window.nexoffice.printReady()
    return unsubscribe
  }, [handleCaptured])

  if (phase.mode === 'docx') {
    return <DocxCapture data={phase.data} onCaptured={handleCaptured} onError={reportFailure} />
  }
  if (phase.mode === 'pages') {
    return <PrintPages set={phase.set} />
  }
  return null
}
