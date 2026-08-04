import { forwardRef, useCallback, useEffect, useImperativeHandle, useRef, useState } from 'react'
import { PptxEditor, type PptxEditorApi } from '@betteroffice/pptx-react'
import type { DeckSnapshot, PptxFontFace, ShapeSnapshot } from '@betteroffice/pptx'
import {
  loadBundledFontBytes,
  resolveLastResortFace,
  resolveMetricCompatFace,
} from '@betteroffice/docx-fonts'
import { en as pptxEn, locales as pptxLocales, type PartialLocaleStrings } from '@betteroffice/pptx-i18n'
import { countCharacters, countWords, type EditorStats } from '../services/textStats.js'
import { useI18n } from '../i18n.js'

const editorLocales = pptxLocales as Record<string, PartialLocaleStrings>

// The pptx renderer resolves a run's family against its registered faces and
// falls back to the first one registered, so Arial leads and the families a
// deck is most likely to name follow, each as its own metric-compatible face.
const FONT_FAMILIES = ['Arial', 'Calibri', 'Cambria', 'Times New Roman', 'Courier New'] as const
const FONT_STYLES = [
  { bold: false, italic: false },
  { bold: true, italic: false },
  { bold: false, italic: true },
  { bold: true, italic: true },
] as const

let bundledFonts: Promise<PptxFontFace[]> | null = null

function presentationFonts(): Promise<PptxFontFace[]> {
  bundledFonts ??= Promise.all(
    FONT_FAMILIES.flatMap((family) =>
      FONT_STYLES.map(async ({ bold, italic }) => {
        const face =
          resolveMetricCompatFace(family, bold, italic) ??
          resolveLastResortFace(family, bold, italic)
        const bytes = await loadBundledFontBytes(face)
        return { family, bold, italic, bytes: new Uint8Array(bytes.slice(0)) }
      }),
    ),
  )
  return bundledFonts
}

// Everything a deck's text lives in already crosses the wasm boundary inside
// the snapshot the editor lays out from, so word count and spell check read it
// instead of asking the engine again. Shapes are visited in the order the deck
// stores them, groups depth-first, and every story, paragraph and shape break
// becomes a newline so counting never fuses the last word of one with the
// first of the next.
function shapeText(shape: ShapeSnapshot, out: string[]): void {
  for (const story of shape.textStories) {
    for (const paragraph of story.paragraphs) {
      out.push(paragraph.runs.map((run) => run.text).join(''))
    }
  }
  for (const child of shape.children) shapeText(child, out)
}

function deckText(snapshot: DeckSnapshot): string {
  return snapshot.slides
    .map((slide) => {
      const lines: string[] = []
      for (const shape of slide.shapes) shapeText(shape, lines)
      return lines.join('\n')
    })
    .join('\n\n')
}

const TEXT_REFRESH_DELAY_MS = 500

export interface PptxEditorViewRef {
  getText: () => string
  getStats: () => EditorStats
  undo: () => void
  redo: () => void
  cut: () => Promise<void>
  copy: () => Promise<void>
  paste: () => Promise<void>
  deleteSelection: () => void
  selectAll: () => void
  getZoom: () => number
  setZoom: (zoom: number) => void
}

interface EditorDocument {
  name: string
  // PptxEditor disposes the presentation and reopens whenever this buffer's
  // identity changes, so it must stay the bytes captured at open — never bytes
  // a later save serialized, which would discard edits made since.
  seed: Uint8Array
}

interface PptxEditorViewProps {
  document: EditorDocument
  onChange?: () => void
}

export const PptxEditorView = forwardRef<PptxEditorViewRef, PptxEditorViewProps>(
  function PptxEditorView({ document, onChange }, ref) {
    const { locale, t } = useI18n()
    const [fonts, setFonts] = useState<PptxFontFace[] | null>(null)
    const [error, setError] = useState<string | null>(null)
    const apiRef = useRef<PptxEditorApi | null>(null)
    const textRef = useRef('')
    const wordsRef = useRef(0)
    const slidesRef = useRef(1)
    const refreshTimer = useRef<ReturnType<typeof setTimeout> | null>(null)
    const onChangeRef = useRef(onChange)
    onChangeRef.current = onChange

    // Walking every story is cheap next to laying a slide out, but an edit
    // arrives per keystroke, so the walk waits for a pause the way docx does.
    const applySnapshot = useCallback((snapshot: DeckSnapshot) => {
      textRef.current = deckText(snapshot)
      wordsRef.current = countWords(textRef.current)
      slidesRef.current = Math.max(1, snapshot.slides.length)
    }, [])

    useEffect(() => {
      textRef.current = ''
      wordsRef.current = 0
      slidesRef.current = 1
      return () => {
        if (refreshTimer.current) clearTimeout(refreshTimer.current)
      }
    }, [document.seed])

    useEffect(() => {
      let canceled = false
      presentationFonts().then(
        (loaded) => { if (!canceled) setFonts(loaded) },
        (err: unknown) => {
          if (!canceled) setError(err instanceof Error ? err.message : String(err))
        },
      )
      return () => { canceled = true }
    }, [])

    useImperativeHandle(ref, () => ({
      getText: () => textRef.current,
      getStats: () => ({
        words: wordsRef.current,
        characters: countCharacters(textRef.current),
        page: (apiRef.current?.getSlideIndex() ?? 0) + 1,
        pages: slidesRef.current,
      }),
      undo: () => apiRef.current?.undo(),
      redo: () => apiRef.current?.redo(),
      cut: () => apiRef.current?.cutSelection() ?? Promise.resolve(),
      copy: () => apiRef.current?.copySelection() ?? Promise.resolve(),
      paste: () => apiRef.current?.pasteSelection() ?? Promise.resolve(),
      deleteSelection: () => apiRef.current?.deleteSelection(),
      selectAll: () => apiRef.current?.selectAll(),
      getZoom: () => apiRef.current?.getZoom() ?? 1,
      setZoom: (zoom: number) => apiRef.current?.setZoom(zoom),
    }))

    if (error) {
      return (
        <div className="flex h-full w-full items-center justify-center p-8">
          <div className="max-w-md text-center">
            <h2 className="text-sm font-semibold text-red-700">
              {t('editor.couldNotOpen', { name: document.name })}
            </h2>
            <p className="mt-2 text-sm text-neutral-600">{error}</p>
          </div>
        </div>
      )
    }

    if (!fonts) {
      return (
        <div className="flex h-full w-full items-center justify-center">
          <span className="text-sm text-neutral-500">
            {t('editor.loading', { name: document.name })}
          </span>
        </div>
      )
    }

    return (
      <PptxEditor
        file={document.seed}
        fonts={fonts}
        i18n={editorLocales[locale] ?? pptxEn}
        onReady={(api: PptxEditorApi) => {
          apiRef.current = api
          applySnapshot(api.handle.snapshot())
        }}
        onChange={(snapshot: DeckSnapshot) => {
          onChangeRef.current?.()
          if (refreshTimer.current) clearTimeout(refreshTimer.current)
          refreshTimer.current = setTimeout(() => applySnapshot(snapshot), TEXT_REFRESH_DELAY_MS)
        }}
        className="h-full w-full"
      />
    )
  },
)
