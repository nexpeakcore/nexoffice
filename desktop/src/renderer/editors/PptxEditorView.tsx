import { forwardRef, useEffect, useImperativeHandle, useRef, useState } from 'react'
import { PptxEditor, type PptxEditorApi } from '@betteroffice/pptx-react'
import type { PptxFontFace } from '@betteroffice/pptx'
import {
  loadBundledFontBytes,
  resolveLastResortFace,
  resolveMetricCompatFace,
} from '@betteroffice/docx-fonts'
import { en as pptxEn, locales as pptxLocales, type PartialLocaleStrings } from '@betteroffice/pptx-i18n'
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

export interface PptxEditorViewRef {
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
}

export const PptxEditorView = forwardRef<PptxEditorViewRef, PptxEditorViewProps>(
  function PptxEditorView({ document }, ref) {
    const { locale, t } = useI18n()
    const [fonts, setFonts] = useState<PptxFontFace[] | null>(null)
    const [error, setError] = useState<string | null>(null)
    const apiRef = useRef<PptxEditorApi | null>(null)

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
        }}
        className="h-full w-full"
      />
    )
  },
)
