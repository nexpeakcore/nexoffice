import { forwardRef, useCallback, useEffect, useImperativeHandle, useMemo, useRef, useState } from 'react'
import { DocxEditor, type BundledFontProvider, type DocxEditorRef } from '@betteroffice/docx-react'
import '@betteroffice/docx-react/styles.css'
import {
  loadBundledFontBytes,
  resolveLastResortFace,
  resolveMetricCompatFace,
  resolveScriptFallbackFace,
} from '@betteroffice/docx-fonts'
import type { OpenedDocument } from '../../shared/ipc.js'

export interface DocxEditorViewRef {
  save: () => Promise<ArrayBuffer | null>
  getText: () => string
  getStats: () => { words: number; characters: number; page: number; pages: number }
}

interface DocxEditorViewProps {
  document: OpenedDocument
  onChange?: () => void
}

function extractText(blocks: unknown[]): string {
  let out = ''
  for (const block of blocks as Array<Record<string, unknown>>) {
    if (!block || typeof block !== 'object') continue
    if (block.type === 'text') {
      out += (block.text as string) ?? ''
    } else if (block.type === 'table') {
      for (const row of (block.rows as unknown[]) ?? []) {
        for (const cell of ((row as Record<string, unknown>).cells as unknown[]) ?? []) {
          out += extractText(((cell as Record<string, unknown>).content as unknown[]) ?? []) + '\t'
        }
        out += '\n'
      }
    } else if (Array.isArray(block.content)) {
      out += extractText(block.content as unknown[])
      if (block.type === 'paragraph') out += '\n'
    }
  }
  return out
}

function countWords(text: string): number {
  const matches = text.match(/[\p{L}\p{N}]+(?:[''-][\p{L}\p{N}]+)*/gu)
  return matches?.length ?? 0
}

const TEXT_REFRESH_DELAY_MS = 500

export const DocxEditorView = forwardRef<DocxEditorViewRef, DocxEditorViewProps>(
  function DocxEditorView({ document, onChange }, ref) {
    const editorRef = useRef<DocxEditorRef>(null)
    const measurementFontProvider = useMemo<BundledFontProvider>(
      () => ({
        resolve(family, bold, italic) {
          const face = resolveMetricCompatFace(family, bold, italic)
          return face ? () => loadBundledFontBytes(face) : undefined
        },
        resolveScriptFallback(script, bold, italic) {
          const face = resolveScriptFallbackFace(script, bold, italic)
          return face ? () => loadBundledFontBytes(face) : undefined
        },
        resolveLastResort(family, bold, italic) {
          return () => loadBundledFontBytes(resolveLastResortFace(family, bold, italic))
        },
      }),
      [],
    )
    const [error, setError] = useState<string | null>(null)
    const textRef = useRef('')
    const statsRef = useRef({ words: 0, characters: 0, page: 1, pages: 1 })
    const onChangeRef = useRef(onChange)
    onChangeRef.current = onChange

    // Projecting the full document out of the CRDT is expensive, so word count
    // and spell check text refresh only after typing pauses — never per edit.
    const refreshText = useCallback(() => {
      const doc = editorRef.current?.getDocument()
      if (!doc) return
      const text = extractText((doc.package?.document?.content as unknown[]) ?? [])
      textRef.current = text
      statsRef.current = { ...statsRef.current, words: countWords(text) }
    }, [])

    useEffect(() => {
      setError(null)
      textRef.current = ''
      statsRef.current = { words: 0, characters: 0, page: 1, pages: 1 }

      let unsubscribe: (() => void) | null = null
      let refreshTimer: ReturnType<typeof setTimeout> | null = null
      const sessionPoll = setInterval(() => {
        const session = editorRef.current?.getEditorRef()?.getYrsSession()
        if (!session) return
        clearInterval(sessionPoll)
        refreshText()
        unsubscribe = session.onUpdate((_update, origin) => {
          if (origin !== 'local') return
          onChangeRef.current?.()
          if (refreshTimer) clearTimeout(refreshTimer)
          refreshTimer = setTimeout(refreshText, TEXT_REFRESH_DELAY_MS)
        })
      }, 300)

      return () => {
        clearInterval(sessionPoll)
        if (refreshTimer) clearTimeout(refreshTimer)
        unsubscribe?.()
      }
    }, [document.path, document.data, refreshText])

    useImperativeHandle(ref, () => ({
      save: async () => (await editorRef.current?.save()) ?? null,
      getText: () => textRef.current,
      getStats: () => {
        const editor = editorRef.current
        return {
          words: statsRef.current.words,
          characters: textRef.current.replace(/\s/g, '').length,
          page: editor?.getCurrentPage() ?? 1,
          pages: editor?.getTotalPages() ?? 1,
        }
      },
    }))

    if (error) {
      return (
        <div className="flex h-full items-center justify-center p-8">
          <div className="max-w-md text-center">
            <h2 className="text-sm font-semibold text-red-700">Could not open {document.name}</h2>
            <p className="mt-2 text-sm text-neutral-600">{error}</p>
          </div>
        </div>
      )
    }

    return (
      <DocxEditor
        ref={editorRef}
        documentBuffer={document.data}
        measurementFontProvider={measurementFontProvider}
        showFileOpen={false}
        className="h-full w-full"
        loadingIndicator={
          <div className="flex h-full items-center justify-center">
            <span className="text-sm text-neutral-500">Loading {document.name}…</span>
          </div>
        }
        onError={(err) => setError(err.message)}
      />
    )
  },
)
