import { forwardRef, useEffect, useImperativeHandle, useRef, useState } from 'react'
import { DocxEditor, type DocxEditorRef } from '@betteroffice/docx-react'
import '@betteroffice/docx-react/styles.css'
import type { Document as DocxDocument } from '@betteroffice/docx/types/document'
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
    if (block.type === 'paragraph') {
      out += extractText((block.content as unknown[]) ?? []) + '\n'
    } else if (block.type === 'table') {
      for (const row of (block.rows as unknown[]) ?? []) {
        for (const cell of ((row as Record<string, unknown>).cells as unknown[]) ?? []) {
          out += extractText(((cell as Record<string, unknown>).content as unknown[]) ?? []) + '\t'
        }
        out += '\n'
      }
    } else if (block.type === 'text') {
      out += (block.text as string) ?? ''
    }
  }
  return out
}

function countWords(text: string): number {
  const matches = text.match(/[\p{L}\p{N}]+(?:[''-][\p{L}\p{N}]+)*/gu)
  return matches?.length ?? 0
}

export const DocxEditorView = forwardRef<DocxEditorViewRef, DocxEditorViewProps>(
  function DocxEditorView({ document, onChange }, ref) {
    const editorRef = useRef<DocxEditorRef>(null)
    const [error, setError] = useState<string | null>(null)
    const textRef = useRef('')
    const statsRef = useRef({ words: 0, characters: 0, page: 1, pages: 1 })
    const parsed = useRef(false)

    useEffect(() => {
      parsed.current = false
      setError(null)
      textRef.current = ''
      statsRef.current = { words: 0, characters: 0, page: 1, pages: 1 }
    }, [document.path, document.data])

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
        showFileOpen={false}
        className="h-full w-full"
        loadingIndicator={
          <div className="flex h-full items-center justify-center">
            <span className="text-sm text-neutral-500">Loading {document.name}…</span>
          </div>
        }
        onChange={(doc: DocxDocument) => {
          const text = extractText((doc.package?.document?.content as unknown[]) ?? [])
          textRef.current = text
          statsRef.current = { ...statsRef.current, words: countWords(text) }
          if (!parsed.current) {
            parsed.current = true
            return
          }
          onChange?.()
        }}
        onError={(err) => setError(err.message)}
      />
    )
  },
)
