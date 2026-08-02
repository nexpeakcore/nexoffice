import { forwardRef, useEffect, useImperativeHandle, useRef, useState } from 'react'
import { XlsxEditor, type XlsxEditorApi } from '@betteroffice/xlsx-react'
import type { OpenedDocument } from '../../shared/ipc.js'

export interface XlsxEditorViewRef {
  save: () => Uint8Array | null
  freezeTopRow: () => void
  freezeFirstColumn: () => void
  unfreeze: () => void
}

interface XlsxEditorViewProps {
  document: OpenedDocument
  onChange?: () => void
}

export const XlsxEditorView = forwardRef<XlsxEditorViewRef, XlsxEditorViewProps>(
  function XlsxEditorView({ document, onChange }, ref) {
    const [error] = useState<string | null>(null)
    const apiRef = useRef<XlsxEditorApi | null>(null)
    const undoDepthRef = useRef(0)
    const onChangeRef = useRef(onChange)
    onChangeRef.current = onChange

    useEffect(() => {
      const timer = setInterval(() => {
        const handle = apiRef.current?.handle
        if (!handle) return
        try {
          const { undoDepth } = handle.historyState()
          if (undoDepth !== undoDepthRef.current) {
            undoDepthRef.current = undoDepth
            onChangeRef.current?.()
          }
        } catch {
          // handle disposed between ticks
        }
      }, 700)
      return () => clearInterval(timer)
    }, [])

    useImperativeHandle(ref, () => ({
      save: () => {
        const handle = apiRef.current?.handle
        return handle ? handle.save() : null
      },
      freezeTopRow: () => {
        apiRef.current?.handle.setFreezePane(0, 1, null)
        onChange?.()
      },
      freezeFirstColumn: () => {
        apiRef.current?.handle.setFreezePane(0, null, 1)
        onChange?.()
      },
      unfreeze: () => {
        apiRef.current?.handle.setFreezePane(0, null, null)
        onChange?.()
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
      <XlsxEditor
        file={document.data}
        fileName={document.name}
        onSave={() => onChange?.()}
        onReady={(api: XlsxEditorApi) => {
          apiRef.current = api
          undoDepthRef.current = 0
          return () => {
            if (apiRef.current === api) apiRef.current = null
          }
        }}
        className="h-full w-full"
      />
    )
  },
)
