import { forwardRef, useImperativeHandle, useRef, useState } from 'react'
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
    const onChangeRef = useRef(onChange)
    onChangeRef.current = onChange

    const setFreezePane = (row: number | null, col: number | null) => {
      const handle = apiRef.current?.handle
      if (!handle) return
      handle.setFreezePane(handle.sheetInfo().activeSheet, row, col)
      onChangeRef.current?.()
    }

    useImperativeHandle(ref, () => ({
      save: () => {
        const handle = apiRef.current?.handle
        return handle ? handle.save() : null
      },
      freezeTopRow: () => setFreezePane(1, null),
      freezeFirstColumn: () => setFreezePane(null, 1),
      unfreeze: () => setFreezePane(null, null),
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
        onEdit={() => onChangeRef.current?.()}
        onReady={(api: XlsxEditorApi) => {
          apiRef.current = api
          return () => {
            if (apiRef.current === api) apiRef.current = null
          }
        }}
        className="h-full w-full"
      />
    )
  },
)
