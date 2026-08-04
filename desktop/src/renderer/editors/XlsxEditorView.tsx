import { forwardRef, useImperativeHandle, useRef, useState } from 'react'
import { XlsxEditor, type XlsxEditorApi } from '@betteroffice/xlsx-react'
import { en as xlsxEn, locales as xlsxLocales, type PartialLocaleStrings } from '@betteroffice/xlsx-i18n'
import type { OpenedDocument } from '../../shared/ipc.js'
import { useI18n } from '../i18n.js'

const editorLocales = xlsxLocales as Record<string, PartialLocaleStrings>

export interface XlsxEditorViewRef {
  save: () => Uint8Array | null
  undo: () => void
  redo: () => void
  cut: () => Promise<void>
  copy: () => Promise<void>
  paste: () => Promise<void>
  freezeTopRow: () => void
  freezeFirstColumn: () => void
  unfreeze: () => void
}

interface XlsxEditorViewProps {
  document: OpenedDocument
  onChange?: () => void
  onSaveRequest?: (bytes: Uint8Array) => void
}

export const XlsxEditorView = forwardRef<XlsxEditorViewRef, XlsxEditorViewProps>(
  function XlsxEditorView({ document, onChange, onSaveRequest }, ref) {
    const { locale, t } = useI18n()
    const [error] = useState<string | null>(null)
    const apiRef = useRef<XlsxEditorApi | null>(null)
    const onChangeRef = useRef(onChange)
    onChangeRef.current = onChange
    const onSaveRequestRef = useRef(onSaveRequest)
    onSaveRequestRef.current = onSaveRequest

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
      undo: () => apiRef.current?.undo(),
      redo: () => apiRef.current?.redo(),
      cut: () => apiRef.current?.cutSelection() ?? Promise.resolve(),
      copy: () => apiRef.current?.copySelection() ?? Promise.resolve(),
      paste: () => apiRef.current?.pasteSelection() ?? Promise.resolve(),
      freezeTopRow: () => setFreezePane(1, null),
      freezeFirstColumn: () => setFreezePane(null, 1),
      unfreeze: () => setFreezePane(null, null),
    }))

    if (error) {
      return (
        <div className="flex h-full items-center justify-center p-8">
          <div className="max-w-md text-center">
            <h2 className="text-sm font-semibold text-red-700">
              {t('editor.couldNotOpen', { name: document.name })}
            </h2>
            <p className="mt-2 text-sm text-neutral-600">{error}</p>
          </div>
        </div>
      )
    }

    return (
      <XlsxEditor
        file={document.data}
        fileName={document.name}
        i18n={editorLocales[locale] ?? xlsxEn}
        onSave={(bytes: Uint8Array) => onSaveRequestRef.current?.(bytes)}
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
