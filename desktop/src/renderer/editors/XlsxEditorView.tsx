import { forwardRef, useImperativeHandle, useRef, useState } from 'react'
import { XlsxEditor, type XlsxEditorApi } from '@betteroffice/xlsx-react'
import { en as xlsxEn, locales as xlsxLocales, type PartialLocaleStrings } from '@betteroffice/xlsx-i18n'
import {
  applyCreateChart,
  applyWriteCells,
  executeXlsxAgentTool,
  validateCreateChart,
  validateWriteCells,
  type CreateChartProposal,
  type WriteCellsProposal,
} from '../services/agentTools.js'
import { useI18n } from '../i18n.js'

const editorLocales = xlsxLocales as Record<string, PartialLocaleStrings>

export interface XlsxEditorViewRef {
  save: () => Uint8Array | null
  undo: () => void
  redo: () => void
  cut: () => Promise<void>
  copy: () => Promise<void>
  paste: () => Promise<void>
  deleteSelection: () => void
  selectAll: () => void
  getZoom: () => number
  setZoom: (zoom: number) => void
  freezeTopRow: () => void
  freezeFirstColumn: () => void
  unfreeze: () => void
  /** Execute one read-only AI-assistant tool against the live workbook. */
  agentTool: (name: string, args: Record<string, unknown>) => unknown
  /** Validate a write_cells request into a reviewable proposal (no mutation). */
  agentValidateWrite: (
    args: Record<string, unknown>
  ) => { proposal: WriteCellsProposal } | { error: string }
  /** Apply a user-approved proposal as one undo step. */
  agentApplyWrite: (proposal: WriteCellsProposal) => unknown
  /** Validate a create_chart request into a reviewable proposal (no mutation). */
  agentValidateChart: (
    args: Record<string, unknown>
  ) => { proposal: CreateChartProposal } | { error: string }
  /** Author a user-approved chart. */
  agentApplyChart: (proposal: CreateChartProposal) => unknown
}

interface EditorDocument {
  name: string
  // XlsxEditor disposes the workbook and reopens whenever this buffer's
  // identity changes, so it must stay the bytes captured at open — never bytes
  // a later save serialized, which would discard edits made since.
  seed: Uint8Array
}

interface XlsxEditorViewProps {
  document: EditorDocument
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
      copy: async () => {
        await apiRef.current?.copySelection()
      },
      paste: () => apiRef.current?.pasteSelection() ?? Promise.resolve(),
      deleteSelection: () => apiRef.current?.clearSelection(),
      selectAll: () => apiRef.current?.selectAll(),
      getZoom: () => apiRef.current?.getZoom() ?? 1,
      setZoom: (zoom: number) => apiRef.current?.setZoom(zoom),
      freezeTopRow: () => setFreezePane(1, null),
      freezeFirstColumn: () => setFreezePane(null, 1),
      unfreeze: () => setFreezePane(null, null),
      agentTool: (name: string, args: Record<string, unknown>) => {
        const handle = apiRef.current?.handle
        if (!handle) return { error: 'no workbook is open' }
        return executeXlsxAgentTool(handle, name, args)
      },
      agentValidateWrite: (args: Record<string, unknown>) => {
        const handle = apiRef.current?.handle
        if (!handle) return { error: 'no workbook is open' }
        return validateWriteCells(handle, args)
      },
      agentApplyWrite: (proposal: WriteCellsProposal) => {
        const handle = apiRef.current?.handle
        if (!handle) return { error: 'no workbook is open' }
        const result = applyWriteCells(handle, proposal)
        onChangeRef.current?.()
        return result
      },
      agentValidateChart: (args: Record<string, unknown>) => {
        const handle = apiRef.current?.handle
        if (!handle) return { error: 'no workbook is open' }
        return validateCreateChart(handle, args)
      },
      agentApplyChart: (proposal: CreateChartProposal) => {
        const api = apiRef.current
        if (!api) return { error: 'no workbook is open' }
        try {
          const result = applyCreateChart(api.handle, proposal)
          api.refreshProposals()
          onChangeRef.current?.()
          return result
        } catch (error) {
          return { error: error instanceof Error ? error.message : String(error) }
        }
      },
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
        file={document.seed}
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
