import { useCallback, useEffect, useRef, useState } from 'react'
import type { DocumentKind, MenuAction, OpenedDocument } from '../shared/ipc.js'
import { SpellCheckPanel } from './components/SpellCheckPanel.js'
import { DocxEditorView, type DocxEditorViewRef } from './editors/DocxEditorView.js'
import { XlsxEditorView, type XlsxEditorViewRef } from './editors/XlsxEditorView.js'
import { spellCheckService, type Misspelling } from './services/spellcheck.js'

interface DocumentState extends Omit<OpenedDocument, 'kind'> {
  kind: DocumentKind
  dirty: boolean
}

const KIND_LABEL = {
  docx: 'Document',
  xlsx: 'Workbook',
  pptx: 'Presentation',
} as const

export function App() {
  const [document, setDocument] = useState<DocumentState | null>(null)
  const [status, setStatus] = useState('Ready')
  const docxRef = useRef<DocxEditorViewRef>(null)
  const [spellPanelOpen, setSpellPanelOpen] = useState(false)
  const [spellLoading, setSpellLoading] = useState(true)
  const [spellError, setSpellError] = useState<string | null>(null)
  const [misspellings, setMisspellings] = useState<Misspelling[]>([])
  const [docStats, setDocStats] = useState({ words: 0, characters: 0, page: 1, pages: 1 })
  const [editRevision, setEditRevision] = useState(0)
  const xlsxRef = useRef<XlsxEditorViewRef>(null)
  const documentRef = useRef<DocumentState | null>(null)
  const closingRef = useRef(false)

  useEffect(() => {
    documentRef.current = document
  }, [document])

  const isDocx = document?.kind === 'docx'

  // Live stats ticker — re-reads editor stats into state every 700ms
  useEffect(() => {
    if (!isDocx) return
    const timer = setInterval(() => {
      const stats = docxRef.current?.getStats()
      if (stats) setDocStats(stats)
    }, 700)
    return () => clearInterval(timer)
  }, [isDocx])

  // Init spell check on mount
  useEffect(() => {
    let canceled = false
    spellCheckService
      .init('en')
      .then(() => { if (!canceled) setSpellLoading(false) })
      .catch((err: Error) => { if (!canceled) { setSpellLoading(false); setSpellError(err.message) } })
    return () => { canceled = true }
  }, [])

  // Periodic spell check for open docx docs
  useEffect(() => {
    if (spellLoading || spellError || !isDocx) {
      setMisspellings([])
      return
    }
    const timer = setTimeout(() => {
      const text = docxRef.current?.getText() ?? ''
      setMisspellings(spellCheckService.check(text))
    }, 800)
    return () => clearTimeout(timer)
  }, [spellLoading, spellError, isDocx, document?.path, editRevision, docStats.words])

  const getEditorText = useCallback(() => docxRef.current?.getText() ?? '', [])

  const markEdited = useCallback(() => {
    setDocument((prev) => (prev && !prev.dirty ? { ...prev, dirty: true } : prev))
    setEditRevision((revision) => revision + 1)
  }, [])

  const getCurrentBytes = useCallback(async (current: DocumentState): Promise<Uint8Array> => {
    if (current.kind === 'docx') {
      const buffer = await docxRef.current?.save()
      if (buffer) return new Uint8Array(buffer)
    } else if (current.kind === 'xlsx') {
      const bytes = xlsxRef.current?.save()
      if (bytes) return bytes
    }
    return current.data
  }, [])

  const saveDocument = useCallback(
    async (forceDialog: boolean): Promise<boolean> => {
      const current = documentRef.current
      if (!current) return false
      try {
        const data = await getCurrentBytes(current)
        const save = forceDialog ? window.nexoffice.saveFileAs : window.nexoffice.saveFile
        const result = await save(current.path, current.kind, data)
        if (result.canceled || !result.path) {
          setStatus('Save canceled')
          return false
        }
        const savedPath = result.path
        const savedName = savedPath.split(/[\\/]/).pop() ?? current.name
        setDocument((prev) =>
          prev ? { ...prev, path: savedPath, name: savedName, data, dirty: false } : prev,
        )
        setStatus(`Saved ${savedPath}`)
        return true
      } catch (error) {
        setStatus(`Save failed: ${error instanceof Error ? error.message : String(error)}`)
        return false
      }
    },
    [getCurrentBytes],
  )

  const ensureSaved = useCallback(async (): Promise<boolean> => {
    const current = documentRef.current
    if (!current?.dirty) return true
    const choice = await window.nexoffice.confirmUnsaved(current.name)
    if (choice === 'cancel') return false
    if (choice === 'discard') return true
    return saveDocument(false)
  }, [saveDocument])

  const applyOpenedDocument = useCallback((opened: OpenedDocument) => {
    const kind = opened.kind
    if (!kind) {
      setStatus(`Unsupported file type: ${opened.name}`)
      return
    }
    setDocument({ ...opened, kind, dirty: false })
    setStatus(`Opened ${opened.name}`)
  }, [])

  const openDocument = useCallback(async () => {
    if (!(await ensureSaved())) return
    try {
      const opened = await window.nexoffice.openFile()
      if (!opened) return
      applyOpenedDocument(opened)
    } catch (error) {
      setStatus(`Open failed: ${error instanceof Error ? error.message : String(error)}`)
    }
  }, [ensureSaved, applyOpenedDocument])

  useEffect(() => {
    return window.nexoffice.onFileOpened((opened) => {
      void ensureSaved().then((proceed) => {
        if (proceed) applyOpenedDocument(opened)
      })
    })
  }, [ensureSaved, applyOpenedDocument])

  useEffect(() => {
    return window.nexoffice.onCloseRequest(() => {
      if (closingRef.current) return
      closingRef.current = true
      void ensureSaved()
        .then((proceed) => window.nexoffice.resolveClose(proceed))
        .finally(() => {
          closingRef.current = false
        })
    })
  }, [ensureSaved])

  useEffect(() => {
    const handle = (action: MenuAction) => {
      switch (action) {
        case 'file:new':
          void ensureSaved().then((proceed) => {
            if (!proceed) return
            setDocument(null)
            setStatus('New document')
          })
          break
        case 'file:open':
          void openDocument()
          break
        case 'file:save':
          void saveDocument(false)
          break
        case 'file:saveAs':
          void saveDocument(true)
          break
        case 'file:exportPdf': {
          const current = documentRef.current
          if (!current) {
            setStatus('Open a document first')
            break
          }
          void window.nexoffice
            .exportPdf(current.name)
            .then((result) => setStatus(result.path ? `Exported ${result.path}` : 'PDF export canceled'))
          break
        }
        case 'view:spellCheck':
          setSpellPanelOpen((prev) => !prev)
          break
        case 'view:wordCount': {
          const stats = docxRef.current?.getStats()
          setStatus(
            stats
              ? `Words: ${stats.words} · Chars: ${stats.characters} · Page ${stats.page}/${stats.pages}`
              : 'Open a Word document first',
          )
          break
        }
        case 'view:freezeTopRow':
          xlsxRef.current?.freezeTopRow()
          setStatus('Frozen top row')
          break
        case 'view:freezeFirstColumn':
          xlsxRef.current?.freezeFirstColumn()
          setStatus('Frozen first column')
          break
        case 'view:unfreeze':
          xlsxRef.current?.unfreeze()
          setStatus('Unfrozen panes')
          break
        default:
          setStatus(`${action} is not implemented yet`)
      }
    }
    return window.nexoffice.onMenuAction(handle)
  }, [openDocument, saveDocument, ensureSaved])

  useEffect(() => {
    window.nexoffice.rendererReady()
  }, [])

  const misspelledCount = misspellings.length

  return (
    <div className="flex h-full flex-col bg-neutral-100">
      <header className="drag-region flex h-11 shrink-0 items-center justify-center border-b border-neutral-200 bg-white">
        <span className="text-sm font-medium text-neutral-700">
          {document ? `${document.name}${document.dirty ? ' — Edited' : ''}` : 'NexOffice'}
        </span>
      </header>

      <div className="flex flex-1 overflow-hidden">
        <main className="flex flex-1 overflow-hidden">
          {document ? (
            document.kind === 'docx' ? (
              <DocxEditorView ref={docxRef} document={document} onChange={markEdited} />
            ) : document.kind === 'xlsx' ? (
              <XlsxEditorView ref={xlsxRef} document={document} onChange={markEdited} />
            ) : (
              <section className="flex w-full items-center justify-center p-8">
                <div className="w-full max-w-3xl rounded-lg border border-neutral-200 bg-white p-8 shadow-sm">
                  <h1 className="text-lg font-semibold text-neutral-900">{document.name}</h1>
                  <dl className="mt-4 grid grid-cols-[8rem_1fr] gap-y-2 text-sm text-neutral-600">
                    <dt>Type</dt>
                    <dd>{KIND_LABEL[document.kind]}</dd>
                    <dt>Size</dt>
                    <dd>{(document.data.byteLength / 1024).toFixed(1)} KB</dd>
                    <dt>Path</dt>
                    <dd className="truncate">{document.path}</dd>
                  </dl>
                  <p className="mt-6 text-sm text-neutral-500">Presentation editor coming soon.</p>
                </div>
              </section>
            )
          ) : (
            <section className="flex w-full items-center justify-center">
              <div className="text-center">
                <h1 className="text-2xl font-semibold text-neutral-900">NexOffice</h1>
                <p className="mt-2 text-sm text-neutral-500">
                  Open a .docx, .xlsx, or .pptx file to get started.
                </p>
                <button
                  type="button"
                  onClick={() => void openDocument()}
                  className="no-drag mt-6 rounded-md bg-neutral-900 px-4 py-2 text-sm font-medium text-white hover:bg-neutral-700"
                >
                  Open File
                </button>
              </div>
            </section>
          )}
        </main>

        {isDocx && (
          <SpellCheckPanel
            visible={spellPanelOpen}
            getText={getEditorText}
            onClose={() => setSpellPanelOpen(false)}
          />
        )}
      </div>

      <footer className="flex h-7 shrink-0 items-center gap-3 border-t border-neutral-200 bg-white px-3 text-xs text-neutral-500">
        <span>{status}</span>
        {isDocx && (
          <>
            <span className="hidden text-neutral-400 sm:inline">
              {docStats.words} words · {docStats.characters} chars · Page {docStats.page}/{docStats.pages}
            </span>
            {!spellLoading && !spellError && (
              <button
                onClick={() => setSpellPanelOpen((prev) => !prev)}
                className={`rounded px-1.5 py-0.5 ${misspelledCount > 0 ? 'text-red-500 hover:bg-red-50' : 'text-green-600 hover:bg-green-50'}`}
              >
                {misspelledCount > 0
                  ? `${misspelledCount} misspelled word${misspelledCount !== 1 ? 's' : ''}`
                  : 'Spelling OK'}
              </button>
            )}
            {spellLoading && <span className="text-neutral-400">Loading dictionary…</span>}
            {spellError && <span className="text-red-400">Dict error</span>}
          </>
        )}
      </footer>
    </div>
  )
}
