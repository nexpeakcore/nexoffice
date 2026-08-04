import { useCallback, useEffect, useRef, useState } from 'react'
import type { DocumentKind, MenuAction, OpenedDocument, WebEditAction } from '../shared/ipc.js'
import { SpellCheckPanel } from './components/SpellCheckPanel.js'
import { UpdateChip } from './components/UpdateChip.js'
import { DocxEditorView, type DocxEditorViewRef } from './editors/DocxEditorView.js'
import { XlsxEditorView, type XlsxEditorViewRef } from './editors/XlsxEditorView.js'
import { spellCheckService, type Misspelling } from './services/spellcheck.js'
import { useI18n } from './i18n.js'

interface DocumentState extends Omit<OpenedDocument, 'kind'> {
  kind: DocumentKind
  dirty: boolean
}

interface StatusMessage {
  key: string
  vars?: Record<string, string | number>
  suffixKey?: string
}

const KIND_LABEL_KEY: Record<DocumentKind, string> = {
  docx: 'app.kind.docx',
  xlsx: 'app.kind.xlsx',
  pptx: 'app.kind.pptx',
}

export function App() {
  const { t } = useI18n()
  const [document, setDocument] = useState<DocumentState | null>(null)
  const [status, setStatus] = useState<StatusMessage>({ key: 'status.ready' })
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
      if (!stats) return
      setDocStats((prev) =>
        prev.words === stats.words &&
        prev.characters === stats.characters &&
        prev.page === stats.page &&
        prev.pages === stats.pages
          ? prev
          : stats,
      )
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

  // Dirty flips state once; the revision (spell check trigger) is debounced so
  // steady typing causes no App re-render per keystroke.
  const editRevisionTimer = useRef<ReturnType<typeof setTimeout> | null>(null)
  const markEdited = useCallback(() => {
    setDocument((prev) => (prev && !prev.dirty ? { ...prev, dirty: true } : prev))
    if (editRevisionTimer.current) clearTimeout(editRevisionTimer.current)
    editRevisionTimer.current = setTimeout(() => {
      setEditRevision((revision) => revision + 1)
    }, 700)
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
    async (forceDialog: boolean, bytesOverride?: Uint8Array): Promise<boolean> => {
      const current = documentRef.current
      if (!current) return false
      try {
        const data = bytesOverride ?? (await getCurrentBytes(current))
        const save = forceDialog ? window.nexoffice.saveFileAs : window.nexoffice.saveFile
        const result = await save(current.path, current.kind, data)
        if (result.canceled || !result.path) {
          setStatus({ key: 'status.saveCanceled' })
          return false
        }
        const savedPath = result.path
        const savedName = savedPath.split(/[\\/]/).pop() ?? current.name
        setDocument((prev) =>
          prev ? { ...prev, path: savedPath, name: savedName, data, dirty: false } : prev,
        )
        setStatus({ key: 'status.saved', vars: { path: savedPath } })
        return true
      } catch (error) {
        setStatus({
          key: 'status.saveFailed',
          vars: { message: error instanceof Error ? error.message : String(error) },
        })
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
      setStatus({ key: 'status.unsupported', vars: { name: opened.name } })
      return
    }
    setDocument({ ...opened, kind, dirty: false })
    setStatus({ key: 'status.opened', vars: { name: opened.name } })
  }, [])

  const openDocument = useCallback(async () => {
    if (!(await ensureSaved())) return
    try {
      const opened = await window.nexoffice.openFile()
      if (!opened) return
      applyOpenedDocument(opened)
    } catch (error) {
      setStatus({
        key: 'status.openFailed',
        vars: { message: error instanceof Error ? error.message : String(error) },
      })
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

  // Edit-menu commands go to the office editor that owns the interaction; a
  // focused plain DOM field (formula bar, dialogs) keeps native behavior via
  // the main process, matching what the previous Electron role items did.
  const editActionTarget = useCallback((): 'docx' | 'xlsx' | 'dom' | null => {
    const kind = documentRef.current?.kind ?? null
    if (kind === 'docx' && docxRef.current?.isEditorFocused()) return 'docx'
    const active = window.document.activeElement
    if (
      active instanceof HTMLInputElement ||
      active instanceof HTMLTextAreaElement ||
      (active instanceof HTMLElement && active.isContentEditable)
    ) {
      return 'dom'
    }
    if (kind === 'xlsx') return 'xlsx'
    if (kind === 'docx') return 'docx'
    return null
  }, [])

  const runEditAction = useCallback(
    (verb: WebEditAction) => {
      const target = editActionTarget()
      if (target === 'dom') {
        window.nexoffice.webEditAction(verb)
        return
      }
      const editor = target === 'xlsx' ? xlsxRef.current : target === 'docx' ? docxRef.current : null
      if (!editor) return
      if (verb === 'undo') editor.undo()
      else if (verb === 'redo') editor.redo()
      else if (verb === 'cut') void editor.cut()
      else if (verb === 'copy') void editor.copy()
      else void editor.paste()
    },
    [editActionTarget],
  )

  useEffect(() => {
    const handle = (action: MenuAction) => {
      switch (action) {
        case 'file:new':
          void ensureSaved().then((proceed) => {
            if (!proceed) return
            setDocument(null)
            setStatus({ key: 'status.newDocument' })
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
            setStatus({ key: 'status.openFirst' })
            break
          }
          setStatus({ key: 'status.exportingPdf' })
          void getCurrentBytes(current)
            .then((data) => window.nexoffice.exportPdf(current.name, current.kind, data))
            .then((result) => {
              if (!result.path) {
                setStatus({ key: 'status.pdfCanceled' })
                return
              }
              const key =
                result.pages == null
                  ? 'status.exported'
                  : result.pages === 1
                    ? 'status.exportedPagesOne'
                    : 'status.exportedPagesMany'
              setStatus({
                key,
                vars: { path: result.path, pages: result.pages ?? 0 },
                ...(result.truncated ? { suffixKey: 'status.truncatedSuffix' } : {}),
              })
            })
            .catch((error: unknown) =>
              setStatus({
                key: 'status.pdfFailed',
                vars: { message: error instanceof Error ? error.message : String(error) },
              }),
            )
          break
        }
        case 'edit:undo':
          runEditAction('undo')
          break
        case 'edit:redo':
          runEditAction('redo')
          break
        case 'edit:cut':
          runEditAction('cut')
          break
        case 'edit:copy':
          runEditAction('copy')
          break
        case 'edit:paste':
          runEditAction('paste')
          break
        case 'view:spellCheck':
          setSpellPanelOpen((prev) => !prev)
          break
        case 'view:wordCount': {
          const stats = docxRef.current?.getStats()
          setStatus(
            stats
              ? {
                  key: 'status.wordCount',
                  vars: {
                    words: stats.words,
                    characters: stats.characters,
                    page: stats.page,
                    pages: stats.pages,
                  },
                }
              : { key: 'status.openWordFirst' },
          )
          break
        }
        case 'view:freezeTopRow':
          xlsxRef.current?.freezeTopRow()
          setStatus({ key: 'status.frozeTopRow' })
          break
        case 'view:freezeFirstColumn':
          xlsxRef.current?.freezeFirstColumn()
          setStatus({ key: 'status.frozeFirstColumn' })
          break
        case 'view:unfreeze':
          xlsxRef.current?.unfreeze()
          setStatus({ key: 'status.unfroze' })
          break
        default:
          setStatus({ key: 'status.notImplemented', vars: { action } })
      }
    }
    return window.nexoffice.onMenuAction(handle)
  }, [openDocument, saveDocument, ensureSaved, getCurrentBytes, runEditAction])

  useEffect(() => {
    window.nexoffice.rendererReady()
  }, [])

  const misspelledCount = misspellings.length

  return (
    <div className="flex h-full flex-col bg-neutral-100">
      <header className="drag-region flex h-11 shrink-0 items-center justify-center border-b border-neutral-200 bg-white">
        <span className="text-sm font-medium text-neutral-700">
          {document ? `${document.name}${document.dirty ? t('app.edited') : ''}` : 'NexOffice'}
        </span>
      </header>

      <div className="flex flex-1 overflow-hidden">
        <main className="flex flex-1 overflow-hidden">
          {document ? (
            document.kind === 'docx' ? (
              <DocxEditorView ref={docxRef} document={document} onChange={markEdited} />
            ) : document.kind === 'xlsx' ? (
              <XlsxEditorView
                ref={xlsxRef}
                document={document}
                onChange={markEdited}
                onSaveRequest={(bytes) => void saveDocument(false, bytes)}
              />
            ) : (
              <section className="flex w-full items-center justify-center p-8">
                <div className="w-full max-w-3xl rounded-lg border border-neutral-200 bg-white p-8 shadow-sm">
                  <h1 className="text-lg font-semibold text-neutral-900">{document.name}</h1>
                  <dl className="mt-4 grid grid-cols-[8rem_1fr] gap-y-2 text-sm text-neutral-600">
                    <dt>{t('app.meta.type')}</dt>
                    <dd>{t(KIND_LABEL_KEY[document.kind])}</dd>
                    <dt>{t('app.meta.size')}</dt>
                    <dd>{(document.data.byteLength / 1024).toFixed(1)} KB</dd>
                    <dt>{t('app.meta.path')}</dt>
                    <dd className="truncate">{document.path}</dd>
                  </dl>
                  <p className="mt-6 text-sm text-neutral-500">{t('app.presentationComingSoon')}</p>
                </div>
              </section>
            )
          ) : (
            <section className="flex w-full items-center justify-center">
              <div className="text-center">
                <h1 className="text-2xl font-semibold text-neutral-900">NexOffice</h1>
                <p className="mt-2 text-sm text-neutral-500">{t('app.empty.subtitle')}</p>
                <button
                  type="button"
                  onClick={() => void openDocument()}
                  className="no-drag mt-6 rounded-md bg-neutral-900 px-4 py-2 text-sm font-medium text-white hover:bg-neutral-700"
                >
                  {t('app.empty.openFile')}
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
        <span>{`${t(status.key, status.vars)}${status.suffixKey ? t(status.suffixKey) : ''}`}</span>
        {isDocx && (
          <>
            <span className="hidden text-neutral-400 sm:inline">
              {t('footer.stats', {
                words: docStats.words,
                characters: docStats.characters,
                page: docStats.page,
                pages: docStats.pages,
              })}
            </span>
            {!spellLoading && !spellError && (
              <button
                onClick={() => setSpellPanelOpen((prev) => !prev)}
                className={`rounded px-1.5 py-0.5 ${misspelledCount > 0 ? 'text-red-500 hover:bg-red-50' : 'text-green-600 hover:bg-green-50'}`}
              >
                {misspelledCount > 0
                  ? misspelledCount === 1
                    ? t('footer.misspelledOne')
                    : t('footer.misspelledMany', { count: misspelledCount })
                  : t('footer.spellingOk')}
              </button>
            )}
            {spellLoading && <span className="text-neutral-400">{t('footer.loadingDictionary')}</span>}
            {spellError && <span className="text-red-400">{t('footer.dictError')}</span>}
          </>
        )}
        <span className="ms-auto">
          <UpdateChip beforeRestart={ensureSaved} />
        </span>
      </footer>
    </div>
  )
}
