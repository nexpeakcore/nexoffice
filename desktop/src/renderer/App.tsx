import { useCallback, useEffect, useRef, useState } from 'react'
import {
  ALL_EDIT_CAPABILITIES,
  sameEditCapabilities,
  type DocumentKind,
  type EditCapabilities,
  type MenuAction,
  type OpenedDocument,
  type WebEditAction,
} from '../shared/ipc.js'
import { SpellCheckPanel } from './components/SpellCheckPanel.js'
import { UpdateChip } from './components/UpdateChip.js'
import { DocxEditorView, type DocxEditorViewRef } from './editors/DocxEditorView.js'
import { PptxEditorView, type PptxEditorViewRef } from './editors/PptxEditorView.js'
import { XlsxEditorView, type XlsxEditorViewRef } from './editors/XlsxEditorView.js'
import {
  editCapabilities,
  exportSuffixes,
  exportedStatusKey,
  saveOutcomeStatus,
  unsavedStep,
  type SaveOutcome,
  type StatusMessage,
  type TextSelectionState,
} from './services/documentPolicy.js'
import { spellCheckService, type Misspelling } from './services/spellcheck.js'
import type { EditorStats } from './services/textStats.js'
import { useI18n } from './i18n.js'

// `seed` replaces `OpenedDocument.data`: the bytes the editor mounted on. Its
// identity must stay stable for the life of the open document — both editors
// reload (disposing the live workbook/CRDT) whenever the buffer they were
// handed changes, so handing back the bytes a save serialized would throw away
// every edit made while that save was in flight. Saved bytes are not tracked:
// what is on disk lives on disk, and dirtiness is tracked by generation.
interface DocumentState extends Omit<OpenedDocument, 'kind' | 'data'> {
  kind: DocumentKind
  dirty: boolean
  id: number
  seed: Uint8Array
}

// The bytes to write, or the writer's account of the change it cannot express.
// A refusal is carried rather than thrown so it never reads as a failed write:
// nothing went wrong, the file on disk is simply still correct.
type Serialized =
  | { ok: true; bytes: Uint8Array }
  | { ok: false; refusal: string }

// The document a PDF export asked about, alongside the refusal that makes the
// question necessary.
interface StaleExport {
  document: DocumentState
  refusal: string
}

const ZOOM_MIN = 0.5
const ZOOM_MAX = 2
const ZOOM_STEP = 0.25

export function App() {
  const { t } = useI18n()
  const [document, setDocument] = useState<DocumentState | null>(null)
  const [status, setStatus] = useState<StatusMessage>({ key: 'status.ready' })
  const docxRef = useRef<DocxEditorViewRef>(null)
  const [spellPanelOpen, setSpellPanelOpen] = useState(false)
  const [spellLoading, setSpellLoading] = useState(true)
  const [spellError, setSpellError] = useState<string | null>(null)
  const [misspellings, setMisspellings] = useState<Misspelling[]>([])
  const [docStats, setDocStats] = useState<EditorStats | null>(null)
  const [editRevision, setEditRevision] = useState(0)
  const [refusal, setRefusal] = useState<string | null>(null)
  const [staleExportTarget, setStaleExportTarget] = useState<StaleExport | null>(null)
  const refusalRef = useRef<string | null>(null)
  const pptxSelectionRef = useRef<TextSelectionState | null>(null)
  const sentEditCapabilitiesRef = useRef<EditCapabilities>(ALL_EDIT_CAPABILITIES)
  const xlsxRef = useRef<XlsxEditorViewRef>(null)
  const pptxRef = useRef<PptxEditorViewRef>(null)
  const documentRef = useRef<DocumentState | null>(null)
  const closingRef = useRef(false)
  const editGenerationRef = useRef(0)
  const savedGenerationRef = useRef(0)
  const documentIdRef = useRef(0)

  useEffect(() => {
    documentRef.current = document
  }, [document])

  const documentKind = document?.kind ?? null

  // The menu lives in the main process and is rebuilt whole, so a caret moving
  // inside one paragraph must not reach it: only a flip in what the verbs can
  // act on is sent.
  const publishEditCapabilities = useCallback(() => {
    const next = editCapabilities(documentRef.current?.kind ?? null, pptxSelectionRef.current)
    if (sameEditCapabilities(next, sentEditCapabilitiesRef.current)) return
    sentEditCapabilitiesRef.current = next
    window.nexoffice.setEditCapabilities(next)
  }, [])

  useEffect(() => {
    window.nexoffice.setDocumentKind(documentKind)
    if (documentKind !== 'pptx') pptxSelectionRef.current = null
    publishEditCapabilities()
  }, [documentKind, publishEditCapabilities])

  const handlePptxSelectionChange = useCallback(
    (state: TextSelectionState) => {
      pptxSelectionRef.current = state
      publishEditCapabilities()
    },
    [publishEditCapabilities],
  )

  const isDocx = document?.kind === 'docx'
  const isPptx = document?.kind === 'pptx'
  const hasProofing = isDocx || isPptx

  // Read through the ref rather than a captured editor: the ticker and the
  // spell-check pass must always see whichever editor is mounted now.
  const proofingEditor = useCallback(
    (): { getText: () => string; getStats: () => EditorStats | null } | null => {
      const kind = documentRef.current?.kind
      if (kind === 'docx') return docxRef.current
      if (kind === 'pptx') return pptxRef.current
      return null
    },
    [],
  )

  useEffect(() => {
    if (!hasProofing) return
    const timer = setInterval(() => {
      const stats = proofingEditor()?.getStats()
      if (!stats) return
      setDocStats((prev) =>
        prev &&
        prev.words === stats.words &&
        prev.characters === stats.characters &&
        prev.page === stats.page &&
        prev.pages === stats.pages
          ? prev
          : stats,
      )
    }, 700)
    return () => clearInterval(timer)
  }, [hasProofing, proofingEditor])

  useEffect(() => {
    let canceled = false
    spellCheckService
      .init('en')
      .then(() => { if (!canceled) setSpellLoading(false) })
      .catch((err: Error) => { if (!canceled) { setSpellLoading(false); setSpellError(err.message) } })
    return () => { canceled = true }
  }, [])

  useEffect(() => {
    if (spellLoading || spellError || !hasProofing) {
      setMisspellings([])
      return
    }
    const timer = setTimeout(() => {
      const text = proofingEditor()?.getText() ?? ''
      setMisspellings(spellCheckService.check(text))
    }, 800)
    return () => clearTimeout(timer)
  }, [spellLoading, spellError, hasProofing, proofingEditor, document?.path, editRevision, docStats?.words])

  const getEditorText = useCallback(() => proofingEditor()?.getText() ?? '', [proofingEditor])

  // The revision (spell check trigger) is debounced so steady typing causes no
  // App re-render per keystroke; dirty, where it applies, flips state once.
  const editRevisionTimer = useRef<ReturnType<typeof setTimeout> | null>(null)
  const scheduleEditRevision = useCallback(() => {
    if (editRevisionTimer.current) clearTimeout(editRevisionTimer.current)
    editRevisionTimer.current = setTimeout(() => {
      setEditRevision((revision) => revision + 1)
    }, 700)
  }, [])

  // The last thing the writer said it could not express, or null while the
  // open document projects cleanly. It outlives the status line on purpose:
  // the refusal arrives after the change was already made, so the footer keeps
  // saying so until a projection succeeds rather than letting the notice scroll
  // away behind the next status.
  const noteRefusal = useCallback((message: string | null) => {
    refusalRef.current = message
    setRefusal(message)
  }, [])

  // An editor that fails to open reports no stats, and the ticker keeps the
  // last value it saw — which would be the previous document's.
  const clearStats = useCallback(() => {
    setDocStats(null)
    setMisspellings([])
  }, [])

  const markEdited = useCallback(() => {
    editGenerationRef.current += 1
    setDocument((prev) => (prev && !prev.dirty ? { ...prev, dirty: true } : prev))
    scheduleEditRevision()
  }, [scheduleEditRevision])

  // Every route to bytes runs through here — save, Save As and PDF export —
  // so one projection attempt decides all three, and the sticky refusal notice
  // is refreshed by whichever of them the user reached for.
  const serializeDocument = useCallback(
    async (current: DocumentState): Promise<Serialized> => {
      if (current.kind === 'docx') {
        const buffer = await docxRef.current?.save()
        if (buffer) return { ok: true, bytes: new Uint8Array(buffer) }
      } else if (current.kind === 'xlsx') {
        const bytes = xlsxRef.current?.save()
        if (bytes) return { ok: true, bytes }
      } else {
        try {
          const bytes = pptxRef.current?.save()
          if (bytes) {
            noteRefusal(null)
            return { ok: true, bytes }
          }
        } catch (error) {
          // The pptx writer refuses by name. Whatever it said is the only
          // account of the change that cannot be written, so it is passed on
          // untouched instead of being folded into a generic failure.
          const message = error instanceof Error ? error.message : String(error)
          noteRefusal(message)
          return { ok: false, refusal: message }
        }
      }
      // No editor yet (still loading) means the bytes the document opened with
      // are the only ones there are.
      return { ok: true, bytes: current.seed }
    },
    [noteRefusal],
  )

  const performSave = useCallback(
    async (
      target: DocumentState,
      forceDialog: boolean,
      bytesOverride: Uint8Array | undefined,
      overrideGeneration: number,
    ): Promise<SaveOutcome> => {
      const finish = (outcome: SaveOutcome): SaveOutcome => {
        setStatus(saveOutcomeStatus(outcome))
        return outcome
      }
      // The saved document is identified by the id captured at enqueue time;
      // a queued save picks up path renames from earlier saves via documentRef,
      // but never crosses over to a different document opened in the meantime.
      const live = documentRef.current
      const targetIsCurrent = live !== null && live.id === target.id
      const current = targetIsCurrent ? live : target
      // The editor only holds the current document, so a queued save whose
      // document was replaced before it ran has nothing valid to serialize.
      if (!bytesOverride && !targetIsCurrent) return { status: 'canceled' }
      // Override bytes were serialized when the save was requested, so they
      // only cover edits up to that point; freshly serialized bytes cover
      // everything up to now.
      const generation = bytesOverride ? overrideGeneration : editGenerationRef.current
      try {
        let data = bytesOverride
        if (!data) {
          const serialized = await serializeDocument(current)
          // Nothing was written, so the document stays dirty and the file on
          // disk stays the one that is still correct.
          if (!serialized.ok) return finish({ status: 'refused', message: serialized.refusal })
          data = serialized.bytes
        }
        const save = forceDialog ? window.nexoffice.saveFileAs : window.nexoffice.saveFile
        const result = await save(current.path, current.kind, data)
        if (result.canceled || !result.path) return finish({ status: 'canceled' })
        const savedPath = result.path
        const savedName = savedPath.split(/[\\/]/).pop() ?? current.name
        if (documentRef.current?.id !== target.id) {
          // The write targeted the captured document's own path, which stays
          // correct; the document open now is a different one, so its state
          // (path, name, bytes, dirty flag, generations) must not be touched.
          return finish({ status: 'saved', path: savedPath })
        }
        savedGenerationRef.current = Math.max(savedGenerationRef.current, generation)
        const dirty = editGenerationRef.current !== savedGenerationRef.current
        // Only the path, name and dirty flag move: `seed` stays the bytes the
        // editor mounted on, so a completed save never reloads the live editor
        // with the older snapshot it just wrote.
        // Queued saves read documentRef before React commits the state update,
        // so the new path/name must be visible synchronously.
        documentRef.current = { ...documentRef.current, path: savedPath, name: savedName, dirty }
        setDocument((prev) =>
          prev && prev.id === target.id
            ? { ...prev, path: savedPath, name: savedName, dirty }
            : prev,
        )
        return finish({ status: 'saved', path: savedPath })
      } catch (error) {
        return finish({
          status: 'failed',
          message: error instanceof Error ? error.message : String(error),
        })
      }
    },
    [serializeDocument],
  )

  // Saves are chained so serialize→write pairs never interleave: a slow older
  // write can otherwise land after a newer one and put stale bytes on disk.
  const saveQueueRef = useRef<Promise<unknown>>(Promise.resolve())

  const saveDocument = useCallback(
    (forceDialog: boolean, bytesOverride?: Uint8Array): Promise<SaveOutcome> => {
      const target = documentRef.current
      if (!target) return Promise.resolve<SaveOutcome>({ status: 'canceled' })
      const generationAtRequest = editGenerationRef.current
      const queued = saveQueueRef.current.then(
        () => performSave(target, forceDialog, bytesOverride, generationAtRequest),
        () => performSave(target, forceDialog, bytesOverride, generationAtRequest),
      )
      saveQueueRef.current = queued
      return queued
    },
    [performSave],
  )

  const hasUnsavedEdits = useCallback(
    () => documentRef.current !== null && editGenerationRef.current !== savedGenerationRef.current,
    [],
  )

  const ensureSaved = useCallback(async (): Promise<boolean> => {
    while (hasUnsavedEdits()) {
      const current = documentRef.current
      if (!current) return true
      const choice = await window.nexoffice.confirmUnsaved(current.name)
      if (choice === 'cancel') return false
      if (choice === 'discard') return true
      const step = unsavedStep(await saveDocument(false))
      if (step.step === 'stop') return false
      // The writer cannot express this change, so asking for the save again
      // would only refuse again. Leaving without it has to stay possible, and
      // the prompt says in the same breath that the change is what gets lost.
      if (step.step === 'escape') {
        return (await window.nexoffice.confirmSaveRefused(current.name, step.message)) === 'discard'
      }
    }
    return true
  }, [hasUnsavedEdits, saveDocument])

  const applyOpenedDocument = useCallback((opened: OpenedDocument) => {
    const kind = opened.kind
    if (!kind) {
      setStatus({ key: 'status.unsupported', vars: { name: opened.name } })
      return
    }
    savedGenerationRef.current = editGenerationRef.current
    noteRefusal(null)
    setStaleExportTarget(null)
    clearStats()
    documentIdRef.current += 1
    const { data, ...rest } = opened
    const next: DocumentState = { ...rest, kind, seed: data, dirty: false, id: documentIdRef.current }
    // In-flight save completions compare against documentRef, so the switch
    // must be visible before React commits the state update.
    documentRef.current = next
    // The editor that reported it is being remounted, so nothing is selected
    // until it says otherwise.
    pptxSelectionRef.current = null
    publishEditCapabilities()
    setDocument(next)
    setStatus({ key: 'status.opened', vars: { name: opened.name } })
  }, [clearStats, noteRefusal, publishEditCapabilities])

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
  const editActionTarget = useCallback((): 'docx' | 'xlsx' | 'pptx' | 'dom' | null => {
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
    if (kind === 'pptx') return 'pptx'
    return null
  }, [])

  const runEditAction = useCallback(
    (verb: WebEditAction) => {
      const target = editActionTarget()
      if (target === 'dom') {
        window.nexoffice.webEditAction(verb)
        return
      }
      const editor =
        target === 'xlsx'
          ? xlsxRef.current
          : target === 'docx'
            ? docxRef.current
            : target === 'pptx'
              ? pptxRef.current
              : null
      if (!editor) return
      if (verb === 'undo') editor.undo()
      else if (verb === 'redo') editor.redo()
      else if (verb === 'cut') void editor.cut()
      else if (verb === 'copy') void editor.copy()
      else if (verb === 'delete') editor.deleteSelection()
      else if (verb === 'selectAll') editor.selectAll()
      else void editor.paste()
    },
    [editActionTarget],
  )

  const applyZoom = useCallback((step: number | null) => {
    const kind = documentRef.current?.kind
    const editor =
      kind === 'docx'
        ? docxRef.current
        : kind === 'xlsx'
          ? xlsxRef.current
          : kind === 'pptx'
            ? pptxRef.current
            : null
    if (!editor) return
    const next =
      step === null
        ? 1
        : Math.min(ZOOM_MAX, Math.max(ZOOM_MIN, Math.round((editor.getZoom() + step) * 100) / 100))
    editor.setZoom(next)
  }, [])

  const runPdfExport = useCallback(
    (target: DocumentState, data: Uint8Array, asOpened: boolean) => {
      setStatus({ key: 'status.exportingPdf' })
      void window.nexoffice
        .exportPdf(target.name, target.kind, data)
        .then((result) => {
          if (!result.path) {
            setStatus({ key: 'status.pdfCanceled' })
            return
          }
          const suffixes = exportSuffixes({
            truncated: result.truncated === true,
            skipped: result.skipped ?? 0,
            asOpened,
          })
          setStatus({
            key: exportedStatusKey(result.pages),
            vars: { path: result.path, pages: result.pages ?? 0 },
            ...(suffixes.length > 0 ? { suffixes } : {}),
          })
        })
        .catch((error: unknown) =>
          setStatus({
            key: 'status.pdfFailed',
            vars: { message: error instanceof Error ? error.message : String(error) },
          }),
        )
    },
    [],
  )

  // The export renders whatever the document can be written as, so it matches
  // the screen wherever the writer can express the edits. Only a refusal sends
  // it back to the bytes the file was opened with, and then the user is asked
  // first rather than handed a PDF that contradicts what they are looking at.
  const exportPdf = useCallback(
    (target: DocumentState) => {
      setStatus({ key: 'status.exportingPdf' })
      void serializeDocument(target).then((serialized) => {
        if (serialized.ok) {
          runPdfExport(target, serialized.bytes, false)
          return
        }
        setStatus(saveOutcomeStatus({ status: 'refused', message: serialized.refusal }))
        setStaleExportTarget({ document: target, refusal: serialized.refusal })
      })
    },
    [runPdfExport, serializeDocument],
  )

  const cancelStaleExport = useCallback(() => {
    setStaleExportTarget(null)
    setStatus({ key: 'status.pdfCanceled' })
  }, [])

  useEffect(() => {
    if (!staleExportTarget) return
    const handle = (event: KeyboardEvent) => {
      if (event.key === 'Escape') cancelStaleExport()
    }
    window.addEventListener('keydown', handle)
    return () => window.removeEventListener('keydown', handle)
  }, [staleExportTarget, cancelStaleExport])

  useEffect(() => {
    const handle = (action: MenuAction) => {
      switch (action) {
        case 'file:new':
          void ensureSaved().then((proceed) => {
            if (!proceed) return
            savedGenerationRef.current = editGenerationRef.current
            noteRefusal(null)
            setStaleExportTarget(null)
            documentRef.current = null
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
          exportPdf(current)
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
        case 'edit:delete':
          runEditAction('delete')
          break
        case 'edit:selectAll':
          runEditAction('selectAll')
          break
        case 'edit:find':
          if (documentRef.current?.kind === 'docx') docxRef.current?.openFind()
          break
        case 'view:zoomIn':
          applyZoom(ZOOM_STEP)
          break
        case 'view:zoomOut':
          applyZoom(-ZOOM_STEP)
          break
        case 'view:zoomReset':
          applyZoom(null)
          break
        case 'view:spellCheck':
          setSpellPanelOpen((prev) => !prev)
          break
        case 'view:wordCount': {
          const kind = documentRef.current?.kind
          const stats = proofingEditor()?.getStats()
          if (!stats) {
            setStatus({ key: kind === 'pptx' ? 'status.openFirst' : 'status.openWordFirst' })
            break
          }
          setStatus(
            kind === 'pptx'
              ? {
                  key: 'status.wordCountSlides',
                  vars: {
                    words: stats.words,
                    characters: stats.characters,
                    slide: stats.page,
                    slides: stats.pages,
                  },
                }
              : {
                  key: 'status.wordCount',
                  vars: {
                    words: stats.words,
                    characters: stats.characters,
                    page: stats.page,
                    pages: stats.pages,
                  },
                },
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
      }
    }
    return window.nexoffice.onMenuAction(handle)
  }, [
    openDocument,
    saveDocument,
    ensureSaved,
    noteRefusal,
    exportPdf,
    runEditAction,
    applyZoom,
    proofingEditor,
  ])

  useEffect(() => {
    window.nexoffice.rendererReady()
  }, [])

  const misspelledCount = misspellings.length

  return (
    <div className="relative flex h-full flex-col bg-neutral-100">
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
              <PptxEditorView
                ref={pptxRef}
                document={document}
                onChange={markEdited}
                onSelectionStateChange={handlePptxSelectionChange}
              />
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

        {hasProofing && (
          <SpellCheckPanel
            visible={spellPanelOpen}
            getText={getEditorText}
            onClose={() => setSpellPanelOpen(false)}
          />
        )}
      </div>

      <footer className="flex h-7 shrink-0 items-center gap-3 border-t border-neutral-200 bg-white px-3 text-xs text-neutral-500">
        <span>
          {`${t(status.key, status.vars)}${(status.suffixes ?? []).map((suffix) => t(suffix.key, suffix.vars)).join('')}`}
        </span>
        {refusal && (
          <button
            type="button"
            onClick={() => setStatus({ key: 'status.saveRefused', vars: { message: refusal } })}
            className="rounded px-1.5 py-0.5 text-amber-600 hover:bg-amber-50"
          >
            {t('footer.saveRefused')}
          </button>
        )}
        {hasProofing && (
          <>
            {docStats && (
            <span className="hidden text-neutral-400 sm:inline">
              {isPptx
                ? t('footer.statsSlides', {
                    words: docStats.words,
                    characters: docStats.characters,
                    slide: docStats.page,
                    slides: docStats.pages,
                  })
                : t('footer.stats', {
                    words: docStats.words,
                    characters: docStats.characters,
                    page: docStats.page,
                    pages: docStats.pages,
                  })}
            </span>
            )}
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

      {staleExportTarget && (
        <div className="absolute inset-0 z-50 flex items-center justify-center bg-neutral-900/40 p-6">
          <div
            role="dialog"
            aria-modal="true"
            aria-labelledby="stale-export-title"
            className="w-full max-w-md rounded-lg bg-white p-5 shadow-xl"
          >
            <h2 id="stale-export-title" className="text-sm font-semibold text-neutral-900">
              {t('exportStale.message', { name: staleExportTarget.document.name })}
            </h2>
            <p className="mt-2 text-sm text-neutral-600">{t('exportStale.detail')}</p>
            <p className="mt-2 text-sm text-neutral-500">{staleExportTarget.refusal}</p>
            <div className="mt-5 flex justify-end gap-2">
              <button
                type="button"
                autoFocus
                onClick={cancelStaleExport}
                className="rounded-md border border-neutral-300 px-3 py-1.5 text-sm text-neutral-700 hover:bg-neutral-100"
              >
                {t('exportStale.cancel')}
              </button>
              <button
                type="button"
                onClick={() => {
                  const target = staleExportTarget.document
                  setStaleExportTarget(null)
                  runPdfExport(target, target.seed, true)
                }}
                className="rounded-md bg-neutral-900 px-3 py-1.5 text-sm font-medium text-white hover:bg-neutral-700"
              >
                {t('exportStale.confirm')}
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  )
}
