import { useCallback, useEffect, useRef, useState } from 'react'
import type { AgentSettings, DocumentKind } from '../../shared/ipc.js'
import type { CreateChartProposal, WriteCellsProposal } from '../services/agentTools.js'
import { useI18n } from '../i18n.js'

interface ChatMessage {
  role: 'user' | 'assistant'
  content: string
}

type PendingProposal =
  | { toolCallId: string; kind: 'write'; proposal: WriteCellsProposal }
  | { toolCallId: string; kind: 'chart'; proposal: CreateChartProposal }

export interface AgentDocumentBridge {
  runReadTool: (name: string, args: Record<string, unknown>) => unknown
  validateWrite: (
    args: Record<string, unknown>
  ) => { proposal: WriteCellsProposal } | { error: string }
  applyWrite: (proposal: WriteCellsProposal) => unknown
  validateChart: (
    args: Record<string, unknown>
  ) => { proposal: CreateChartProposal } | { error: string }
  applyChart: (proposal: CreateChartProposal) => unknown
}

interface AgentPanelProps {
  visible: boolean
  documentKind: DocumentKind
  documentName: string
  bridge: AgentDocumentBridge
  /** A write proposal arriving while the panel is hidden must surface it. */
  onRequestOpen: () => void
  onClose: () => void
}

export function AgentPanel({
  visible,
  documentKind,
  documentName,
  bridge,
  onRequestOpen,
  onClose,
}: AgentPanelProps) {
  const { locale, t } = useI18n()
  const [settings, setSettings] = useState<AgentSettings | null>(null)
  const [keyDraft, setKeyDraft] = useState('')
  const [messages, setMessages] = useState<ChatMessage[]>([])
  const [draft, setDraft] = useState('')
  const [busy, setBusy] = useState(false)
  const [notice, setNotice] = useState<string | null>(null)
  const [pending, setPending] = useState<PendingProposal | null>(null)
  const scrollRef = useRef<HTMLDivElement | null>(null)
  const bridgeRef = useRef(bridge)
  bridgeRef.current = bridge
  const onRequestOpenRef = useRef(onRequestOpen)
  onRequestOpenRef.current = onRequestOpen
  const pendingRef = useRef<PendingProposal | null>(null)
  pendingRef.current = pending

  useEffect(() => {
    if (!visible) return
    void window.nexoffice.agentGetSettings().then(setSettings)
  }, [visible])

  useEffect(() => {
    const offEvent = window.nexoffice.onAgentEvent((event) => {
      switch (event.type) {
        case 'text':
          setMessages((prev) => {
            const last = prev.at(-1)
            if (last?.role === 'assistant') {
              return [...prev.slice(0, -1), { role: 'assistant', content: last.content + event.delta }]
            }
            return [...prev, { role: 'assistant', content: event.delta }]
          })
          break
        case 'tool':
          setNotice(`${event.name}(${event.summary})`)
          break
        case 'done':
          setBusy(false)
          setNotice(null)
          break
        case 'error':
          setBusy(false)
          setNotice(null)
          setMessages((prev) => [
            ...prev,
            {
              role: 'assistant',
              content:
                event.message === 'missing-api-key'
                  ? '⚠ missing-api-key'
                  : `⚠ ${event.message}`,
            },
          ])
          break
      }
    })
    const offTool = window.nexoffice.onAgentToolRequest((request) => {
      if (request.name === 'write_cells') {
        const validated = bridgeRef.current.validateWrite(request.args)
        if ('error' in validated) {
          window.nexoffice.agentToolResult(request.id, validated)
          return
        }
        setPending({ toolCallId: request.id, kind: 'write', proposal: validated.proposal })
        onRequestOpenRef.current()
        return
      }
      if (request.name === 'create_chart') {
        const validated = bridgeRef.current.validateChart(request.args)
        if ('error' in validated) {
          window.nexoffice.agentToolResult(request.id, validated)
          return
        }
        setPending({ toolCallId: request.id, kind: 'chart', proposal: validated.proposal })
        onRequestOpenRef.current()
        return
      }
      let result: unknown
      try {
        result = bridgeRef.current.runReadTool(request.name, request.args)
      } catch (error) {
        result = { error: error instanceof Error ? error.message : String(error) }
      }
      window.nexoffice.agentToolResult(request.id, result ?? null)
    })
    return () => {
      offEvent()
      offTool()
    }
  }, [])

  useEffect(() => {
    scrollRef.current?.scrollTo({ top: scrollRef.current.scrollHeight })
  }, [messages, busy])

  const send = useCallback(() => {
    const text = draft.trim()
    if (!text || busy) return
    const next = [...messages, { role: 'user' as const, content: text }]
    setMessages(next)
    setDraft('')
    setBusy(true)
    void window.nexoffice.agentRun({
      messages: next,
      documentKind,
      documentName,
      locale,
    })
  }, [draft, busy, messages, documentKind, documentName, locale])

  const resolveProposal = useCallback((approve: boolean) => {
    const current = pendingRef.current
    if (!current) return
    setPending(null)
    let result: unknown
    if (approve) {
      try {
        result =
          current.kind === 'chart'
            ? bridgeRef.current.applyChart(current.proposal)
            : bridgeRef.current.applyWrite(current.proposal)
      } catch (error) {
        result = { error: error instanceof Error ? error.message : String(error) }
      }
    } else {
      result = { rejected: true, reason: 'the user rejected this proposal' }
    }
    window.nexoffice.agentToolResult(current.toolCallId, result ?? null)
  }, [])

  const cancelRun = useCallback(() => {
    // A pending proposal belongs to the run being cancelled — resolve it so
    // the main-process tool round-trip never dangles until its timeout.
    resolveProposal(false)
    window.nexoffice.agentCancel()
  }, [resolveProposal])

  const saveKey = useCallback(() => {
    const apiKey = keyDraft.trim()
    if (!apiKey) return
    void window.nexoffice.agentSetSettings({ apiKey }).then((updated) => {
      setSettings(updated)
      setKeyDraft('')
    })
  }, [keyDraft])

  if (!visible) return null

  return (
    <aside className="flex h-full w-80 shrink-0 flex-col border-s border-neutral-200 bg-white">
      <div className="flex items-center justify-between border-b border-neutral-200 px-3 py-2">
        <h2 className="text-xs font-semibold uppercase tracking-wide text-neutral-500">
          {t('agentPanel.title')}
        </h2>
        <button
          onClick={onClose}
          className="rounded p-0.5 text-neutral-400 hover:bg-neutral-100 hover:text-neutral-600"
          aria-label={t('agentPanel.close')}
        >
          ✕
        </button>
      </div>

      {settings && !settings.hasApiKey ? (
        <div className="flex flex-1 flex-col gap-2 p-3">
          <p className="text-xs text-neutral-600">{t('agentPanel.needKey')}</p>
          <input
            type="password"
            value={keyDraft}
            onChange={(event) => setKeyDraft(event.target.value)}
            onKeyDown={(event) => {
              if (event.key === 'Enter') saveKey()
            }}
            placeholder={t('agentPanel.keyPlaceholder')}
            className="rounded border border-neutral-300 px-2 py-1.5 text-xs"
          />
          <button
            onClick={saveKey}
            disabled={!keyDraft.trim()}
            className="rounded bg-neutral-900 px-3 py-1.5 text-xs font-medium text-white disabled:opacity-40"
          >
            {t('agentPanel.saveKey')}
          </button>
        </div>
      ) : (
        <>
          <div ref={scrollRef} className="flex-1 space-y-3 overflow-auto p-3">
            {messages.length === 0 && (
              <p className="text-xs text-neutral-400">{t('agentPanel.empty')}</p>
            )}
            {messages.map((message, index) =>
              message.content === '⚠ missing-api-key' ? (
                <p key={index} className="text-xs text-amber-700">
                  {t('agentPanel.needKey')}
                </p>
              ) : (
                <div
                  key={index}
                  className={
                    message.role === 'user'
                      ? 'ms-6 rounded-lg bg-neutral-100 px-2.5 py-1.5 text-xs text-neutral-800'
                      : 'me-2 whitespace-pre-wrap text-xs text-neutral-800'
                  }
                >
                  {message.content}
                </div>
              )
            )}
            {pending && (
              <div className="rounded-lg border border-amber-300 bg-amber-50 p-2">
                <p className="text-xs font-medium text-amber-900">
                  {pending.kind === 'write'
                    ? t('agentPanel.proposalTitle', {
                        count: pending.proposal.edits.length,
                        sheet: pending.proposal.sheetName,
                      })
                    : t('agentPanel.chartProposalTitle', {
                        type: pending.proposal.chartType,
                        sheet: pending.proposal.sheetName,
                      })}
                </p>
                <div className="mt-1.5 max-h-40 overflow-auto">
                  {pending.kind === 'write' ? (
                    <table className="w-full text-xs text-neutral-800">
                      <tbody>
                        {pending.proposal.edits.map((edit) => (
                          <tr key={edit.a1} className="border-t border-amber-200/60">
                            <td className="py-0.5 pe-2 font-mono text-neutral-500">{edit.a1}</td>
                            <td className="break-all py-0.5 font-mono">
                              {edit.input === '' ? '∅' : edit.input}
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  ) : (
                    <div className="space-y-0.5 text-xs text-neutral-800">
                      {pending.proposal.title && (
                        <p className="font-medium">{pending.proposal.title}</p>
                      )}
                      <p>
                        <span className="text-neutral-500">{t('agentPanel.chartAnchor')} </span>
                        <span className="font-mono">{pending.proposal.anchor}</span>
                      </p>
                      {pending.proposal.categories && (
                        <p>
                          <span className="text-neutral-500">{t('agentPanel.chartCategories')} </span>
                          <span className="font-mono">{pending.proposal.categories}</span>
                        </p>
                      )}
                      {pending.proposal.series.map((series, index) => (
                        <p key={index}>
                          <span className="text-neutral-500">
                            {series.name ?? t('agentPanel.chartSeries', { index: index + 1 })}{' '}
                          </span>
                          <span className="font-mono">{series.values}</span>
                        </p>
                      ))}
                    </div>
                  )}
                </div>
                <div className="mt-2 flex justify-end gap-2">
                  <button
                    onClick={() => resolveProposal(false)}
                    className="rounded px-2 py-1 text-xs text-neutral-600 hover:bg-amber-100"
                  >
                    {t('agentPanel.reject')}
                  </button>
                  <button
                    onClick={() => resolveProposal(true)}
                    className="rounded bg-amber-600 px-3 py-1 text-xs font-medium text-white hover:bg-amber-700"
                  >
                    {t('agentPanel.apply')}
                  </button>
                </div>
              </div>
            )}
            {busy && !pending && (
              <p className="text-xs text-neutral-400">
                {notice ? t('agentPanel.working', { tool: notice }) : t('agentPanel.thinking')}
              </p>
            )}
          </div>
          <div className="border-t border-neutral-200 p-2">
            <textarea
              value={draft}
              onChange={(event) => setDraft(event.target.value)}
              onKeyDown={(event) => {
                if (event.key === 'Enter' && !event.shiftKey) {
                  event.preventDefault()
                  send()
                }
              }}
              placeholder={t('agentPanel.placeholder')}
              rows={2}
              className="w-full resize-none rounded border border-neutral-300 px-2 py-1.5 text-xs"
            />
            <div className="mt-1 flex justify-end gap-2">
              {busy && (
                <button
                  onClick={cancelRun}
                  className="rounded px-2 py-1 text-xs text-neutral-500 hover:bg-neutral-100"
                >
                  {t('agentPanel.cancel')}
                </button>
              )}
              <button
                onClick={send}
                disabled={busy || !draft.trim()}
                className="rounded bg-neutral-900 px-3 py-1 text-xs font-medium text-white disabled:opacity-40"
              >
                {t('agentPanel.send')}
              </button>
            </div>
          </div>
        </>
      )}
    </aside>
  )
}
