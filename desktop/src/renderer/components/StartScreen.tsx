import { useCallback, useEffect, useState } from 'react'
import type { DocumentKind, RecentFile } from '../../shared/ipc.js'
import { useI18n } from '../i18n.js'
import aiIcon from '../assets/brand/nex-ai-icon.svg'
import docsIcon from '../assets/brand/nex-docs-icon.svg'
import sheetsIcon from '../assets/brand/nex-sheets-icon.svg'
import slidesIcon from '../assets/brand/nex-slides-icon.svg'

const KIND_ICONS: Record<DocumentKind, string> = {
  docx: docsIcon,
  xlsx: sheetsIcon,
  pptx: slidesIcon,
}

interface StartScreenProps {
  onNew: (kind: DocumentKind) => void
  onOpen: () => void
  onOpenRecent: (path: string) => void
}

export function StartScreen({ onNew, onOpen, onOpenRecent }: StartScreenProps) {
  const { t } = useI18n()
  const [recents, setRecents] = useState<RecentFile[]>([])
  const [version, setVersion] = useState('')

  const refresh = useCallback(() => {
    void window.nexoffice.recentsList().then(setRecents)
  }, [])

  useEffect(() => {
    refresh()
    window.addEventListener('focus', refresh)
    return () => window.removeEventListener('focus', refresh)
  }, [refresh])

  useEffect(() => {
    void window.nexoffice.version().then(setVersion)
  }, [])

  const remove = useCallback((path: string) => {
    void window.nexoffice.recentsRemove(path).then(setRecents)
  }, [])

  const newKinds: Array<{ kind: DocumentKind; label: string }> = [
    { kind: 'docx', label: t('start.newDocument') },
    { kind: 'xlsx', label: t('start.newWorkbook') },
    { kind: 'pptx', label: t('start.newPresentation') },
  ]

  return (
    <section className="flex w-full justify-center overflow-auto">
      <div className="w-full max-w-2xl px-8 py-12">
        <div className="flex items-center gap-3">
          <img src={aiIcon} alt="" className="h-10 w-10" />
          <h1 className="text-3xl font-bold tracking-tight">
            <span className="bg-gradient-to-r from-violet-600 via-purple-500 to-indigo-600 bg-clip-text text-transparent">
              Nex
            </span>
            <span className="text-neutral-900">Office</span>
          </h1>
        </div>
        <p className="mt-2 flex items-center gap-2 text-sm text-neutral-500">
          <span>{t('app.empty.subtitle')}</span>
          {version && (
            <button
              type="button"
              onClick={() => void window.nexoffice.checkForUpdates()}
              title={t('menu.app.checkForUpdates')}
              className="no-drag text-xs text-neutral-400 transition hover:text-neutral-600"
            >
              v{version}
            </button>
          )}
        </p>

        <div className="mt-8 grid grid-cols-3 gap-3">
          {newKinds.map(({ kind, label }) => (
            <button
              key={kind}
              type="button"
              onClick={() => onNew(kind)}
              className="no-drag flex flex-col items-center gap-3 rounded-xl border border-neutral-200 bg-white px-4 py-6 text-sm font-medium text-neutral-800 shadow-sm transition hover:border-neutral-300 hover:shadow"
            >
              <img src={KIND_ICONS[kind]} alt="" className="h-12 w-12" />
              {label}
            </button>
          ))}
        </div>

        <div className="mt-10 flex items-center justify-between">
          <h2 className="text-sm font-semibold text-neutral-700">{t('start.recent')}</h2>
          <button
            type="button"
            onClick={onOpen}
            className="no-drag rounded-md px-3 py-1.5 text-sm font-medium text-neutral-600 hover:bg-neutral-100"
          >
            {t('app.empty.openFile')}
          </button>
        </div>
        {recents.length === 0 ? (
          <p className="mt-3 text-sm text-neutral-400">{t('start.noRecent')}</p>
        ) : (
          <ul className="mt-3 divide-y divide-neutral-100 rounded-xl border border-neutral-200 bg-white shadow-sm">
            {recents.map((entry) => (
              <li key={entry.path} className="group flex items-center gap-3 px-4 py-2.5">
                <button
                  type="button"
                  disabled={!entry.exists}
                  onClick={() => onOpenRecent(entry.path)}
                  className="no-drag flex min-w-0 flex-1 items-center gap-3 text-left disabled:cursor-default"
                >
                  {entry.kind ? (
                    <img src={KIND_ICONS[entry.kind]} alt="" className="h-7 w-7 shrink-0" />
                  ) : (
                    <span className="h-7 w-7 shrink-0 rounded bg-neutral-200" />
                  )}
                  <span className="min-w-0">
                    <span
                      className={`block truncate text-sm ${entry.exists ? 'text-neutral-800' : 'text-neutral-400 line-through'}`}
                    >
                      {entry.name}
                    </span>
                    <span className="block truncate text-xs text-neutral-400">{entry.path}</span>
                  </span>
                </button>
                <button
                  type="button"
                  onClick={() => remove(entry.path)}
                  aria-label={t('start.remove')}
                  title={t('start.remove')}
                  className="no-drag rounded p-1 text-neutral-300 opacity-0 transition group-hover:opacity-100 hover:bg-neutral-100 hover:text-neutral-600"
                >
                  ✕
                </button>
              </li>
            ))}
          </ul>
        )}
      </div>
    </section>
  )
}
