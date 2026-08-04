import { useEffect, useRef, useState } from 'react'
import type { UpdateEvent } from '../../shared/ipc.js'
import { useI18n } from '../i18n.js'

const TRANSIENT_STATUSES = new Set<UpdateEvent['status']>(['none', 'dev', 'error'])
const TRANSIENT_DISMISS_MS = 6000

interface UpdateChipProps {
  beforeRestart: () => Promise<boolean>
}

export function UpdateChip({ beforeRestart }: UpdateChipProps) {
  const { t } = useI18n()
  const [event, setEvent] = useState<UpdateEvent | null>(null)
  const [restarting, setRestarting] = useState(false)
  const dismissTimer = useRef<ReturnType<typeof setTimeout> | null>(null)

  useEffect(() => {
    const unsubscribe = window.nexoffice.onUpdateEvent((update) => {
      if (dismissTimer.current) clearTimeout(dismissTimer.current)
      setEvent(update)
      if (TRANSIENT_STATUSES.has(update.status)) {
        dismissTimer.current = setTimeout(() => setEvent(null), TRANSIENT_DISMISS_MS)
      }
    })
    return () => {
      unsubscribe()
      if (dismissTimer.current) clearTimeout(dismissTimer.current)
    }
  }, [])

  if (!event) return null

  if (event.status === 'downloaded') {
    return (
      <span className="flex items-center gap-1.5 text-neutral-400">
        <span>{t('update.ready', { version: event.version })}</span>
        <button
          type="button"
          disabled={restarting}
          onClick={() => {
            setRestarting(true)
            void beforeRestart()
              .then((proceed) => {
                if (proceed) window.nexoffice.installUpdate()
              })
              .finally(() => setRestarting(false))
          }}
          className="rounded bg-neutral-900 px-1.5 py-0.5 font-medium text-white hover:bg-neutral-700 disabled:opacity-50"
        >
          {t('update.restart')}
        </button>
      </span>
    )
  }

  const label =
    event.status === 'checking'
      ? t('update.checking')
      : event.status === 'available'
        ? t('update.downloading', { version: event.version })
        : event.status === 'progress'
          ? t('update.downloadingPercent', { percent: event.percent })
          : event.status === 'none'
            ? t('update.upToDate')
            : event.status === 'dev'
              ? t('update.devDisabled')
              : t('update.failed', { message: event.message })

  return <span className={event.status === 'error' ? 'text-red-400' : 'text-neutral-400'}>{label}</span>
}
