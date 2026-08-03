import { useEffect, useRef, useState } from 'react'
import type { UpdateEvent } from '../../shared/ipc.js'

const TRANSIENT_STATUSES = new Set<UpdateEvent['status']>(['none', 'dev', 'error'])
const TRANSIENT_DISMISS_MS = 6000

interface UpdateChipProps {
  beforeRestart: () => Promise<boolean>
}

export function UpdateChip({ beforeRestart }: UpdateChipProps) {
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
        <span>Update {event.version} ready</span>
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
          Restart to update
        </button>
      </span>
    )
  }

  const label =
    event.status === 'checking'
      ? 'Checking for updates…'
      : event.status === 'available'
        ? `Downloading update ${event.version}…`
        : event.status === 'progress'
          ? `Downloading update… ${event.percent}%`
          : event.status === 'none'
            ? 'NexOffice is up to date'
            : event.status === 'dev'
              ? 'Updates are disabled in dev builds'
              : `Update check failed: ${event.message}`

  return <span className={event.status === 'error' ? 'text-red-400' : 'text-neutral-400'}>{label}</span>
}
