import { useEffect, useState } from 'react'
import { spellCheckService } from '../services/spellcheck.js'

interface SpellCheckPanelProps {
  visible: boolean
  getText: () => string
  onClose: () => void
}

export function SpellCheckPanel({ visible, getText, onClose }: SpellCheckPanelProps) {
  const [misspelled, setMisspelled] = useState<string[]>([])
  const [suggestions, setSuggestions] = useState<Record<string, string[]>>({})
  const [selectedWord, setSelectedWord] = useState<string | null>(null)

  useEffect(() => {
    if (!visible) return
    const refresh = () => {
      if (!spellCheckService.ready) return
      const results = spellCheckService.check(getText())
      setMisspelled(results.map((m) => m.word))
    }
    refresh()
    const interval = setInterval(refresh, 2000)
    return () => clearInterval(interval)
  }, [visible, getText])

  const handleWordClick = (word: string) => {
    setSelectedWord((prev) => (prev === word ? null : word))
    if (!suggestions[word]) {
      setSuggestions((prev) => ({
        ...prev,
        [word]: spellCheckService.suggest(word),
      }))
    }
  }

  if (!visible) return null

  return (
    <aside className="flex h-full w-64 shrink-0 flex-col border-l border-neutral-200 bg-white">
      <div className="flex items-center justify-between border-b border-neutral-200 px-3 py-2">
        <h2 className="text-xs font-semibold uppercase tracking-wide text-neutral-500">Spell Check</h2>
        <button
          onClick={onClose}
          className="rounded p-0.5 text-neutral-400 hover:bg-neutral-100 hover:text-neutral-600"
          aria-label="Close"
        >
          ✕
        </button>
      </div>

      <div className="flex-1 overflow-auto p-3">
        {misspelled.length === 0 ? (
          <p className="text-xs text-neutral-400">No misspelled words found.</p>
        ) : (
          <>
            <p className="mb-2 text-xs text-neutral-500">
              {misspelled.length} misspelled word{misspelled.length !== 1 ? 's' : ''}
            </p>
            <ul className="space-y-1">
              {misspelled.map((word) => (
                <li key={word}>
                  <button
                    onClick={() => handleWordClick(word)}
                    className={`w-full rounded px-2 py-1 text-left text-xs ${
                      selectedWord === word
                        ? 'bg-red-50 text-red-700'
                        : 'text-red-600 hover:bg-red-50'
                    }`}
                  >
                    {word}
                  </button>
                  {selectedWord === word && (suggestions[word]?.length ?? 0) > 0 && (
                    <ul className="ml-3 mt-1 space-y-0.5">
                      {(suggestions[word] ?? []).slice(0, 5).map((s) => (
                        <li
                          key={s}
                          className="rounded px-2 py-0.5 text-xs text-green-700"
                        >
                          {s}
                        </li>
                      ))}
                    </ul>
                  )}
                </li>
              ))}
            </ul>
          </>
        )}
      </div>
    </aside>
  )
}
