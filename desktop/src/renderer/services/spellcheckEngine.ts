import Typo from 'typo-js'

export interface Misspelling {
  word: string
}

const MAX_UNIQUE_WORDS = 2000
const VERDICT_CACHE_CAP = 20000

/**
 * Hunspell-backed proofing over plain document text. Building this costs
 * ~40MB of heap (typo-js expands en_US into JS structures), so it is meant to
 * live in the spell-check worker, not the renderer's main thread.
 */
export class SpellCheckEngine {
  private spellchecker: Typo | null = null
  private verdicts = new Map<string, boolean>()
  ready = false

  async init(): Promise<void> {
    if (this.ready) return
    const [dic, aff] = await Promise.all([
      fetch('/dictionaries/en_US.dic').then((r) => r.text()),
      fetch('/dictionaries/en_US.aff').then((r) => r.text()),
    ])
    this.spellchecker = new Typo('en_US', aff, dic, { platform: 'web' })
    this.ready = true
  }

  // Whole words only; non-ASCII words are outside the en_US dictionary's
  // judgment, and suggestions are looked up lazily because Typo.suggest is
  // orders of magnitude slower than Typo.check.
  check(text: string): Misspelling[] {
    if (!this.ready || !this.spellchecker) return []
    const words = text.match(/\p{L}[\p{L}''-]*/gu) ?? []
    const unique = new Set<string>()
    for (const raw of words) {
      if (unique.size >= MAX_UNIQUE_WORDS) break
      const word = raw.toLowerCase().replace(/['']/g, "'")
      if (word.length < 2 || !/^[a-z'-]+$/.test(word)) continue
      unique.add(word)
    }
    if (this.verdicts.size > VERDICT_CACHE_CAP) this.verdicts.clear()
    const misspelled: Misspelling[] = []
    for (const word of unique) {
      let ok = this.verdicts.get(word)
      if (ok === undefined) {
        ok = this.spellchecker.check(word)
        this.verdicts.set(word, ok)
      }
      if (!ok) misspelled.push({ word })
    }
    return misspelled
  }

  suggest(word: string): string[] {
    if (!this.ready || !this.spellchecker) return []
    return (this.spellchecker.suggest(word) as string[]) ?? []
  }
}
