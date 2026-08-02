import Typo from 'typo-js'

export interface Misspelling {
  word: string
  suggestions: string[]
}

class SpellCheckServiceImpl {
  private spellchecker: Typo | null = null
  ready = false

  async init(_lang: string): Promise<void> {
    if (this.ready) return
    const [dic, aff] = await Promise.all([
      fetch('/dictionaries/en_US.dic').then((r) => r.text()),
      fetch('/dictionaries/en_US.aff').then((r) => r.text()),
    ])
    this.spellchecker = new Typo('en_US', aff, dic, { platform: 'web' })
    this.ready = true
  }

  check(text: string): Misspelling[] {
    if (!this.ready || !this.spellchecker) return []
    const words = text.match(/\b[a-zA-Z]{2,}\b/g) ?? []
    const unique = [...new Set(words.map((w) => w.toLowerCase()))]
    return unique
      .filter((w) => !this.spellchecker!.check(w))
      .map((word) => ({ word, suggestions: this.suggest(word) }))
  }

  suggest(word: string): string[] {
    if (!this.ready || !this.spellchecker) return []
    return (this.spellchecker.suggest(word) as string[]) ?? []
  }
}

export const spellCheckService = new SpellCheckServiceImpl()
