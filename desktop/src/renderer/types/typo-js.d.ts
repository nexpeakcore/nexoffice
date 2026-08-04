declare module 'typo-js' {
  interface TypoSettings {
    platform?: string
    dictionaryPath?: string
    flags?: Record<string, unknown>
  }

  class Typo {
    constructor(
      dictionary?: string,
      affData?: string | null,
      wordsData?: string | null,
      settings?: TypoSettings,
    )
    check(word: string): boolean
    suggest(word: string, limit?: number): string[]
  }

  export default Typo
}

declare module '*.dic?url' {
  const src: string
  export default src
}

declare module '*.aff?url' {
  const src: string
  export default src
}
