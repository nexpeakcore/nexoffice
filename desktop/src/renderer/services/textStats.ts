// `page`/`pages` carry the slide position and slide count for a presentation,
// so one footer and one Word Count command serve every editor kind.
export interface EditorStats {
  words: number
  characters: number
  page: number
  pages: number
}

// U+2019 counts alongside the ASCII apostrophe because Word and PowerPoint
// autocorrect straight quotes to it, so "don’t" is what a real document holds.
// Combining marks are part of the letter they sit on: Vietnamese and other
// heavily accented text arrives composed or decomposed depending on where it
// was typed, and a decomposed "Tiếng" must not count as two words.
const WORD = /[\p{L}\p{N}\p{M}]+(?:['’-][\p{L}\p{N}\p{M}]+)*/gu

export function countWords(text: string): number {
  return text.match(WORD)?.length ?? 0
}

export function countCharacters(text: string): number {
  return text.replace(/[\s\p{M}]/gu, '').length
}
