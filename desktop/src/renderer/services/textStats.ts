// `page`/`pages` carry the slide position and slide count for a presentation,
// so one footer and one Word Count command serve every editor kind.
export interface EditorStats {
  words: number
  characters: number
  page: number
  pages: number
}

export function countWords(text: string): number {
  const matches = text.match(/[\p{L}\p{N}]+(?:[''-][\p{L}\p{N}]+)*/gu)
  return matches?.length ?? 0
}

export function countCharacters(text: string): number {
  return text.replace(/\s/g, '').length
}
