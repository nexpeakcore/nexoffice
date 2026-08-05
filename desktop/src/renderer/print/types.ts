export interface PageImage {
  dataUrl: string
  width: number
  height: number
}

export interface PageSet {
  pages: PageImage[]
  padding: number
  truncated: boolean
  // 1-based numbers of the source pages the renderer could not produce at all.
  // A page missing because it failed is a different outcome from a page missing
  // because the export stopped at the page cap, and the two are never merged.
  skippedPages: number[]
}
