export interface PageImage {
  dataUrl: string
  width: number
  height: number
}

export interface PageSet {
  pages: PageImage[]
  padding: number
  truncated: boolean
}
