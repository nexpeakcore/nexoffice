export interface PageImage {
  dataUrl: string
  width: number
  height: number
}

export interface PageSet {
  pages: PageImage[]
  pageWidth: number
  pageHeight: number
  padding: number
  truncated: boolean
}
