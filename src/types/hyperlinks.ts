import type { CellRef } from './cells'

export interface Hyperlink {
  ref: string
  'r:id': string
  'xr:uid': string
  xml: Element
}

export interface Hyperlinks {
  add: ({ range, url }: { range: CellRef, url: string }) => void
  save: () => void
}
