import type { Cell, RangeRef } from './cells'

export interface MergedCell {
  xml: Element
  ref: string
  height: number
  width: number
  range: RangeRef
}

export interface MergedCells {
  add: (s: RangeRef) => void
  get: (s: string) => MergedCell | undefined
  findByCell: (cell: Cell) => string | null
}
