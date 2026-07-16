import type { RangeRef, Cell } from './cells'

export interface ConditionalFormatting {
  sqref: string
  formula?: string
  extend: (cell: Cell) => void
  extension: { cols: number, rows: number }
  xml: Element
  range: RangeRef
}

export interface ConditionalFormattings {
  get: (s: string) => ConditionalFormatting | undefined
  save: () => void
  findByCell: (cell: Cell) => string | null
}
