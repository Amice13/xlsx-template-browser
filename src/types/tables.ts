import { type Cell, type RangeRef } from './cells'

export interface Table {
  xml: Document
  ref: string
  extend: (cell: Cell) => Cell | undefined
  extension: { cols: number, rows: number }
  lastHeaderCell?: Cell
  range: RangeRef
  isDirty: boolean
  finalize: () => void
  save: () => void
}

export interface Tables {
  get: (s: string) => Table | undefined
  save: () => void
  findByCell: (cell: Cell) => string | null
}
