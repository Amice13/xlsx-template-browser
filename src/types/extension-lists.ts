import type { RangeRef, Cell } from './cells'

export interface ExtenstionList {
  xml: Element
  sheet: string

  sqref: string
  range: RangeRef

  formula?: string
  listRange?: {
    sheet?: string
    ref: string
    range?: RangeRef
  }

  extend: (cell: Cell) => void
  extendFormula: (cell: Cell) => void

  extension: { cols: number, rows: number }
  formulaExtension: { cols: number, rows: number }
}

export interface ExtenstionLists {
  get: (s: string) => ExtenstionList | undefined
  save: () => void
  findByCell: (cell: Cell) => string | null
  findListByCell: (name: string, cell: Cell) => string | null
}
