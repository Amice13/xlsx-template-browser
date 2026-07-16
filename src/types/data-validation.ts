import type { RangeRef, Cell } from './cells'

export interface DataValidation {
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
  extension: { cols: number, rows: number }
  extendFormula: (cell: Cell) => void
  formulaExtension: { cols: number, rows: number }
}

export interface DataValidations {
  add: (xml: Document, dataValidation: CreateDataValidation) => void
  get: (s: string) => DataValidation | undefined
  save: () => void
  findByCell: (name: string, cell: Cell) => string | null
  findListByCell: (name: string, cell: Cell) => string | null
}

export interface CreateDataValidation {
  type: 'list' | 'formula'
  sheet?: string
  sqref?: string
  range?: RangeRef
  formula?: string
  listRange?: {
    sheet?: string
    ref?: string
    range?: RangeRef
  }
}
