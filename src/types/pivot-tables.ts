import { type Cell, type RangeRef } from './cells'

export interface PivotTable {
  definitionXml: Document
  cacheXml: Document
  sheet: string
  ref: string
  range: RangeRef
  extend: (cell: Cell) => void
  extension: { cols: number, rows: number }
  save: () => void
}

export interface PivotTables {
  get: (s: string) => PivotTable | undefined
  save: () => void
  findByCell: (sheet: string, cell: Cell) => string | null
}
