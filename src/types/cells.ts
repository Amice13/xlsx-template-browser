export type CellType = 'string' | 'number' | 'formula' | 'url' | 'date' | 'boolean' | 'error'

export interface CellRef {
  row: number
  column: number
  rowAbsolute?: boolean
  columnAbsolute?: boolean
}

export interface RangeRef {
  rowStart: number
  columnStart: number
  rowAbsoluteStart?: boolean
  columnAbsoluteStart?: boolean
  rowEnd: number
  columnEnd: number
  rowAbsoluteEnd?: boolean
  columnAbsoluteEnd?: boolean
}

export type Cell = {
  ref: string
  value?: string | number | null
  newValue?: unknown
  type: CellType | null
  style: string | null
  formula?: string | null
  inTable?: string
  inMergedCell?: string
  inNamedRange?: string
  inDataValidation?: string
  inDataValidationList?: string
  inExtensionList?: string
  inExtensionListSource?: string
  inPivotTableSource?: string
  inConditionalFormatting?: string
  hasComment?: string
  hasUrl?: boolean
} & CellRef
