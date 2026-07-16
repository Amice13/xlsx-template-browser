import type { Cell } from '../types/cells'
import { parseCellRef, getExcelColumnIndex, getExcelColumnName } from './excel-helpers'

interface ShiftMap {
  rows: Map<number, number>
  cols: Map<number, number>
}

const typesDict: Record<string, Cell['type']> = {
  f: 'formula',
  e: 'error',
  b: 'boolean',
  s: 'string',
  n: 'number'
}

export const cellToModel = (cell: Element): Cell => {
  const currentType = cell.getAttribute('t')
  const cellType = typesDict[currentType as keyof typeof typesDict ?? 'n']
  const cellRef = cell.getAttribute('r')
  const cellStyle = cell.getAttribute('s')
  let value: string | number | undefined = cell.querySelector('v')?.textContent
  if (cellType === 'number' && typeof value === 'string') value = parseFloat(value)
  const formula = cell.querySelector('f')?.textContent
  if (cellRef === null) { throw new Error('Excel template is mailformed') }
  const { row, column } = parseCellRef(cellRef)
  return {
    ref: cellRef,
    row,
    column,
    style: cellStyle,
    type: cellType,
    ...(formula !== undefined ? { formula } : {}),
    ...(value !== undefined ? { value } : {})
  }
}

interface ShiftMap {
  rows: Map<number, number>
  cols: Map<number, number>
}

export const shiftCellRef = (ref: string, shift: ShiftMap): string => {
  const match = ref.match(/^([A-Z]+)(\d+)$/)
  if (match === null) return ref

  const [, colLetters, rowStr] = match
  const row = Number(rowStr)

  const column = getExcelColumnIndex(colLetters)

  const newRow = shift.rows.get(row)
  const newColumn = shift.cols.get(column)
  if (newRow === undefined || newColumn === undefined) throw new Error('Cell is not defined')

  return getExcelColumnName(newColumn) + String(newRow)
}

export const shiftRange = (ref: string, shift: ShiftMap): string => {
  const parts = ref.split(':')

  if (parts.length === 1) {
    return shiftCellRef(parts[0], shift)
  }

  const [start, end] = parts
  return `${shiftCellRef(start, shift)}:${shiftCellRef(end, shift)}`
}

export const updateTables = (tableXml: Document, shift: ShiftMap): void => {
  const table = tableXml.querySelector('table')
  if (table === null) return

  const ref = table.getAttribute('ref')
  if (typeof ref === 'string') {
    table.setAttribute('ref', shiftRange(ref, shift))
  }

  const autoFilter = table.querySelector('autoFilter')
  const afRef = autoFilter?.getAttribute('ref')

  if (typeof afRef === 'string') {
    autoFilter?.setAttribute('ref', shiftRange(afRef, shift))
  }
}
