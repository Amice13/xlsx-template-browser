import type { CellType, CellRef, RangeRef, Cell } from '../types/cells'

export const toCellRef = (ref: CellRef): string => {
  let col = ref.column

  if (col <= 0) throw new Error(`Invalid column: ${col}`)

  let column = ''
  while (col > 0) {
    const rem = (col - 1) % 26
    column = String.fromCharCode(65 + rem) + column
    col = Math.floor((col - 1) / 26)
  }

  const colPart = `${(ref.columnAbsolute === true ? '$' : '')}${column}`
  const rowPart = `${(ref.rowAbsolute === true ? '$' : '')}${ref.row}`

  if (ref.row <= 0) throw new Error(`Invalid row: ${ref.row}`)

  return colPart + rowPart
}

export const parseCellRef = (ref: string): CellRef => {
  const normalized = ref.trim().toUpperCase()

  let column = 0
  let i = 0

  let columnAbsolute = false
  let rowAbsolute = false

  // Check column $
  if (normalized[i] === '$') {
    columnAbsolute = true
    i++
  }

  // Parse column letters
  for (; i < normalized.length; i++) {
    const code = normalized.charCodeAt(i)
    if (code < 65 || code > 90) break
    column = column * 26 + (code - 64)
  }

  // Check row $
  if (normalized[i] === '$') {
    rowAbsolute = true
    i++
  }

  // Parse row digits
  const rowStr = normalized.slice(i)
  const row = Number(rowStr)

  if (column === 0 || rowStr === '' || Number.isNaN(row)) {
    throw new Error(`Invalid cell reference: ${ref}`)
  }

  return { row, column, rowAbsolute, columnAbsolute }
}

export const parseCellRange = (ref: string): RangeRef => {
  const parts = ref.split(':')
  if (parts.length === 0 || parts.length > 2) {
    throw new Error(`Invalid cell range: ${ref}`)
  }
  if (parts.length === 1) parts.push(parts[0])
  const [start, end] = parts
  if (start === undefined || end === undefined) throw new Error(`Cell ${ref} is broken`)
  const startRef = parseCellRef(start)
  const endRef = parseCellRef(end)
  return {
    rowStart: startRef.row,
    columnStart: startRef.column,
    rowAbsoluteStart: startRef.rowAbsolute,
    columnAbsoluteStart: startRef.columnAbsolute,
    rowEnd: endRef.row,
    columnEnd: endRef.column,
    rowAbsoluteEnd: endRef.rowAbsolute,
    columnAbsoluteEnd: endRef.columnAbsolute
  }
}

export const toCellRange = (range: Partial<RangeRef>): string => {
  if (range.rowStart === undefined && range.rowEnd === undefined) {
    if (range.columnStart === undefined || range.columnEnd === undefined) {
      throw new Error('Range is not defined')
    }
    return [
      range.columnAbsoluteStart === true ? '$' : '',
      String(range.columnStart),
      ':',
      range.columnAbsoluteEnd === true ? '$' : '',
      String(range.columnEnd)
    ].join('')
  }

  if (range.columnStart === undefined && range.columnEnd === undefined) {
    if (range.rowStart === undefined || range.rowEnd === undefined) {
      throw new Error('Range is not defined')
    }
    return [
      range.rowAbsoluteStart === true ? '$' : '',
      getExcelColumnName(range.rowStart),
      ':',
      range.rowAbsoluteEnd === true ? '$' : '',
      getExcelColumnName(range.rowEnd)
    ].join('')
  }

  if (
    range.rowStart === undefined ||
    range.columnStart === undefined ||
    range.rowEnd === undefined ||
    range.columnEnd === undefined
  ) {
    throw new Error('Range is not defined')
  }
  const start = toCellRef({
    row: range.rowStart,
    column: range.columnStart,
    rowAbsolute: range.rowAbsoluteStart,
    columnAbsolute: range.columnAbsoluteStart
  })
  if (range.rowEnd === 0 && range.columnEnd === 0) return start

  const end = toCellRef({
    row: range.rowEnd,
    column: range.columnEnd,
    rowAbsolute: range.rowAbsoluteEnd,
    columnAbsolute: range.columnAbsoluteEnd
  })

  return `${start}:${end}`
}

export const getExcelColumnIndex = (cellRef: string): number => {
  let index = 0

  for (let i = 0; i < cellRef.length; i++) {
    const code = cellRef.charCodeAt(i)

    // Stop when we hit a digit
    if (code < 65 || code > 90) { break }

    index = index * 26 + (code - 64)
  }

  return index
}

export const getExcelColumnName = (index: number): string => {
  let result = ''
  while (index > 0) {
    index--
    result = String.fromCharCode(65 + (index % 26)) + result
    index = Math.floor(index / 26)
  }
  return result
}

export const dateToExcel = (value: string | Date): string => {
  const date = new Date(value)
  const dateValue = 25_569 + ((date.getTime() - (date.getTimezoneOffset() * 60 * 1000)) / (1000 * 60 * 60 * 24))
  return String(dateValue)
}

/**
 * Converts a JavaScript value into an Excel-compatible cell value
 * based on the specified cell type.
 *
 * Notes:
 * - Performs loose coercion for objects via String() / JSON.stringify()
 * - Falls back to String(value) if no branch matches
 *
 * @param value - Input value of any type
 * @param cellType - Target Excel cell type
 * @returns Excel-compatible string value or null
 */

interface ConvertValueToExcel {
  value: unknown
  cellType: CellType | null
}

export const convertValueToExcel = ({ value, cellType }: ConvertValueToExcel): string | null => {
  if (cellType === null) { return null }
  switch (cellType) {
    case 'formula': {
      if (typeof value === 'object') { value = String(value) }
      if (typeof value === 'string') {
        return value.startsWith('=') ? value.slice(1) : value
      }

      break
    }
    case 'date': {
      if (value instanceof Date || typeof value === 'string') {
        return dateToExcel(value)
      }
      if (typeof value === 'object') {
        return dateToExcel(String(value))
      }

      break
    }
    case 'boolean': {
      if (typeof value === 'boolean') { return String(Number(value)) }
      if (typeof value === 'object') { return String(Number(/true/i.test(String(value)))) }

      break
    }
    case 'number': {
      if (typeof value === 'object') { return String(value) }
      if (typeof value === 'number') { return value.toString() }

      break
    }
    default: { if (typeof value === 'object') {
      return JSON.stringify(value)
    }
    }
  }
  return String(value)
}

/**
 * Guesses the most appropriate Excel data type for a given value.
 *
 * Notes:
 * - Invalid Date objects are treated as 'null'
 * - URL detection uses a permissive regex
 * - Objects default to string unless explicitly typed
 *
 * @param value - Any input value
 * @returns Inferred data type for Excel cell handling
 */

const urlPatten = String.raw`^(?!mailto:)(?:(?:http|https|ftp)://)(?:\S+(?::\S*)?@)?(?:(?:(?:[1-9]\d?|1\d\d|2[01]\d|22[0-3])(?:\.(?:1?\d{1,2}|2[0-4]\d|25[0-5])){2}(?:\.(?:[0-9]\d?|1\d\d|2[0-4]\d|25[0-4]))|(?:(?:[a-z\u00a1-\uffff0-9]+-?)*[a-z\u00a1-\uffff0-9]+)(?:\.(?:[a-z\u00a1-\uffff0-9]+-?)*[a-z\u00a1-\uffff0-9]+)*(?:\.(?:[a-z\u00a1-\uffff]{2,})))|localhost)(?::\d{2,5})?(?:(/|\?|#)[^\s]*)?$`
const urlRegex = new RegExp(urlPatten, 'i')

export const guessDataType = (value: unknown): CellType | null => {
  if (value === undefined || value === null || value === '') { return null }
  if (typeof value === 'number') { return 'number' }
  if (typeof value === 'boolean') { return 'boolean' }

  if (value instanceof Date) {
    return isNaN(value.getTime()) ? null : 'date'
  }

  if (typeof value === 'string') {
    if (value.startsWith('=')) return 'formula' // cheaper than regex
    if (urlRegex.test(value)) return 'url'
    return 'string'
  }

  if (typeof value === 'object') {
    if ('$cellType' in value) { return value.$cellType as CellType }
    return 'string'
  }

  return null
}

export const getRelationsPath = (partPath: string): string => {
  const slash = partPath.lastIndexOf('/')
  const file = partPath.slice(slash + 1)
  return `_rels/${file}.rels`
}

export const isInRange = (c: Cell, r: RangeRef): boolean =>
  c.column >= r.columnStart &&
  c.column <= r.columnEnd &&
  c.row >= r.rowStart &&
  c.row <= r.rowEnd

interface CompleteLocation {
  row: number
  column: number
  ref: string
}

export const completeLocation = (
  input: Partial<CompleteLocation>
): CompleteLocation => {
  let { row, column, ref } = input

  if (ref !== undefined) {
    const parsed = parseCellRef(ref)
    row ??= parsed.row
    column ??= parsed.column
  }

  if (row !== undefined && column !== undefined) {
    ref ??= toCellRef({ row, column })
  }

  if (row === undefined || column === undefined || ref === undefined) {
    throw new Error('Location is not defined')
  }

  return { row, column, ref }
}
