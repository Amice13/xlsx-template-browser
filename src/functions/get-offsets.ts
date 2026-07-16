import type { Cell } from '../types/cells'
import type { MergedCells } from '../types/merged-cells'
import type { Tables } from '../types/tables'
import type { Offset } from '../types/offsets'

export function getOffsets ({
  cells,
  mergedCells,
  tables
}: {
  cells: Cell[]
  mergedCells: MergedCells
  tables: Tables
}): Offset {
  let rowOffset = 0
  let initialRowOffset = 0
  let columnOffset = 0
  let currentRow = -1

  const offsets: Offset = {
    row: {},
    col: {}
  }

  for (const cell of cells) {
    if (cell.inTable !== undefined) {
      const table = tables.get(cell.inTable)
      if (table === undefined) continue
      // Extend layout on the end of the header
      if (cell.column === table.range.columnEnd && cell.row === table.range.rowStart) {
        const tableHeight = table.range.rowEnd - table.range.rowStart + 1
        for (let i = 0; i < tableHeight; i++) {
          const row = cell.row + i
          offsets.col[row] ??= {}
          offsets.col[row][table.range.columnEnd] = table.extension.cols
        }
        table.finalize()
        continue
      }
    }

    if (!Array.isArray(cell.newValue)) continue
    if (currentRow !== cell.row) {
      initialRowOffset = rowOffset
      currentRow = cell.row
      columnOffset = 0
    }
    if (offsets.col[currentRow] !== undefined) {
      for (const column of Object.keys(offsets.col[currentRow])) {
        const index = parseInt(column)
        if (index < cell.column) columnOffset = columnOffset = Math.max(columnOffset, offsets.col[currentRow][index])
      }
    }
    let unit = 1
    let rows = 1

    if (cell.inMergedCell !== undefined) {
      const merge = mergedCells.get(cell.inMergedCell)
      if (merge !== undefined) {
        unit = '$isTable' in cell.newValue ? merge.height : merge.width
        rows = merge.height
      }
    }
    if ('$isTable' in cell.newValue) {
      const existingOffset = offsets.row[cell.row] ?? 0
      const newOffset = initialRowOffset + cell.newValue.length * unit - unit
      rowOffset = Math.max(newOffset, existingOffset ?? 0)
      offsets.row[cell.row] = rowOffset
      continue
    }
    const delta = cell.newValue.length * unit - unit
    columnOffset = columnOffset + delta
    for (let i = 0; i < rows; i++) {
      const row = cell.row + i
      offsets.col[row] ??= {}
      offsets.col[row][cell.column] = columnOffset
    }
  }
  return offsets
}

export const getRowShift = (cell: { row: number }, offsets: Offset): number => {
  let shift = 0
  for (const r in offsets.row) {
    const ri = Number(r)
    if (ri < cell.row) shift = offsets.row[ri]
  }
  return shift
}

export const getColumnShift = (cell: Cell, offsets: Offset): number => {
  const rowOffsets = offsets.col
  const colOffsets = rowOffsets[cell.row]
  if (colOffsets === undefined) return 0
  let shift = 0
  for (const c in colOffsets) {
    const ci = Number(c)
    if (ci < cell.column) shift = colOffsets[ci]
  }
  return shift
}
