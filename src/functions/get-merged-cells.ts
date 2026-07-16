import type { Cell, RangeRef } from '../types/cells'
import type { MergedCell, MergedCells } from '../types/merged-cells'
import { NS } from './constants'
import { isInRange, parseCellRange, toCellRange } from './excel-helpers'

export const getMergedCells = (xml: Document): MergedCells => {
  const mergedCells = new Map<string, MergedCell>()
  const mergeCellsParent = xml.querySelector('mergeCells') ?? xml.createElementNS(NS, 'mergeCells')
  const mergeCellsXml = mergeCellsParent.querySelectorAll('mergeCell')

  for (const mergedCell of Array.from(mergeCellsXml)) {
    const ref = mergedCell.getAttribute('ref')
    if (ref === null) continue
    const range = parseCellRange(ref)
    mergedCells.set(ref, {
      range,
      width: range.columnEnd - range.columnStart + 1,
      height: range.rowEnd - range.rowStart + 1,
      ref,
      xml: mergedCell
    })
  }

  mergeCellsParent.replaceChildren()
  const add = (ref: RangeRef): void => {
    const mergedCell = xml.createElementNS(NS, 'mergeCell')
    mergedCell.setAttribute('ref', toCellRange(ref))
    mergeCellsParent.appendChild(mergedCell)
  }

  const get = (ref: string): MergedCell | undefined => {
    return mergedCells.get(ref)
  }

  const findByCell = (cell: Cell): string | null => {
    for (const [name, value] of mergedCells) {
      if (isInRange(cell, value.range)) return name
    }
    return null
  }

  return {
    add,
    get,
    findByCell
  }
}
