import type { Cell } from '../types/cells'
import type { ConditionalFormatting, ConditionalFormattings } from '../types/conditional-formatting'
import { isInRange, parseCellRange, toCellRange } from './excel-helpers'

export const getConditionalFormatting = (xml: Document): ConditionalFormattings => {
  const conditionalFormattingXml = xml.querySelectorAll('conditionalFormatting')
  const conditionalFormattings: Map<string, ConditionalFormatting> = new Map()

  let index = -1
  for (const formatting of conditionalFormattingXml) {
    const sqref = formatting.getAttribute('sqref')
    const formula = formatting.querySelector('formula')?.textContent
    if (sqref === null) continue
    index++
    const range = parseCellRange(sqref)
    const extension = {
      rows: 0,
      cols: 0
    }

    let columnExtension = 0
    let rowExtension = 0
    let currentRow = 0
    const extend = (cell: Cell): void => {
      if (Array.isArray(cell.newValue)) {
        if (cell.row !== currentRow) {
          currentRow = cell.row
          columnExtension = 0
          rowExtension = 0
        }
        const length = cell.newValue.length - 1
        if ('$isTable' in cell.newValue) {
          rowExtension = Math.max(length, rowExtension) - Math.min(length, rowExtension)
          extension.rows = extension.rows + rowExtension
        } else {
          columnExtension = columnExtension + length
          extension.cols = Math.max(columnExtension, extension.cols)
        }
      }
    }

    const currentFormatting = {
      sqref,
      range,
      extension,
      extend,
      ...(formula === undefined ? {} : { formula }),
      xml: formatting
    }
    conditionalFormattings.set(String(index), currentFormatting)
  }

  const save = (): void => {
    for (const formatting of conditionalFormattings.values()) {
      formatting.range.columnEnd = formatting.range.columnEnd + formatting.extension.cols
      formatting.range.rowEnd = formatting.range.rowEnd + formatting.extension.rows
      const ref = toCellRange(formatting.range)
      formatting.xml.setAttribute('sqref', ref)
    }
  }

  const get = (ref: string): ConditionalFormatting | undefined => {
    return conditionalFormattings.get(ref)
  }

  const findByCell = (cell: Cell): string | null => {
    for (const [name, value] of conditionalFormattings) {
      if (isInRange(cell, value.range)) return name
    }
    return null
  }

  return {
    get,
    save,
    findByCell
  }
}
