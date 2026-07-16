import type { Cell } from '../types/cells'
import { NS } from './constants'
import { cellToModel } from './cells-helpers'

export const getCells = (xml: Document): Cell[] => {
  const sheetData = xml.querySelector('sheetData') ?? xml.createElementNS(NS, 'sheetData')
  const cellsXml = sheetData.querySelectorAll('c')
  const cells = cellsXml === undefined ? [] : Array.from(cellsXml).map(c => cellToModel(c))
  return cells
}
