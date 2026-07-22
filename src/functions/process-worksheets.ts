import { type Cell } from '../types/cells'
import type { AddComment } from '../types/comments'
import { type DataStore } from '../types/data-store'
import type { CreateDataValidation } from '../types/data-validation'
import { type Workbook } from '../types/workbook'
import { NS } from './constants'
import { isElement } from './data-helpers'
import { convertValueToExcel, guessDataType } from './excel-helpers'
import { getColumnShift, getOffsets, getRowShift } from './get-offsets'
import { getReplacements } from './get-replacements'
import { createCell } from './xml-helpers'

interface ProcessWorkSheets {
  workbook: Workbook
  dataStore: DataStore
}

export const processWorksheets = async ({
  workbook,
  dataStore
}: ProcessWorkSheets): Promise<void> => {
  const {
    styles,
    sharedStrings,
    dataValidations,
    pivotTables,
    extensionLists
  } = workbook

  const oldValues = sharedStrings.oldValues
  const replacements = getReplacements(oldValues, dataStore)

  for (const worksheet of workbook.sheets) {
    const {
      name,
      cells,
      mergedCells,
      hyperlinks,
      comments,
      conditionalFormattings,
      tables,
      xml,
      replaceCells
    } = worksheet
    comments.changeComments(dataStore)

    // Fill cells with new data
    const processedCells = cells
      .filter((cell: Cell) => {
        if (cell.inMergedCell !== undefined) {
          if ((cell.value === undefined || cell.value === null) &&
            (cell.formula === undefined || cell.formula === null)
          ) return false
        }
        return true
      })
      .map((cell: Cell) => {
        if (cell.type !== 'string') return cell
        if (cell.value === undefined) return cell
        if (cell.value === null) return cell
        cell.newValue = replacements.get(String(cell.value))
        if (cell.inTable !== undefined) {
          const table = tables.get(cell.inTable)
          table?.extend(cell)
        }

        if (cell.inConditionalFormatting !== undefined) {
          const formatting = conditionalFormattings.get(cell.inConditionalFormatting)
          formatting?.extend(cell)
        }
        if (cell.inDataValidation !== undefined) {
          const dataValidation = dataValidations.get(cell.inDataValidation)
          dataValidation?.extend(cell)
        }
        if (cell.inDataValidationList !== undefined) {
          const dataValidation = dataValidations.get(cell.inDataValidationList)
          dataValidation?.extendFormula(cell)
        }
        if (cell.inExtensionList !== undefined) {
          const extension = extensionLists.get(cell.inExtensionList)
          extension?.extend(cell)
        }
        if (cell.inExtensionListSource !== undefined) {
          const extension = extensionLists.get(cell.inExtensionListSource)
          extension?.extendFormula(cell)
        }
        if (cell.inPivotTableSource !== undefined) {
          const pivotTableSource = pivotTables.get(cell.inPivotTableSource)
          pivotTableSource?.extend(cell)
        }
        return cell
      })

    processedCells.sort((cell1: Cell, cell2: Cell) => {
      if (cell1.row !== cell2.row) return cell1.row - cell2.row
      return cell1.column - cell2.column
    })

    // Get offsets for new cells
    const offsets = getOffsets({
      cells: processedCells,
      mergedCells,
      tables
    })

    // Change cell location
    const movedCells = processedCells.map((cell: Cell) => {
      const colOffset = getColumnShift(cell, offsets)
      const rowOffset = getRowShift(cell, offsets)
      cell.column = cell.column + colOffset
      cell.row = cell.row + rowOffset
      return cell
    })

    // Extend cells
    const extendedCells: Cell[] = []
    for (const cell of movedCells) {
      if (cell.type !== 'string' || !Array.isArray(cell.newValue)) {
        extendedCells.push(cell)
        continue
      }
      if (Array.isArray(cell.newValue)) {
        const mergedCell = mergedCells.get(cell.inMergedCell ?? '')
        let width = 1
        let height = 1
        if (mergedCell !== undefined) {
          width = mergedCell.width
          height = mergedCell.height
        }
        for (let i = 0; i < cell.newValue.length; i++) {
          const newCell = { ...cell }
          newCell.newValue = cell.newValue[i]
          newCell.column = newCell.column + ('$isTable' in cell.newValue ? 0 : width) * i
          newCell.row = newCell.row + ('$isTable' in cell.newValue ? height : 0) * i
          if (mergedCell !== undefined) {
            mergedCells.add({
              columnStart: newCell.column,
              rowStart: newCell.row,
              columnEnd: newCell.column + width - 1,
              rowEnd: newCell.row + height - 1
            })
          }
          extendedCells.push(newCell)
        }
      }
    }

    const newCells = extendedCells.map((cell: Cell) => {
      if (isElement(cell.newValue)) {
        const sharedValue = sharedStrings.get(cell.newValue)
        cell.value = sharedValue
        return cell
      }
      if (cell.type !== 'string') {
        if (cell.type === 'formula') cell.value = null
        return cell
      }
      const newValue = cell.newValue === undefined ? cell.value : cell.newValue
      const cellType = cell.type === 'string' ? guessDataType(newValue) : cell.type
      const value = convertValueToExcel({ value: newValue, cellType })
      cell.type = cellType

      if ((cellType === 'formula' || cellType === 'error') && typeof value === 'string') {
        if (cell.formula === undefined) cell.formula = value
        cell.type = 'formula'
        cell.value = null
        return cell
      }
      if (cellType === 'string' && typeof value === 'string') {
        const sharedValue = sharedStrings.get(value)
        cell.value = sharedValue
        return cell
      }
      if (cellType === 'url' && typeof value === 'string') {
        cell.style = styles.getHyperlinkFormat()
        const sharedValue = sharedStrings.get(value)
        cell.value = sharedValue
        hyperlinks.add({ range: cell, url: value })
        return cell
      }
      if (cellType === 'date') {
        cell.style = styles.getDateFormat()
      }
      cell.value = value
      return cell
    })

    const rowsMap = new Map<number, Element>()
    for (const cell of newCells) {
      let row = rowsMap.get(cell.row)
      if (row === undefined) {
        row = xml.createElementNS(NS, 'row')
        row.setAttribute('r', String(cell.row))
        rowsMap.set(cell.row, row)
        row = rowsMap.get(cell.row)
      }
      row?.append(createCell({ ...cell, xml }))
    }
    const original = xml.querySelector('sheetData')

    if (original === null) throw new Error('Excel workbook is malformed')
    const newSheetData = original.cloneNode() as Element
    if (newSheetData === undefined) throw new Error('Excel workbook is mailformed')
    for (const row of rowsMap.values()) { newSheetData.append(row) }

    const validations = dataStore.get('$validations')
    if (Array.isArray(validations)) {
      for (const validation of (validations as CreateDataValidation[])) {
        if (validation.sheet !== name) continue
        dataValidations.add(xml, validation)
      }
    }
    const currentComments = dataStore.get('$comments')
    if (Array.isArray(currentComments)) {
      for (const comment of (currentComments as Array<AddComment & { sheet: string }>)) {
        if (comment.sheet !== name) continue
        comments.add(comment)
      }
    }
    replaceCells(newSheetData)
    comments.save()
    conditionalFormattings.save()
    tables.save()
    hyperlinks.save()
  }
  dataValidations.save()
  extensionLists.save()
  pivotTables.save()
  for (const worksheet of workbook.sheets) worksheet.save()
}
