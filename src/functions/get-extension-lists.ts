import type { Cell } from '../types/cells'
import type { ExtenstionList, ExtenstionLists } from '../types/extension-lists'
import { NS } from './constants'
import { isInRange, parseCellRange, toCellRange } from './excel-helpers'
import { getFormulasFromReferences } from './get-formulas'

export const getExtensionLists = (sheets: Array<{
  xml: Document
  name: string
}>): ExtenstionLists => {
  const extensions = new Map<string, ExtenstionList>()

  let index = -1
  for (const sheet of sheets) {
    const extenstionListsParent = sheet.xml.querySelector('extLst dataValidations') ??
      sheet.xml.createElementNS(NS, 'dataValidations')

    const extenstionListsXml = extenstionListsParent.querySelectorAll('dataValidation')

    for (const extensionList of extenstionListsXml) {
      const sheetName = sheet.name
      const extension = {
        rows: 0,
        cols: 0
      }
      const formulaExtension = {
        rows: 0,
        cols: 0
      }

      const sqref = extensionList.querySelector('sqref')?.textContent
      const formula = extensionList.querySelector('formula1 f')?.textContent
      if (sqref === undefined || formula === undefined) continue
      index++

      const range = parseCellRange(sqref)

      const currentExtension: Partial<ExtenstionList> = {
        xml: extensionList,
        sheet: sheetName,
        sqref,
        range
      }

      if (formula !== undefined) {
        currentExtension.formula = formula
        const references = getFormulasFromReferences(formula)
        if (references.length !== 1) continue
        const reference = references[0]
        if (reference.ref === undefined) continue
        currentExtension.listRange = {
          ref: reference.ref,
          sheet: reference.sheet ?? sheetName
        }
        if (`${String(reference.sheet)}!${String(reference.ref)}` === formula) {
          currentExtension.listRange.range = parseCellRange(reference.ref)
        }
      }

      const extend = (cell: Cell): void => {
        if (Array.isArray(cell.newValue)) {
          const length = cell.newValue.length - 1
          if ('$isTable' in cell.newValue) {
            extension.rows = Math.max(length, extension.rows)
          } else {
            extension.cols = Math.max(length, extension.cols)
          }
        }
      }

      const extendFormula = (cell: Cell): void => {
        if (Array.isArray(cell.newValue)) {
          const length = cell.newValue.length - 1
          if ('$isTable' in cell.newValue) {
            formulaExtension.rows = Math.max(length, formulaExtension.rows)
          } else {
            formulaExtension.cols = Math.max(length, formulaExtension.cols)
          }
        }
      }

      currentExtension.extension = extension
      currentExtension.formulaExtension = formulaExtension
      currentExtension.extend = extend
      currentExtension.extendFormula = extendFormula

      extensions.set(String(index), currentExtension as ExtenstionList)
    }
  }

  const save = (): void => {
    for (const extension of extensions.values()) {
      extension.range.columnEnd = extension.range.columnEnd + extension.extension.cols
      extension.range.rowEnd = extension.range.rowEnd + extension.extension.rows
      const ref = toCellRange(extension.range)
      const sqref = extension.xml.querySelector('sqref')
      if (sqref !== null) sqref.textContent = ref
      extension.xml.setAttribute('sqref', ref)
      if (extension.listRange?.range !== undefined) {
        extension.listRange.range.columnEnd = extension.listRange.range.columnEnd +
          extension.formulaExtension.cols
        extension.listRange.range.rowEnd = extension.listRange.range.rowEnd +
          extension.formulaExtension.rows
        const listRef = toCellRange(extension.listRange.range)
        const formula = extension.xml.querySelector('formula1 f')
        if (formula !== null) {
          formula.textContent = String(extension.listRange.sheet) + '!' + listRef
        }
      }
    }
  }

  const get = (ref: string): ExtenstionList | undefined => {
    return extensions.get(ref)
  }

  const findByCell = (cell: Cell): string | null => {
    for (const [name, value] of extensions) {
      if (isInRange(cell, value.range)) return name
    }
    return null
  }

  const findListByCell = (sheetName: string, cell: Cell): string | null => {
    for (const [name, value] of extensions) {
      if (value.listRange?.range === undefined) continue
      if (value.listRange.sheet !== sheetName) continue
      if (isInRange(cell, value.listRange.range)) return name
    }
    return null
  }

  return {
    save,
    get,
    findByCell,
    findListByCell
  }
}
