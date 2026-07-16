import type { Cell } from '../types/cells'
import type { CreateDataValidation, DataValidation, DataValidations } from '../types/data-validation'
import { NS, TAGS_ORDER } from './constants'
import { isInRange, parseCellRange, toCellRange } from './excel-helpers'
import { getFormulasFromReferences } from './get-formulas'

const createDataValidation = (doc: Document, dataValidation: CreateDataValidation): Element => {
  if (dataValidation.sqref === undefined) throw new Error('The validation area is not defined')
  if (dataValidation.formula === undefined) throw new Error('The data validation formula is not defined')
  const dataValidationTag = doc.createElementNS(NS, 'dataValidation')
  const formulaTag = doc.createElementNS(NS, 'formula1')
  dataValidationTag.setAttribute('allowBlank', '1')
  dataValidationTag.setAttribute('showInputMessage', '1')
  dataValidationTag.setAttribute('showErrorMessage', '1')
  dataValidationTag.setAttribute('sqref', dataValidation.sqref)
  dataValidationTag.setAttribute('type', dataValidation.type)
  formulaTag.textContent = dataValidation.formula
  dataValidationTag.appendChild(formulaTag)
  return dataValidationTag
}

export const getDataValidations = (sheets: Array<{
  xml: Document
  name: string
}>): DataValidations => {
  const dataValidations = new Map<string, DataValidation>()

  let index = -1
  for (const sheet of sheets) {
    const dataValidationsParent = sheet.xml.querySelector('dataValidations')
    if (dataValidationsParent === null) continue

    const dataValidationsXml = dataValidationsParent.querySelectorAll('dataValidation')
    for (const dataValidation of dataValidationsXml) {
      const sheetName = sheet.name
      const extension = { rows: 0, cols: 0 }
      const formulaExtension = { rows: 0, cols: 0 }
      const sqref = dataValidation.getAttribute('sqref')
      const formula = dataValidation.querySelector('formula1')?.textContent
      if (sqref === null) continue
      index++
      const range = parseCellRange(sqref)
      const validation: Partial<DataValidation> = {
        sheet: sheetName,
        sqref,
        range,
        xml: dataValidation
      }
      if (formula !== undefined) {
        validation.formula = formula
        const references = getFormulasFromReferences(formula)
        if (references.length !== 1) continue
        const reference = references[0]
        if (reference.ref === undefined) continue
        validation.listRange = {
          ref: reference.ref,
          sheet: reference.sheet ?? sheetName
        }
        if (reference.ref === formula) {
          validation.listRange.range = parseCellRange(reference.ref)
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
      validation.extension = extension
      validation.formulaExtension = formulaExtension
      validation.extend = extend
      validation.extendFormula = extendFormula
      dataValidations.set(String(index), validation as DataValidation)
    }
  }

  const save = (): void => {
    for (const dataValidation of dataValidations.values()) {
      dataValidation.range.columnEnd = dataValidation.range.columnEnd + dataValidation.extension.cols
      dataValidation.range.rowEnd = dataValidation.range.rowEnd + dataValidation.extension.rows
      const ref = toCellRange(dataValidation.range)
      dataValidation.xml.setAttribute('sqref', ref)
      if (dataValidation.listRange?.range !== undefined) {
        dataValidation.listRange.range.columnEnd = dataValidation.listRange.range.columnEnd +
          dataValidation.formulaExtension.cols
        dataValidation.listRange.range.rowEnd = dataValidation.listRange.range.rowEnd +
          dataValidation.formulaExtension.rows
        const listRef = toCellRange(dataValidation.listRange.range)
        const formula = dataValidation.xml.querySelector('formula1')
        if (formula !== null) {
          formula.textContent = listRef
        }
      }
    }
  }

  const add = (xml: Document, dataValidation: CreateDataValidation): void => {
    if (dataValidation.sqref === undefined && dataValidation.range === undefined) {
      throw new Error('The validation area is not defined')
    }
    if (dataValidation.type === 'formula' && dataValidation.formula === undefined) {
      throw new Error('Data validation formula is not defined')
    }
    if (dataValidation.type === 'list' &&
      dataValidation.formula === undefined &&
      dataValidation.range === undefined
    ) {
      throw new Error('Data validation list is not defined')
    }
    if (dataValidation.range !== undefined) {
      dataValidation.sqref = toCellRange(dataValidation.range)
    }
    if (dataValidation.formula === undefined) {
      if (dataValidation.listRange?.range === undefined) {
        throw new Error('The data validation list is not defined')
      }
      dataValidation.formula = toCellRange(dataValidation.listRange.range)
    }
    const dataValidationXml = createDataValidation(sheets[0].xml, dataValidation)
    const dataValidationsTag = xml.querySelector('dataValidations')
    if (dataValidationsTag === null) {
      const elements = TAGS_ORDER.slice(TAGS_ORDER.indexOf('dataValidations') + 1)
      let element: Element | null = null
      for (const elName of elements) {
        const foundElement = xml.querySelector(elName)
        if (foundElement !== null) {
          element = foundElement
          break
        }
      }
      xml.querySelector('worksheet')?.insertBefore(xml.createElementNS(NS, 'dataValidations'), element)
    }
    xml.querySelector('dataValidations')?.appendChild(dataValidationXml)
  }

  const get = (ref: string): DataValidation | undefined => {
    return dataValidations.get(ref)
  }

  const findByCell = (sheetName: string, cell: Cell): string | null => {
    for (const [name, value] of dataValidations) {
      if (sheetName !== value.sheet) continue
      if (isInRange(cell, value.range)) return name
    }
    return null
  }

  const findListByCell = (sheetName: string, cell: Cell): string | null => {
    for (const [name, value] of dataValidations) {
      if (value.listRange?.range === undefined) continue
      if (value.listRange.sheet !== sheetName) continue
      if (isInRange(cell, value.listRange.range)) return name
    }
    return null
  }

  return {
    add,
    save,
    get,
    findByCell,
    findListByCell
  }
}
