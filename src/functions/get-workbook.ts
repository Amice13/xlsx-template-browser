import type JSZip from 'jszip'
import { parser } from './global-helpers'
import { GLOBAL_RELS, OFFICE_DOCUMENT_TYPE } from './constants'
import { getRelationsPath } from './excel-helpers'
import type { Sheet } from '../types/sheets'
import type { Workbook } from '../types/workbook'
import { createStyles } from './get-styles'
import { getSharedStrings } from './get-shared-strings'
import { getSheet } from './get-sheets'
import { getRelations } from './get-relations'
import { getPivotTables } from './get-pivot-tables'
import type { Relations } from '../types/relations'
import { getDataValidations } from './get-data-validations'
import { getExtensionLists } from './get-extension-lists'
import { getPersons } from './get-persons'

export const getWorkbook = async (xlsx: JSZip): Promise<Workbook> => {
  // Get global path for the workbook
  const relsXmlText = await xlsx.file(GLOBAL_RELS)?.async('string')
  if (relsXmlText === undefined) throw new Error('This Excel file has no relations')
  const relsXml = parser.parseFromString(relsXmlText, 'application/xml')
  const workBookElement = relsXml.querySelector(`Relationship[Type="${OFFICE_DOCUMENT_TYPE}/officeDocument"]`)
  if (workBookElement === null) throw new Error('This Excel file has no workbook')

  const workbookPath = workBookElement.getAttribute('Target')
  if (workbookPath === null) throw new Error('This Excel file has no workbook')

  // Process the workbook
  const workbookXmlText = await xlsx.file(workbookPath)?.async('string')
  if (workbookXmlText === undefined) throw new Error('This Excel file has no workbook')

  const workbookXml = parser.parseFromString(workbookXmlText, 'application/xml')
  const sheetsParent = workbookXml.querySelector('sheets')
  if (sheetsParent === null) throw new Error('This Excel file has no sheets')
  const sheetsXml = Array.from(sheetsParent.children)

  const sheets: Array<Partial<Sheet>> = sheetsXml.map(el => ({
    id: el.getAttribute('r:id') ?? undefined,
    name: el.getAttribute('name') ?? undefined,
    sheetId: el.getAttribute('sheetId') ?? undefined
  }))

  // Process relations
  const relationsPath = getRelationsPath(workbookPath)
  const relations = await getRelations({ xlsx, filename: relationsPath })

  const workbook: Partial<Workbook> = {}
  workbook.relations = relations

  // Get styles
  const styleProps = relations.getElements('styles')[0]
  if (styleProps === undefined) throw new Error('Styles are mailformed')
  const styles = await createStyles(xlsx, styleProps.target)
  workbook.styles = styles

  // Shared strings
  const sharedStringsProps = relations.getElements('sharedStrings')[0]
  if (sharedStringsProps === undefined) throw new Error('This file does not contain strings')
  const sharedStrings = await getSharedStrings(xlsx, sharedStringsProps.target)
  workbook.sharedStrings = sharedStrings

  // Get system user
  const persons = await getPersons(xlsx, 'persons/person.xml')
  workbook.persons = persons

  const pivotTables = await getPivotTables(xlsx)
  workbook.pivotTables = pivotTables

  const sheetsRels = relations.getElements('worksheet')

  const sheetsData = await Promise.all(sheetsRels.map(async (s): Promise<{
    xml: Document
    relations: Relations
    target: string
    id: string
    name: string
  }> => {
    if (s?.target === undefined || s?.id === undefined) throw new Error('Excel file is mailformed')
    const currentSheet = sheets.find(sheet => sheet.id === s.id)
    if (currentSheet === undefined || currentSheet.name === undefined) throw new Error('Excel file is mailformed')
    const relsTarget = s.target.replace(/worksheets\//, 'worksheets/_rels/') + '.rels'
    const relations = await getRelations({ xlsx, filename: relsTarget })
    const sheetXmlText = await xlsx.file(`xl/${s.target}`)?.async('string')
    if (sheetXmlText === undefined) throw new Error(`${s.target} sheet is mailformed`)
    const xml = parser.parseFromString(sheetXmlText, 'application/xml')
    const sheet = {
      xml,
      relations,
      target: s.target,
      id: s.id,
      name: currentSheet.name
    }
    Object.assign(currentSheet, sheet)
    return sheet
  }))

  const dataValidations = getDataValidations(sheetsData)
  workbook.dataValidations = dataValidations

  const extensionLists = getExtensionLists(sheetsData)
  workbook.extensionLists = extensionLists

  const enrichedSheets = await Promise.all(sheetsData.map(async sheet => await getSheet({
    xlsx,
    workbook: workbook as Workbook,
    ...sheet
  })))

  workbook.sheets = enrichedSheets as Sheet[] // eslint-disable-line

  workbook.save = (): void => {
    const calcChain = relations.get({ by: 'realType', value: 'calcChain' })
    for (const { target } of calcChain) {
      xlsx.remove(`xl/${target}`)
    }
    relations.remove({ by: 'realType', value: 'calcChain' })
    sharedStrings.save()
    styles.save()
  }

  return workbook as Workbook // eslint-disable-line
}
