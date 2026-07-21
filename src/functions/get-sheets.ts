import type JSZip from 'jszip'
import type { Workbook } from '../types/workbook'
import { getCells } from './get-cells'
import { getMergedCells } from './get-merged-cells'
import { getHyperlinks } from './get-hyperlinks'
import { getComments } from './get-comments'
import { getConditionalFormatting } from './get-conditional-formatting'
import { getTables } from './get-tables'
import { serializeXml } from './xml-helpers'
import type { Relations } from '../types/relations'
import { getDrawings } from './get-drawings'
import { getOldComments } from './get-old-comments'
import { type Sheet } from '../types/sheets'

export const getSheet = async ({
  xlsx,
  xml,
  relations,
  workbook,
  target,
  id,
  name
}: {
  xlsx: JSZip
  xml: Document
  relations: Relations
  workbook: Workbook
  target: string
  id: string
  name: string
}): Promise<Sheet> => {
  const filename = `xl/${target}`
  const {
    pivotTables,
    dataValidations,
    extensionLists
  } = workbook
  const initialCells = getCells(xml)
  const mergedCells = getMergedCells(xml)
  const hyperlinks = getHyperlinks({ xml, relations })
  const comments = await getComments({ xlsx, workbook, relations, id })
  const conditionalFormattings = getConditionalFormatting(xml)
  const tables = await getTables({ xlsx, relations })
  const drawings = await getDrawings({ xlsx, id, relations })
  const oldComments = await getOldComments({ xlsx, id, relations })
  comments.setDrawings(drawings)
  comments.setOldComments(oldComments)

  // TO DO
  // print areas
  // named ranges
  // drawings
  // charts
  // sparklines

  const cells = initialCells.map(cell => {
    const comment = comments.getComment(cell.ref)
    if (comment !== undefined) cell.hasComment = comment
    const table = tables.findByCell(cell)
    if (table !== null) cell.inTable = table
    const mergeCell = mergedCells.findByCell(cell)
    if (mergeCell !== null) cell.inMergedCell = mergeCell
    const pivotTable = pivotTables.findByCell(name, cell)
    if (pivotTable !== null) cell.inPivotTableSource = pivotTable
    const dataValidation = dataValidations.findByCell(name, cell)
    if (dataValidation !== null) cell.inDataValidation = dataValidation
    const dataValidationList = dataValidations.findListByCell(name, cell)
    if (dataValidationList !== null) cell.inDataValidationList = dataValidationList
    const conditionalFormatting = conditionalFormattings.findByCell(cell)
    if (conditionalFormatting !== null) cell.inConditionalFormatting = conditionalFormatting
    const extension = conditionalFormattings.findByCell(cell)
    if (extension !== null) cell.inExtensionList = extension
    const extensionListSource = extensionLists.findListByCell(name, cell)
    if (extensionListSource !== null) cell.inExtensionListSource = extensionListSource
    return cell
  })

  const replaceCells = (element: Element): void => {
    xml.querySelector('sheetData')?.replaceWith(element)
  }

  const cleanUp = (): void => {
    const hyperlinks = xml.querySelector('hyperlinks')
    if (hyperlinks?.childNodes.length === 0) hyperlinks.remove()
    const mergeCells = xml.querySelector('mergeCells')
    if (mergeCells?.childNodes.length === 0) mergeCells.remove()
  }

  const save = (): void => {
    cleanUp()
    const newData = serializeXml(xml)
    relations.save()
    xlsx.file(filename, newData)
  }

  return {
    xml,
    sheetId: id,
    relations,
    target,
    id,
    name,
    cells,
    mergedCells,
    hyperlinks,
    comments,
    dataValidations,
    extensionLists,
    conditionalFormattings,
    drawings,
    oldComments,
    tables,
    replaceCells,
    save
  }
}
