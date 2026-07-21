import JSZip from 'jszip'
import type { Cell } from '../types/cells'
import { parser } from './global-helpers'
import { NS_PREFIX } from './constants'
import { isInRange, parseCellRange, toCellRange } from './excel-helpers'
import type { PivotTable, PivotTables } from '../types/pivot-tables'

const serializer = new XMLSerializer()

export const getPivotTables = async (xlsx: JSZip): Promise<PivotTables> => {
  const pivotTables = new Map<string, PivotTable>()

  const pivotFiles = Object.keys(xlsx.files)
    .filter(f => f.startsWith('xl/pivotTables/') && f.endsWith('.xml'))
    .map(f => ({
      definitionFilename: f,
      relsFilename: f.replace('/pivotTables', '/pivotTables/_rels') + '.rels'
    }))

  for (const { definitionFilename, relsFilename } of pivotFiles) {
    // Get table
    const tableDefinitionText = await xlsx.file(definitionFilename)?.async('string')
    if (tableDefinitionText === undefined) continue

    const tableDefinitionXml = parser.parseFromString(tableDefinitionText, 'application/xml')
    const tableDefinition = tableDefinitionXml.querySelector('pivotTableDefinition')
    if (tableDefinition === null) continue

    // Get cache
    const xmlRelsText = await xlsx.file(relsFilename)?.async('string')
    if (xmlRelsText === undefined) continue
    const rels = parser.parseFromString(xmlRelsText, 'application/xml')
    const cacheTarget = rels
      .querySelector(`Relationship[Type="${NS_PREFIX}/relationships/pivotCacheDefinition"]`)
      ?.getAttribute('Target')
    if (cacheTarget === undefined || cacheTarget === null) continue
    const cacheFile = cacheTarget.replace('../', 'xl/')
    const cacheText = await xlsx.file(cacheFile)?.async('string')
    if (cacheText === undefined) continue
    const cache = parser.parseFromString(cacheText, 'application/xml')
    const workSheetSource = cache.querySelector('worksheetSource')
    if (workSheetSource === null) continue

    const ref = workSheetSource.getAttribute('ref')
    const sheet = workSheetSource.getAttribute('sheet')
    if (ref === null || sheet === null) continue
    const range = parseCellRange(ref)

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

    const save = (): void => {
      tableDefinition.setAttribute('refreshOnLoad', '1')
      tableDefinition.setAttribute('saveData', '0')
      const cacheDefinition = cache.querySelector('pivotCacheDefinition')
      if (cacheDefinition !== null) {
        cacheDefinition.setAttribute('refreshOnLoad', '1')
        cacheDefinition.setAttribute('saveData', '0')
        cacheDefinition.setAttribute('invalid', '1')
        cacheDefinition.setAttribute('recordCount', '1')
      }
      const newDefinition = serializer.serializeToString(tableDefinition)
      xlsx.file(definitionFilename, newDefinition)
      range.columnEnd = range.columnEnd + extension.cols
      range.rowEnd = range.rowEnd + extension.rows
      const ref = toCellRange(range)
      workSheetSource.setAttribute('ref', ref)
      const newCache = serializer.serializeToString(cache)
      xlsx.file(cacheFile, newCache)
    }

    const pivotTable: PivotTable = {
      ref,
      sheet,
      range,
      definitionXml: tableDefinitionXml,
      cacheXml: cache,
      extension,
      extend,
      save
    }
    pivotTables.set(definitionFilename, pivotTable)
  }

  const get = (tableName: string): PivotTable | undefined => {
    return pivotTables.get(tableName)
  }

  const save = (): void => {
    for (const table of pivotTables.values()) {
      table.save()
    }
  }

  const findByCell = (sheet: string, cell: Cell): string | null => {
    for (const [name, value] of pivotTables) {
      if (value.sheet !== sheet) continue
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
