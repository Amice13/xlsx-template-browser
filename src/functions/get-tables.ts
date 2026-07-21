import type JSZip from 'jszip'
import type { Cell } from '../types/cells'
import type { Table, Tables } from '../types/tables'
import type { Relations } from '../types/relations'
import { isInRange, parseCellRange, toCellRange } from './excel-helpers'
import { generateUUID, parser } from './global-helpers'
import { NS } from './constants'

const serializer = new XMLSerializer()

export const getTables = async ({
  xlsx,
  relations
}: {
  xlsx: JSZip
  relations: Relations
}): Promise<Tables> => {
  const tablesProps = relations.getElements('table')
  const tableFiles = tablesProps.map(el => el.target)
  const tables = new Map<string, Table>()

  for (let file of tableFiles) {
    file = file.replace(/^\.\./, 'xl')
    const xmlText = await xlsx.file(file)?.async('string')
    if (xmlText === undefined) { continue }
    const xml = parser.parseFromString(xmlText, 'application/xml')
    const table = xml.querySelector('table')
    if (table === null) continue
    const ref = table.getAttribute('ref')
    if (ref === null) continue
    const range = parseCellRange(ref)
    const headers = Array.from(xml.querySelectorAll('tableColumn'))
      .map(el => el.getAttribute('name'))
      .filter(el => el !== null)
    const initialHeadersLength = headers.length
    const extension = {
      rows: 0,
      cols: 0
    }
    const existingHeaders: Set<string> = new Set(headers)

    const addHeader = (name: string, index: number): string => {
      let newName = name
      let j = 0
      while (existingHeaders.has(newName)) {
        j++
        newName = name + String(j)
      }
      existingHeaders.add(newName)
      headers.splice(index, 0, newName)
      return newName
    }

    const createHeaders = (names: string[]): Element[] => {
      return names.map((name, i) => {
        const h = xml.createElementNS(NS, 'tableColumn')
        h.setAttribute('id', String(i + 1))
        h.setAttribute('name', name)
        h.setAttribute('xr3:uid', `{${generateUUID().toUpperCase()}}`)
        return h
      })
    }

    let columnExtension = 0
    let rowExtension = 0
    let currentRow = 0
    const extend = (cell: Cell): Cell | undefined => {
      const t = tables.get(file)
      // Extend headers row
      if (t === undefined) return
      if (Array.isArray(cell.newValue)) {
        const isHeader = cell.row === range.rowStart
        if (isHeader) {
          const index = cell.column - range.columnStart
          const removed = headers.splice(index, 1)
          existingHeaders.delete(removed[0])
          cell.newValue = cell.newValue.map((header: string, i: number) => {
            return addHeader(header, i + index)
          })
        }
      }

      // Set in advance extensions for the table
      if (Array.isArray(cell.newValue)) {
        if (cell.row !== currentRow) {
          currentRow = cell.row
          columnExtension = 0
          rowExtension = 0
        }
        const length = cell.newValue.length - 1
        if ('$isTable' in cell.newValue) {
          rowExtension = Math.max(length, rowExtension) - Math.min(length, rowExtension)
          t.extension.rows = t.extension.rows + rowExtension
        } else {
          columnExtension = columnExtension + length
          t.extension.cols = Math.max(columnExtension, t.extension.cols)
        }
      }
      t.isDirty = true
    }

    const finalize = (): void => {
      const headersLength = range.columnEnd - range.columnStart + extension.cols + 1
      const headersDifference = headersLength - initialHeadersLength
      if (headersDifference !== 0) {
        for (let i = 0; i < headersDifference; i++) {
          addHeader('Column', range.columnEnd + i - 1)
        }
        if (tableInMap.lastHeaderCell !== undefined) {
          tableInMap.lastHeaderCell.newValue = headers.slice(initialHeadersLength - 1)
        }
      }
      const newHeaders = createHeaders(headers)
      xml.querySelector('tableColumns')?.replaceChildren(...newHeaders)
      range.columnEnd = range.columnEnd + extension.cols
      range.rowEnd = range.rowEnd + extension.rows
      const ref = toCellRange(range)
      table.setAttribute('ref', ref)
      const autofilter = table.querySelector('autoFilter')
      if (autofilter !== null) autofilter.setAttribute('ref', ref)
    }

    const save = (): void => {
      const newData = serializer.serializeToString(xml)
      xlsx.file(file, newData)
    }

    const tableInMap: Table = {
      xml,
      ref,
      range,
      extend,
      extension,
      finalize,
      save,
      isDirty: false
    }

    tables.set(file, tableInMap)
  }

  const get = (tableName: string): Table | undefined => {
    return tables.get(tableName)
  }

  const save = (): void => {
    for (const value of tables.values()) {
      if (value.isDirty) value.save()
    }
  }

  const findByCell = (cell: Cell): string | null => {
    for (const [name, table] of tables) {
      // Set the last header row to extend the table columns
      if (cell.column === table.range.columnEnd && cell.row === table.range.rowStart) {
        table.lastHeaderCell = cell
      }
      if (isInRange(cell, table.range)) return name
    }
    return null
  }

  return {
    get,
    save,
    findByCell
  }
}
