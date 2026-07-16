import JSZip from 'jszip'
import { type Cell } from '../types/cells'
import { NS } from './constants'
import { toCellRef } from './excel-helpers'

export const readXlsx = async (buffer: ArrayBuffer): Promise<JSZip> => {
  const zip = new JSZip()
  const content = await zip.loadAsync(buffer)
  return content
}

export const createXml = ({
  coreSchema,
  specificSchema,
  tagName
}: {
  coreSchema?: string
  specificSchema: string
  tagName: string
}): XMLDocument => {
  const doc = document.implementation.createDocument(specificSchema, tagName, null)
  const root = doc.documentElement
  root.setAttributeNS('http://www.w3.org/2000/xmlns/', 'xmlns', specificSchema)
  if (coreSchema !== undefined) {
    root.setAttributeNS('http://www.w3.org/2000/xmlns/', 'xmlns:x', coreSchema)
  }
  return doc
}

const serializer = new XMLSerializer()
export const serializeXml = (doc: Document): string => {
  let serialized = serializer.serializeToString(doc)
  if (!serialized.startsWith('<?xml')) {
    serialized = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + serialized
  }
  return serialized
}

export const valueToString = (value: unknown): string => {
  if (value === undefined || value === null) { return '' }
  if (Array.isArray(value)) { return value.map(valueToString).toString() }
  if (typeof value === 'object') {
    if ('$cellType' in value) { return String(value) }
    return JSON.stringify(value)
  }
  if (typeof value === 'number') { return value.toString() }
  return String(value)
}

/**
 * Creates an Excel shared string (`<si>`) element for a given text value.
 *
 * - Wraps the text in `<si><t>...</t></si>` using the provided XML document
 * - Adds `xml:space="preserve"` when leading/trailing whitespace is present
 *   (required by Excel to avoid trimming)
 * - Uses the sharedStrings namespace (NS)
 *
 * @param text - Text content of the shared string
 * @param doc - XML document used to create elements
 * @returns `<si>` element ready to be inserted into sharedStrings.xml
 */

export const createSharedString = (
  text: string,
  doc: Document = document
): Element => {
  const si = doc.createElementNS(NS, 'si')
  const t = doc.createElementNS(NS, 't')

  // Excel requires xml:space="preserve" when whitespace matters
  if (/^\s|\s$/.test(text)) {
    t.setAttributeNS(
      'http://www.w3.org/XML/1998/namespace',
      'xml:space',
      'preserve'
    )
  }

  t.textContent = text
  si.append(t)
  return si
}

/**
 * Creates an Excel `<c>` (cell) XML element.
 *
 * @param params - Cell definition + XML document
 * @returns `<c>` element ready to append to worksheet
 */

const toType = {
  null: 's',
  error: 'e',
  number: null,
  boolean: 'b',
  string: 's',
  date: null,
  formula: null,
  url: 's'
}

export const createCell = ({
  value,
  type,
  style,
  row,
  column,
  formula,
  xml
}: Cell & { xml: Document }): Element => {
  const c = xml.createElementNS(NS, 'c')
  if (style !== null) c.setAttribute('s', style)
  const ref = toCellRef({ row, column })
  c.setAttribute('r', ref)
  if (type !== null) {
    const excelType = toType[type]
    if (excelType !== undefined && excelType !== null) c.setAttribute('t', excelType)
  }

  if (formula !== null && formula !== undefined) {
    const f = xml.createElementNS(NS, 'f')
    f.textContent = formula
    c.append(f)
  }

  if (value !== null && value !== undefined) {
    const v = xml.createElementNS(NS, 'v')
    v.textContent = String(value)
    c.append(v)
  }

  return c
}
