import type JSZip from 'jszip'
import type { SharedStrings } from '../types/shared-strings'

import { serializeXml, createSharedString } from './xml-helpers'
import { parser } from './global-helpers'

export const getSharedStrings = async (
  xlsx: JSZip,
  target: string
): Promise<SharedStrings> => {
  let isDirty = false
  const filename = `xl/${target}`
  // Read all shared text values
  const xmlText = await xlsx.file(filename)?.async('string')
  if (xmlText === undefined) { throw new Error('This Excel template has no strings') }

  // Get the xml with shared shared strings
  const xml = parser.parseFromString(xmlText, 'application/xml')

  // Get shared strings element
  const sst = xml.querySelector('sst')
  if (sst === null) throw new Error('This Excel template has no strings')

  const oldValues = Array.from(sst.children)

  // Create new shared strings
  const strings = new Map<string, string>()
  const elements = new WeakMap<Element, string>()
  const list: Element[] = []

  // Create new shared strings
  const get = (el: string | Element): string => {
    if (typeof el === 'string') {
      const value = strings.get(el)
      if (value !== undefined) { return value }
      const newString = createSharedString(el, xml)
      const index = String(list.length)
      strings.set(el, index)
      list.push(newString)
      isDirty = true
      return index
    }
    const value = elements.get(el)
    if (value !== undefined) { return value }
    const index = String(list.length)
    elements.set(el, index)
    list.push(el)
    isDirty = true
    return index
  }

  const save = (): void => {
    if (!isDirty) return
    sst?.replaceChildren(...list)
    const text = serializeXml(xml)
    xlsx.file(filename, text)
  }

  return {
    oldValues,
    get,
    save
  }
}
