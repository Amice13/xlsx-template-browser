import type { DataStore } from '../types/data-store'
import { valueToString } from './xml-helpers'

// Regular expression for parsing a property accessor enclosed in ${}
const accessorRegex = /(?<=^\$\{)(?<table>\s*table:\s*)?(?<accessor>[^}]+?)(?=\s*}$)/

// Global regular expression for extracting all property accessors enclosed in ${}
const globalAccessorRegex = /\$\{(?<accessor>[^}]+)}/g

export const replaceAccessors = (text: string, dataStore: DataStore): string => {
  return text.replace(globalAccessorRegex, (_match, accessor) => {
    return valueToString(dataStore.get(accessor))
  })
}

export const getReplacements = (oldValues: Element[], dataStore: DataStore): Map<string, unknown> => {
  const replacements: Map<string, unknown> = new Map()

  let index = -1
  for (const si of oldValues) {
    index++
    if (si.children.length === 1 && si.children[0].tagName === 't') {
      const element = si.children[0]
      const text = element.textContent
      const match = text.match(accessorRegex)
      if (match === null) {
        const value = replaceAccessors(text, dataStore)
        replacements.set(String(index), value)
        continue
      }
      if (match.groups === undefined) {
        replacements.set(String(index), text)
        continue
      }
      const isTable = typeof match.groups.table === 'string'
      const value = dataStore.get(match.groups.accessor)
      if (typeof value !== 'object') {
        replacements.set(String(index), '')
        continue
      }
      const cloned = structuredClone(value)
      if (isTable) Object.defineProperty(cloned, '$isTable', { value: true, enumerable: false })
      replacements.set(String(index), cloned)
      continue
    }
    for (const t of Array.from(si.querySelectorAll('t'))) {
      const text = t.textContent ?? ''
      t.textContent = replaceAccessors(text, dataStore)
      replacements.set(String(index), si)
    }
  }
  return replacements
}
