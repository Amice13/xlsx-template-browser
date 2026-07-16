import { type DataStore } from '../types/data-store'

const isNumeric = (v: string): boolean => {
  return /^\d+$/.test(v)
}

type Accessor = string | number

export const parseNotation = (input: string): Accessor[] => {
  const result: Accessor[] = []
  let i = 0

  while (i < input.length) {
    const char = input[i]

    // skip dots
    if (char === '.') {
      i++
      continue
    }

    // bracket notation
    if (char === '[') {
      i++

      const quote = input[i] === '"' || input[i] === '\'' ? input[i++] : null

      let value = ''

      while (i < input.length) {
        if (quote !== null) {
          if (input[i] === quote && input[i + 1] === ']') break
        } else {
          if (input[i] === ']') break
        }
        value += input[i++]
      }

      // skip closing quote + bracket OR just bracket
      i += quote !== null ? 2 : 1

      result.push(isNumeric(value) ? Number(value) : value)
      continue
    }

    // dot notation
    let value = ''
    while (i < input.length) {
      const c = input[i]
      if (c === '.' || c === '[') break
      value += c
      i++
    }

    result.push(value)
  }
  return result
}

/**
 * Retrieves nested properties from an object using an array of accessors.
 * @param {*} obj - The object from which to retrieve the nested properties.
 * @param {Array} accessors - An array of accessors representing the path to the desired properties.
 * @returns {*} - The value of the nested properties, or undefined if any accessor along the path is undefined.
 */

const getDeep = (obj: unknown, accessors: Array<string | number>): unknown => {
  for (let i = 0; i < accessors.length; i++) {
    if (obj === undefined) { return undefined }
    if (obj === null) { return obj }
    if (Array.isArray(obj) && typeof accessors[i] === 'string') {
      return obj.map(el => getDeep(el, accessors.slice(i)))
    }
    const accessor = accessors[i]
    if (accessor === undefined) { continue }
    obj = (obj as Record<string | number, unknown>)[accessor]
  }
  return obj
}

/**
 * Retrieves nested properties from an object using dot and bracket notation.
 * @param {*} obj - The object from which to retrieve the nested properties.
 * @param {string} props - The notation representing the path to the desired properties.
 * @returns {*} - The value of the nested properties, or the original object if the notation is empty.
 */

export const getByNotation = (obj: Record<string, unknown>, props: string): unknown => {
  const parsed = parseNotation(props)
  const value = getDeep(obj, parsed)
  return value ?? ''
}

/**
 * Creates a simple accessor for nested data using string paths (notation).
 *
 * - Resolves values via `getByNotation` (e.g. "user.address.city")
 * - Caches resolved values to avoid repeated lookups
 * - Accessor strings are trimmed before lookup
 *
 * @param data - Source object
 * @returns DataStore with cached `get` accessor
 */

export const createDataStore = (data: Record<string, unknown>): DataStore => {
  const accessors = new Map<string, unknown>()
  const get = (accessor: string): unknown => {
    accessor = accessor.trim()
    if (accessors.has(accessor)) { return accessors.get(accessor) }
    const value = getByNotation(data, accessor)
    accessors.set(accessor, value)
    return value
  }
  return { get }
}

export const isElement = (value: unknown): value is Element => {
  return (
    value instanceof Element ||
    (value instanceof Node && value.nodeType === Node.ELEMENT_NODE)
  )
}
