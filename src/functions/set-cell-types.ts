/* eslint-disable no-new-wrappers */
type CellTypes = 'formula' | 'number' | 'date' | 'string'

type EnrichedString = string & { $cellType: CellTypes }

export const isFormula = (s: string): EnrichedString => {
  if (typeof s !== 'string') { throw new Error('This is not a string') }
  const wrapped = new String(s) as EnrichedString
  Object.defineProperty(wrapped, '$cellType', { value: 'formula', enumerable: false })
  return wrapped
}

export const isBoolean = (s: string | number): EnrichedString => {
  if (!['string', 'number'].includes(typeof s)) { throw new Error('This is not a string or a number') }
  const wrapped = new String(s) as EnrichedString
  Object.defineProperty(wrapped, '$cellType', { value: 'string', enumerable: false })
  return wrapped
}

export const isString = (s: string | number): EnrichedString => {
  if (!['string', 'number'].includes(typeof s)) { throw new Error('This is not a string or a number') }
  const wrapped = new String(s) as EnrichedString
  Object.defineProperty(wrapped, '$cellType', { value: 'string', enumerable: false })
  return wrapped
}

export const isDate = (s: unknown): EnrichedString => {
  if (!(
    typeof s === 'string' ||
    typeof s === 'number' ||
    s instanceof Date
  )) { throw new Error('This is not a date') }
  const wrapped = new String(s) as EnrichedString
  Object.defineProperty(wrapped, '$cellType', { value: 'date', enumerable: false })
  return wrapped
}

export const isNumber = (s: string): EnrichedString => {
  if (!['string', 'number'].includes(typeof s)) { throw new Error('This is not a string or a number') }
  const wrapped = new String(s) as EnrichedString
  Object.defineProperty(wrapped, '$cellType', { value: 'number', enumerable: false })
  return wrapped
}
/* eslint-disable no-new-wrappers */
