const sheetRegex = '(?:(?<sheet>(?:\'(?:[^\'!]|\'\')+\'|[^\'!()+/*-]+))!)'
const completeRefRegex = '\\$?[A-Z]+\\$?[0-9]+(?::\\$?[A-Z]+\\$?[0-9]+)?'
const partialRefRegex = '\\$?[A-Z]+(?::\\$?[A-Z]+)|\\$?[0-9]+(?::\\$?[0-9]+)'
const re = new RegExp(`(?<![A-Z0-9])(?:${sheetRegex})?(?<ref>${completeRefRegex}|${partialRefRegex})(?![A-Z0-9])`, 'g')

export const getFormulasFromReferences = (s: string): Array<{ sheet?: string, ref?: string }> =>
  [...s.matchAll(re)].map(m => ({
    sheet: m.groups?.sheet
      ?.replace(/^'|'$/g, '')
      .replace(/''/g, "'"),
    ref: m.groups?.ref
  }))
