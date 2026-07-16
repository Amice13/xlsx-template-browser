import type JSZip from 'jszip'
import { type Styles } from '../types/styles'

import { NS } from './constants'
import { parser } from './global-helpers'

const DATE_FORMAT = String.raw`yyyy\-mm\-dd;@`
const serializer = new XMLSerializer()

const hasAtLeastTwoDateParts = (s: string): boolean => {
  let count = 0
  const lower = s.toLowerCase()
  for (const p of ['y', 'm', 'd', 'h']) {
    if (lower.includes(p) && ++count >= 2) return true
  }
  return false
}

export const createStyles = async (zip: JSZip, target: string): Promise<Styles> => {
  let numFmts: Element | null
  let numFmtId: number
  let cachedDateFormatId: string | null = null
  let cachedLinkFormatId: string | null = null
  let isDirty: boolean = false

  // Get styles file
  const xmlStyle = await zip.file(`xl/${target}`)?.async('string')
  if (xmlStyle === undefined) throw new Error('This Excel file is corrupted: styles.xml missing')

  // Parse data to parts
  const domStyle = parser.parseFromString(xmlStyle, 'application/xml')
  const styleSheet = domStyle.querySelector('styleSheet')
  const cellXfs = domStyle.querySelector('cellXfs')
  if (styleSheet === null || cellXfs === null) throw new Error('Invalid styles.xml structure')

  // Get custom formats
  numFmts = domStyle.querySelector('numFmts')
  if (numFmts === null) {
    numFmts = domStyle.createElementNS(NS, 'numFmts')
    styleSheet.prepend(numFmts)
  }

  const existingIds = Array.from(
    domStyle.querySelectorAll('numFmt'),
    el => Number(el.getAttribute('numFmtId'))
  ).filter(Number.isFinite)

  // Create an unique id for the new format
  numFmtId = Math.max(163, ...existingIds)

  const save = (): void => {
    if (!isDirty) return
    const newXfData = serializer.serializeToString(domStyle)
    zip.file(`xl/${target}`, newXfData)
  }

  const getDateFormat = (): string => {
    const existingDateFormat = Array.from(
      domStyle.querySelectorAll('cellXfs xf')
    ).findIndex(el => {
      return ['14', '15', '16', '17', '18', '22'].includes(el.getAttribute('numFmtId') ?? '') ||
      hasAtLeastTwoDateParts(el.getAttribute('formatCode') ?? '')
    })
    if (existingDateFormat !== -1) {
      cachedDateFormatId = String(existingDateFormat)
      return cachedDateFormatId
    }
    // Create custom date format

    numFmtId++
    const numFmt = domStyle.createElementNS(NS, 'numFmt')
    numFmt.setAttribute('numFmtId', String(numFmtId))
    numFmt.setAttribute('formatCode', DATE_FORMAT)
    if (numFmts === null || cellXfs === null) throw new Error('Number formats are mailformed')
    numFmts.append(numFmt)
    numFmts.setAttribute('count', String(numFmts.children.length))

    // Create number format
    const xf = domStyle.createElementNS(NS, 'xf')
    xf.setAttribute('numFmtId', String(numFmtId))
    xf.setAttribute('fontId', '0')
    xf.setAttribute('fillId', '0')
    xf.setAttribute('borderId', '0')
    xf.setAttribute('xfId', '0')
    xf.setAttribute('applyNumberFormat', '1')
    cellXfs.append(xf)
    cellXfs.setAttribute('count', String(cellXfs.children.length))
    isDirty = true
    cachedDateFormatId = String(cellXfs.children.length - 1)
    return cachedDateFormatId
  }

  const getHyperlinkFormat = (): string => {
    const fonts = domStyle.querySelector('fonts')
    if (fonts === null || cellXfs === null) {
      throw new Error('styles.xml malformed')
    }

    // Try to reuse existing hyperlink font (underline + theme color 10)
    let fontId = Array.from(fonts.children).findIndex(font => {
      return (
        font.querySelector('u') !== undefined &&
        font.querySelector('color')?.getAttribute('theme') === '10'
      )
    })

    if (fontId === -1) {
      const font = domStyle.createElementNS(NS, 'font')

      const color = domStyle.createElementNS(NS, 'color')
      color.setAttribute('theme', '10') // Excel hyperlink theme color

      const underline = domStyle.createElementNS(NS, 'u')

      font.append(color, underline)

      fonts.append(font)
      fonts.setAttribute('count', String(fonts.children.length))
      fontId = fonts.children.length - 1
    }

    // Reuse existing XF if present
    const xfIndex = Array.from(cellXfs.children).findIndex(
      xf => xf.getAttribute('fontId') === String(fontId)
    )

    if (xfIndex !== -1) return String(xfIndex)

    // Create new XF
    const xf = domStyle.createElementNS(NS, 'xf')
    xf.setAttribute('fontId', String(fontId))
    xf.setAttribute('fillId', '0')
    xf.setAttribute('borderId', '0')
    xf.setAttribute('xfId', '0')
    xf.setAttribute('applyFont', '1')

    cellXfs.append(xf)
    cellXfs.setAttribute('count', String(cellXfs.children.length))
    isDirty = true
    cachedLinkFormatId = String(cellXfs.children.length - 1)
    return cachedLinkFormatId
  }

  return { save, getDateFormat, getHyperlinkFormat }
}
