import type { CellRef } from '../types/cells'
import type { Hyperlink, Hyperlinks } from '../types/hyperlinks'
import type { Relations } from '../types/relations'

import { NS } from './constants'
import { toCellRef } from './excel-helpers'
import { generateUUID } from './global-helpers'

const addHyperlink = (doc: Document, hyperlinks: Element, ref: string, rel: string): void => {
  const id = generateUUID()
  const el = doc.createElementNS(NS, 'hyperlink')
  el.setAttribute('ref', ref)
  el.setAttribute('r:id', rel)
  el.setAttribute('xr:uid', `{${id.toUpperCase()}}`)
  hyperlinks.appendChild(el)
}

export const getHyperlinks = ({
  xml,
  relations
}: {
  xml: Document
  relations: Relations
}): Hyperlinks => {
  let exists = true
  let hyperLinksParent = xml.querySelector('hyperlinks')
  if (hyperLinksParent === null) {
    hyperLinksParent = xml.createElementNS(NS, 'hyperlinks')
    exists = false
  }
  const hyperLinksXml = hyperLinksParent.querySelectorAll('hyperlink')
  const existingLinks: Map<string, Hyperlink> = new Map()

  for (const hyperlink of hyperLinksXml) {
    const ref = hyperlink.getAttribute('ref')
    const rId = hyperlink.getAttribute('r:id')
    const xrUid = hyperlink.getAttribute('xr:uid')
    if (ref === null || rId === null || xrUid === null) continue
    existingLinks.set(ref, {
      ref,
      'r:id': rId,
      'xr:uid': xrUid,
      xml: hyperlink
    })
  }

  const save = (): void => {
    if (exists) return
    const pageMargins = xml.querySelector('pageMargins')
    if (pageMargins !== null) pageMargins.parentNode?.insertBefore(hyperLinksParent, pageMargins)
  }

  const add = ({ range, url }: {
    range: CellRef
    url: string
  }): void => {
    let id
    const ref = toCellRef(range)
    const oldUrlElemnt = xml.querySelector(`Relationship[Target="${url}"]`)
    if (oldUrlElemnt !== null) {
      const oldUrl = oldUrlElemnt.getAttribute('Id')
      if (oldUrl !== null) id = oldUrl
    } else {
      id = relations.add({
        type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
        target: url,
        targetMode: 'External'
      })
    }
    if (id === undefined) throw new Error('The URL is mailformed')
    addHyperlink(xml, hyperLinksParent, ref, id)
  }

  return {
    add,
    save
  }
}
