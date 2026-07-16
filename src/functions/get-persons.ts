import type JSZip from 'jszip'
import { generateUUID, parser } from './global-helpers'
import { type Person, type Persons } from '../types/people'
import { createXml, serializeXml } from './xml-helpers'
import { NS, NS_TC } from './constants'

const addPerson = (doc: Document, person: Person): void => {
  const el = doc.createElementNS(NS_TC, 'person')
  el.setAttribute('displayName', person.displayName)
  el.setAttribute('id', person.id)
  el.setAttribute('userId', person.userId)
  el.setAttribute('providerId', person.providerId)
  doc.documentElement.appendChild(el)
}

export const getPersons = async (xlsx: JSZip, target: string): Promise<Persons> => {
  let isDirty = false
  const xmlText = await xlsx.file(`xl/${target}`)?.async('string')
  if (xmlText === undefined) isDirty = true
  const xml = xmlText === undefined
    ? createXml({ coreSchema: NS, specificSchema: NS_TC, tagName: 'personList' })
    : parser.parseFromString(xmlText, 'application/xml')

  if (xml.getElementsByTagName('parsererror').length > 0) {
    throw new Error('Invalid person.xml')
  }

  const getSystemUserId = (): string => {
    const existing = Array.from(xml.getElementsByTagNameNS(NS_TC, 'person'))
      .find(p => p.getAttribute('displayName') === 'System')?.getAttribute('id')
    if (existing !== null && existing !== undefined) return existing
    const systemPerson = {
      id: `{${generateUUID().toUpperCase()}}`,
      displayName: 'System',
      userId: 'S::system@local::00000000-0000-0000-0000-000000000000',
      providerId: 'AD'
    }
    addPerson(xml, systemPerson)
    isDirty = true
    return systemPerson.id
  }

  const save = (): void => {
    if (!isDirty) return
    const newData = serializeXml(xml)
    xlsx.file(`xl/${target}`, newData)
  }

  return {
    getSystemUserId,
    save
  }
}
