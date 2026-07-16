import type JSZip from 'jszip'
import { createXml, serializeXml } from './xml-helpers'
import { parser } from './global-helpers'
import { NS_REL } from './constants'
import type { Relation, Relations } from '../types/relations'

export const getRelations = async ({ xlsx, filename }: {
  xlsx: JSZip
  filename: string
}): Promise<Relations> => {
  let isDirty = false
  const relFilename = `xl/${filename}`
  const xmlText = await xlsx.file(relFilename)?.async('string')
  const xml = xmlText === undefined
    ? createXml({ specificSchema: NS_REL, tagName: 'Relationships' })
    : parser.parseFromString(xmlText, 'application/xml')

  if (xmlText === undefined) isDirty = true
  if (xml.getElementsByTagName('parsererror').length > 0) {
    throw new Error(`Invalid relations file for ${filename}`)
  }
  const realationsXml = Array.from(xml.querySelectorAll('Relationship'))
  let relations: Relation[] = []

  for (const el of realationsXml) {
    const id = el.getAttribute('Id')
    const target = el.getAttribute('Target')
    const elementType = el.getAttribute('Type')
    const targetMode = el.getAttribute('TargetMode')
    if (id === null || target === null || elementType === null) continue
    const realType = elementType.slice(elementType.lastIndexOf('/') + 1)
    relations.push({
      id,
      target,
      type: elementType,
      realType,
      ...(targetMode === null ? {} : { targetMode })
    })
  }

  const getNextId = (): string => {
    let max = 0
    xml.querySelectorAll('Relationship[Id^=\'rId\']').forEach(el => {
      const n = Number((el.getAttribute('Id') ?? 'Id1').slice(3))
      if (!Number.isNaN(n)) max = Math.max(max, n)
    })
    return `rId${max + 1}`
  }

  const add = (rel: Relation): string => {
    const id = getNextId()
    const root = xml.querySelector('Relationships')
    if (root === null) return 'NA'
    const el = xml.createElementNS(NS_REL, 'Relationship')
    el.setAttribute('Id', id)
    el.setAttribute('Type', rel.type)
    el.setAttribute('Target', rel.target)
    if (rel.targetMode !== undefined) el.setAttribute('TargetMode', rel.targetMode)
    root.appendChild(el)
    relations.push({ id, type: rel.type, target: rel.target })
    isDirty = true
    return id
  }

  const get = ({ by, value }: { by: keyof Relation, value: string }): Relation[] => {
    return relations.filter(relation => relation[by] === value)
  }

  const remove = ({ by, value }: { by: keyof Relation, value: string }): void => {
    const valuesToRemove = relations.filter(relation => relation[by] === value)
    for (const value of valuesToRemove) {
      xml.querySelector(`Relationship[id="${String(value)}"]`)?.remove()
    }
    relations = relations.filter(relation => relation[by] !== value)
  }

  const save = (): void => {
    if (!isDirty) return
    const newData = serializeXml(xml)
    xlsx.file(relFilename, newData)
  }

  const getElements = (element: string): Relation[] => {
    return relations.filter(el => el.realType === element)
  }

  return {
    add,
    remove,
    get,
    save,
    getElements
  }
}
