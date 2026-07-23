import type JSZip from 'jszip'
import type { ThreadedComment, Comments, Comment, AddComment } from '../types/comments'
import { createXml, serializeXml } from './xml-helpers'
import type { Relations } from '../types/relations'
import { NS, NS_TC } from './constants'
import { parser } from './global-helpers'
import type { Workbook } from '../types/workbook'
import type { DataStore } from '../types/data-store'
import { replaceAccessors } from './get-replacements'
import type { OldComments } from '../types/old-comments'
import type { Drawings } from '../types/drawings'
import { parseCellRef, toCellRef } from './excel-helpers'

const addThreadedComment = (doc: Document, c: ThreadedComment): string => {
  const id = crypto.randomUUID()

  const el = doc.createElementNS(NS_TC, 'threadedComment')
  el.setAttribute('ref', c.ref)
  el.setAttribute('dT', new Date().toISOString())
  el.setAttribute('personId', c.personId ?? '')
  el.setAttribute('id', '{' + id.toUpperCase() + '}')

  const text = doc.createElementNS(NS_TC, 'text')
  text.textContent = c.text

  el.appendChild(text)
  doc.documentElement.appendChild(el)

  return id
}

export const getComments = async ({
  xlsx,
  relations,
  workbook,
  id
}: {
  xlsx: JSZip
  relations: Relations
  workbook: Workbook
  id: string
}): Promise<Comments> => {
  let isDirty = false
  const guessedFilename = `../threadedComments/threadedComment${id.replace('rId', '')}.xml`

  const commentsRelations = relations.getElements('threadedComment')
  if (commentsRelations.length > 1) throw new Error('Comments are mailformed')
  let relationExists = commentsRelations.length > 0

  const commentsFileName = (commentsRelations[0]?.target ?? guessedFilename).replace('..', 'xl')

  const xmlText = await xlsx.file(commentsFileName)?.async('string')
  const xml = xmlText === undefined
    ? createXml({ coreSchema: NS, specificSchema: NS_TC, tagName: 'threadedComments' })
    : parser.parseFromString(xmlText, 'application/xml')

  if (xmlText === undefined) isDirty = true
  if (xml.getElementsByTagName('parsererror').length > 0) {
    throw new Error(`Invalid relations file for ${commentsFileName}`)
  }

  const oldCommentsMap = new Map<string, Comment>()
  const commentsXml = xml.querySelectorAll('threadedComment')

  for (const comment of commentsXml) {
    const ref = comment.getAttribute('ref')
    const dT = comment.getAttribute('dT')
    const personId = comment.getAttribute('personId')
    const id = comment.getAttribute('id')
    const text = comment.querySelector('text')?.textContent
    if (ref === null || dT === null || personId === null || id === null) continue
    oldCommentsMap.set(ref, {
      ref,
      dT,
      personId,
      id,
      text: text ?? '',
      xml: comment
    })
  }

  const changeComments = (datastore: DataStore): void => {
    for (const [, comment] of oldCommentsMap) {
      comment.text = replaceAccessors(comment.text, datastore)
      const text = comment.xml.querySelector('text')
      if (text === null) continue
      text.textContent = comment.text
      isDirty = true
    }
  }

  let oldComments: OldComments
  let drawings: Drawings

  const setOldComments = (props: OldComments): void => {
    oldComments = props
  }

  const setDrawings = (props: Drawings): void => {
    drawings = props
  }

  const add = ({ ref, text, row, column }: AddComment): void => {
    addRelation()
    if (ref === undefined) {
      if (row === undefined || column === undefined) {
        throw new Error('Comment position is undefined')
      }
      ref = toCellRef({ row, column })
    }
    if (row === undefined && column === undefined) {
      if (ref === undefined) {
        throw new Error('Comment position is undefined')
      }
      const location = parseCellRef(ref)
      row = location.row
      column = location.column
    }
    // TODO: Support person creation
    // const { persons } = workbook
    const systemPersonId = '{00000000-0000-0000-0000-000000000000}'
    const id = addThreadedComment(xml, {
      personId: systemPersonId,
      ref,
      text
    })
    oldComments?.add({ id: `{${id.toUpperCase()}}`, ref, text })
    drawings?.add({ row: String(row), column: String(column) })
    isDirty = true
  }

  const getComment = (ref: string): string | undefined => {
    return oldCommentsMap.get(ref)?.ref
  }

  const addRelation = (): void => {
    if (relationExists) return
    const commentsRelation = {
      type: 'http://schemas.microsoft.com/office/2017/10/relationships/threadedComment',
      target: `../threadedComments/threadedComment${id.replace('rId', '')}.xml`
    }
    relations.add(commentsRelation)
    relationExists = true
  }

  const save = (): void => {
    if (!isDirty) return
    if (oldComments !== undefined) oldComments.save()
    if (drawings !== undefined) drawings.save()
    const text = serializeXml(xml)
    xlsx.file(commentsFileName, text)
  }

  return {
    add,
    changeComments,
    getComment,
    setDrawings,
    setOldComments,
    save
  }
}
