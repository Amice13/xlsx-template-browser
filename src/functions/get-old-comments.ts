import type JSZip from 'jszip'
import type { Relations } from '../types/relations'
import { parser } from './global-helpers'
import { createXml, serializeXml } from './xml-helpers'
import { NS } from './constants'
import type { OldComment, OldComments } from '../types/old-comments'

const EXCEL_NOTIFICATION = `[Threaded comment]

Your version of Excel allows you to read this threaded comment; however, any edits to it will get removed if the file is opened in a newer version of Excel. Learn more: https://go.microsoft.com/fwlink/?linkid=870924

Comment:
    `

const NS_MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
const NS_XR = 'http://schemas.microsoft.com/office/spreadsheetml/2014/revision'

const createComment = (
  { xml, ref, id, text, author }: OldComment & { xml: Document, author: string }
): Element => {
  const comment = xml.createElementNS(NS, 'comment')
  comment.setAttribute('ref', ref)
  comment.setAttribute('xr:uid', id)
  comment.setAttribute('authorId', author)
  comment.setAttribute('shapeId', '0')
  const textTag = xml.createElementNS(NS, 'text')
  const tTag = xml.createElementNS(NS, 't')
  tTag.textContent = text
  textTag.appendChild(tTag)
  comment.appendChild(textTag)
  return comment
}

const createAuthor = ({ xml, id }: { xml: Document, id: string }): Element => {
  const author = xml.createElementNS(NS, 'author')
  author.textContent = `tc=${id}`
  return author
}

export const getOldComments = async ({
  id,
  relations,
  xlsx
}: {
  id: string
  relations: Relations
  xlsx: JSZip
}): Promise<OldComments> => {
  const oldCommentsRelation = relations.getElements('comments')
  const guessedFilename = `../comments${id.replace('rId', '')}.xml`
  const noRelation = oldCommentsRelation.length === 0
  const filename = (oldCommentsRelation[0]?.target ?? guessedFilename).replace('..', 'xl')

  let isDirty = false
  const comments: OldComment[] = []

  const xmlText = await xlsx.file(filename)?.async('string')
  const xml = xmlText === undefined
    ? createXml({ specificSchema: NS, tagName: 'comments' })
    : parser.parseFromString(xmlText, 'application/xml')

  if (xmlText === undefined) {
    xml.querySelector('comments')?.setAttribute('xmlns:mc', NS_MC)
    xml.querySelector('comments')?.setAttribute('xmlns:xr', NS_XR)
    xml.querySelector('comments')?.setAttribute('mc:Ignorable', 'xr')
    xml.querySelector('comments')?.appendChild(xml.createElementNS(NS, 'authors'))
    xml.querySelector('comments')?.appendChild(xml.createElementNS(NS, 'commentList'))
  }
  const parent = xml.querySelector('comments')
  const authors = xml.querySelector('authors')
  const commentList = xml.querySelector('commentList')
  if (parent === null || authors === null || commentList === null) throw new Error('Comments are mailformed')

  const uniqueRefs = new Set<string>()

  if (xmlText !== undefined) {
    const currentComments = commentList.getElementsByTagName('comment')
    for (const comment of currentComments) {
      const ref = comment.getAttribute('ref')
      const id = comment.getAttribute('xr:uid')
      let text = comment.querySelector('t')?.textContent ?? ''
      text = text.replace(EXCEL_NOTIFICATION, '')
      if (ref === null || id === null) continue
      comments.push({ ref, id, text })
    }
  }

  const add = ({ ref, id, text }: OldComment): void => {
    if (uniqueRefs.has(ref)) return
    uniqueRefs.add(ref)
    comments.push({ ref, id, text })
    isDirty = true
  }

  const save = (): void => {
    if (!isDirty) return
    if (noRelation) {
      relations.add({
        target: filename.replace('xl', '..'),
        type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments'
      })
    }
    const newComments: Element[] = []
    const newAuthors: Element[] = []
    let index = 0
    for (const comment of comments) {
      const newComment = createComment({ xml, author: String(index), ...comment })
      const author = createAuthor({ xml, id: comment.id })
      newComments.push(newComment)
      newAuthors.push(author)
      index++
    }
    authors.replaceChildren(...newAuthors)
    commentList?.replaceChildren(...newComments)
    const data = serializeXml(xml)
    xlsx.file(filename, data)
  }

  return {
    add,
    save
  }
}
