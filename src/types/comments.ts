import type { DataStore } from './data-store'
import type { Drawings } from './drawings'
import type { OldComments } from './old-comments'

export interface AddComment {
  row?: number
  column?: number
  ref?: string
  text: string
}

export interface ThreadedComment {
  ref: string
  text: string
  personId?: string
}

export interface Comment {
  ref: string
  dT: string
  personId: string
  id: string
  text: string
  xml: Element
}

export interface Comments {
  add: (comment: AddComment) => void
  getComment: (ref: string) => string | undefined
  changeComments: (dataStore: DataStore) => void
  setDrawings: (params: Drawings) => void
  setOldComments: (params: OldComments) => void
  save: () => void
}

export interface GetComments {
  get: (sheet: string | number) => Promise<Comments>
  save: () => void
}
