import type JSZip from 'jszip'
import type { GetComments } from './comments'

export interface Relation {
  id?: string
  target: string
  type: string
  realType?: string
  targetMode?: string
}

export interface Relations {
  add: (rel: Relation) => string
  get: ({ by, value }: { by: keyof Relation, value: string }) => Relation[]
  remove: ({ by, value }: { by: keyof Relation, value: string }) => void
  getElements: (el: string) => Relation[]
  save: () => void
}

export interface GetRelationsParams {
  xlsx: JSZip
  comments: GetComments
}

export interface GetRelations {
  get: (worksheetName: string) => Promise<Relation>
  save: () => void
}
