import { type Cell } from './cells'
import type { Comments } from './comments'
import type { ConditionalFormattings } from './conditional-formatting'
import type { DataValidations } from './data-validation'
import type { Drawings } from './drawings'
import type { ExtenstionLists } from './extension-lists'
import type { Hyperlinks } from './hyperlinks'
import type { MergedCells } from './merged-cells'
import type { OldComments } from './old-comments'
import type { Relations } from './relations'
import type { Tables } from './tables'

export interface Sheet {
  xml: Document
  name: string
  sheetId: string
  relations: Relations
  id: string
  target: string
  cells: Cell[]
  mergedCells: MergedCells
  hyperlinks: Hyperlinks
  comments: Comments
  dataValidations: DataValidations
  conditionalFormattings: ConditionalFormattings
  extensionLists: ExtenstionLists
  drawings: Drawings
  oldComments: OldComments
  tables: Tables
  replaceCells: (element: Element) => void
  save: () => void
}
