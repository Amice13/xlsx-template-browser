import type { Cell } from './cells'

export interface Replacement {
  source: Element
  target?: number
  replacement?: unknown
  processedReplacement?: Pick<Cell, 'value' | 'type'> | Array<Pick<Cell, 'value' | 'type'>>
}
