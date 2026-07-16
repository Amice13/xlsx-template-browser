import type { Persons } from './people'
import type { PivotTables } from './pivot-tables'
import type { Relations } from './relations'
import type { SharedStrings } from './shared-strings'
import type { Sheet } from './sheets'
import type { Styles } from './styles'
import type { DataValidations } from './data-validation'
import type { ExtenstionLists } from './extension-lists'

export interface Workbook {
  sheets: Sheet[]
  theme?: {
    id: string
    target: string
  }
  calcChain?: {
    id: string
    target: string
  }
  sharedStrings: SharedStrings
  styles: Styles
  pivotTables: PivotTables
  dataValidations: DataValidations
  extensionLists: ExtenstionLists
  relations: Relations
  persons: Persons
  save: () => void
}
