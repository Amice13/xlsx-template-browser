export interface Drawings {
  add: ({ row, column }: { row: string, column: string }) => void
  save: () => void
}
