export interface Offset {
  row: Record<number, number>
  col: Record<number, Record<number, number>>
}

export interface Offsets {
  [sheet: string]: Offset
}
