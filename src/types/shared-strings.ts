export interface SharedStrings {
  get: (s: string | Element) => string
  oldValues: Element[]
  save: () => void
}
