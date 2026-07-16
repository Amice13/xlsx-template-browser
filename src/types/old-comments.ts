export interface OldComment {
  ref: string
  id: string
  text: string
}

export interface OldComments {
  add: (comment: OldComment) => void
  save: () => void
}
