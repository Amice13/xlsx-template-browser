export interface Person {
  id: string
  displayName: string
  userId: string
  providerId: string
}

export interface Persons {
  getSystemUserId: () => string
  save: () => void
}
