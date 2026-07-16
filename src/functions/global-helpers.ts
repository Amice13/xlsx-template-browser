// This one is for a strategy pattern in the future
export const parser = new DOMParser()

export const generateUUID = (): string => {
  return crypto.randomUUID()
}
