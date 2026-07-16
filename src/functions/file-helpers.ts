/**
 * Fetches a remote file and returns its contents as an ArrayBuffer.
 *
 * @param url - File URL
 * @returns Raw binary data as ArrayBuffer
 * @throws Error if the response is not OK
 */

export const fetchFile = async (url: string): Promise<ArrayBuffer> => {
  const response = await fetch(url)

  if (!response.ok) {
    throw new Error(
      `Failed to download file: ${response.status} ${response.statusText}`
    )
  }

  return await response.arrayBuffer()
}

/**
 * Downloads a file using the provided data and file name.
 * @param {string} data - The file BASE64 data
 * @param {string} fileName - The desired name for the downloaded file
 */

const downloadURL = (data: string, fileName: string): void => {
  const a = document.createElement('a')
  a.href = data
  a.download = fileName
  a.style = 'display: none'
  document.body.append(a)
  a.click()
  a.remove()
}

/**
 * Converts an array to a Blob and initiates the download with the specified file name and MIME type.
 * @param {Array} data - The array data to be converted to a Blob.
 * @param {string} fileName - The desired name for the downloaded file.
 * @param {string} mimeType - The MIME type of the file.
 */

export const downloadFile = (
  data: Blob,
  fileName: string,
  mimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
): void => {
  const blob = new Blob([data], { type: mimeType })
  const url = window.URL.createObjectURL(blob)
  downloadURL(url, fileName)
  setTimeout(() => { return window.URL.revokeObjectURL(url) }, 100)
}
