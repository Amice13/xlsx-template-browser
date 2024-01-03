/**
 * Downloads a file using the provided data and file name.
 * @param {string} data - The file data
 * @param {string} fileName - The desired name for the downloaded file
 */

const downloadURL = (data, fileName) => {
  const a = document.createElement('a')
  a.href = data
  a.download = fileName
  a.style = 'display: none'
  document.body.appendChild(a)
  a.click()
  a.remove()
}

/**
 * Converts an array to a Blob and initiates the download with the specified file name and MIME type.
 * @param {Array} data - The array data to be converted to a Blob.
 * @param {string} fileName - The desired name for the downloaded file.
 * @param {string} mimeType - The MIME type of the file.
 */
export const downloadBlob = (data, fileName, mimeType) => {
  const blob = new Blob([data], { type: mimeType })
  const url = window.URL.createObjectURL(blob)
  downloadURL(url, fileName)
  setTimeout(() => { return window.URL.revokeObjectURL(url) }, 100)
}


// Regular expression for extracting fields with deep notation data
const notationRegex = /(?<=\[(?<qoute>['"]))[^'"\]].*?(?=\k<qoute>\])|(?<=\[)[^'"\]].*?(?=\])|[^.\["''\]]+(?=\.|\[|$)/g

/**
 * Retrieves nested properties from an object using an array of accessors.
 * @param {*} obj - The object from which to retrieve the nested properties.
 * @param {Array} accessors - An array of accessors representing the path to the desired properties.
 * @returns {*} - The value of the nested properties, or undefined if any accessor along the path is undefined.
 */
const getDeep = (obj, accessors) => {
  let length = accessors.length
  for (let i = 0; i < length; i++) {
    if (typeof obj === 'undefined') return undefined
    if (Array.isArray(obj) && typeof accessors[i] === 'string') return obj.map(el => getDeep(el, accessors.slice(i)))
    obj = obj[accessors[i]]
  }
  return obj
}

/**
 * Retrieves nested properties from an object using dot and bracket notation.
 * @param {*} obj - The object from which to retrieve the nested properties.
 * @param {string} props - The notation representing the path to the desired properties.
 * @returns {*} - The value of the nested properties, or the original object if the notation is empty.
 */
export const getByNotation = (obj, props) => {
  let accessors = props.match(notationRegex)
  if (!accessors.length) return obj
  accessors = accessors.map(el => /^\d+$/.test(el) ? parseInt(el) : el)
  return getDeep(obj, accessors)
}

