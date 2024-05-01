import { downloadBlob } from './helpers'
import replace from './main'

/**
 * Generates an XLSX file by replacing data in a template file.
 * @param {string} templateURL - The URL of the template XLSX file.
 * @param {Object} data - The data object containing values to replace in the template.
 * @returns {Promise<Blob>} A promise that resolves with the Blob of the generated XLSX file.
 * @throws {Error} Throws an error if either the templateURL or data is not provided, or if unable to fetch the template file.
 */
export async function generateXlsx(templateURL, data) {
  if (!templateURL || !data) throw new Error('No templateURL or data provided')
  return await fetch(templateURL)
    .then(response => {
      if (response.ok) return response.arrayBuffer()
      throw new Error(`Unable to fetch the template at: ${templateURL}`)
    })
    .then(template => replace(template, data))
    .catch(e => console.error(e))
}

/**
 * Downloads an XLSX file by replacing data in a template file and initiates the download in the browser.
 * @param {string} templateURL - The URL of the template XLSX file.
 * @param {Object} data - The data object containing values to replace in the template.
 * @param {string} [fileName] - The name of the file to be downloaded. If not provided, a default name based on the current date will be used.
 * @returns {Promise<void>} A promise that resolves once the download is initiated.
 */
export async function downloadXlsx(templateURL, data, fileName) {
  fileName = fileName || `${new Date().toISOString().substring(0, 10)} - Report.xlsx`
  const mimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  downloadBlob(await generateXlsx(templateURL, data), fileName, mimeType)
}
