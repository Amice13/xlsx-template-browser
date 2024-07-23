import { downloadBlob } from './helpers'
import replace from './main'

/**
 * Generates an XLSX file by replacing data in a template file.
 * @param {string} templateURL - The URL of the template XLSX file.
 * @param {Object} data - The data object containing values to replace in the template.
 * @returns {Promise<Blob>} A promise that resolves with the Blob of the generated XLSX file.
 * @throws {Error} Throws an error if either the template URL or data is not provided, or if unable to fetch the template file.
 */
export const generateXlsx = async (template, data) => {
  if (!template || !data) throw Error('No template or data provided')
  if (typeof template === 'string' && template.match(/^https?:/)) {
    const response = await fetch(template).catch(err => { return false})
    if (!response) throw Error('Can\'t fetch the template URL')
    template = response.arrayBuffer()
  }
  return await replace(template, data)
}

/**
 * Downloads an XLSX file by replacing data in a template file and initiates the download in the browser.
 * @param {string} templateURL - The URL of the template XLSX file.
 * @param {Object} data - The data object containing values to replace in the template.
 * @param {string} [fileName] - The name of the file to be downloaded. If not provided, a default name based on the current date will be used.
 * @returns {Promise<void>} A promise that resolves once the download is initiated.
 */
export const downloadXlsx = async (template, data, fileName, mimeType) => {
  if (!fileName) fileName = new Date().toISOString().substring(0, 10) + ' - Report.xlsx'
  downloadBlob(await generateXlsx(template, data), fileName, mimeType)
}
