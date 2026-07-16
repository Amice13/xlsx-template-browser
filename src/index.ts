import { downloadFile, fetchFile } from './functions/file-helpers'
import { generateTemplate } from './functions/generate-file'

export const generateXlsx = async (template: string | ArrayBuffer, data: Record<string, unknown>): Promise<Blob> => {
  if (template === undefined || data === undefined) throw Error('No template or data provided')
  if (typeof template === 'string') {
    if (!/^https|ftp?:/.test(template)) throw new Error('Template URL is invalid')
    const response = await fetchFile(template)
    const fileContent = await generateTemplate(response, data)
    return fileContent
  }
  return await generateTemplate(template, data)
}

/**
 * Downloads an XLSX file by replacing data in a template file and initiates the download in the browser.
 * @param {string} templateURL - The URL of the template XLSX file.
 * @param {Object} data - The data object containing values to replace in the template.
 * @param {string} [fileName] - The name of the file to be downloaded. If not provided, a default name based on the current date will be used.
 * @returns {Promise<void>} A promise that resolves once the download is initiated.
 */

export const downloadXlsx = async (
  template: string | ArrayBuffer,
  data: Record<string, unknown>,
  fileName?: string,
  mimeType?: string
): Promise<void> => {
  if (fileName === undefined) fileName = new Date().toLocaleDateString('sv') + ' - Report.xlsx'
  const fileContent = await generateXlsx(template, data)
  downloadFile(fileContent, fileName, mimeType)
}
