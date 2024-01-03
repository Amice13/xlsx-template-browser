import { downloadBlob } from './helpers'
import replace from './main'

export const generateXlsx = async (template, data) => {
  if (!template || !data) throw Error('No template or data provided')
  if (typeof template.match === 'string' && template.match(/^https:/)) {
    const response = await fetch(template)
    template = response.arrayBuffer()
  }
  return await replace(template, data)
}

export const downloadXlsx = async (template, data, fileName, mimeType) => {
  if (!fileName) fileName = new Date().toISOString().substring(0, 10) + ' - Report.xlsx'
  downloadBlob(await generateXlsx(template, data), fileName, mimeType)
}
