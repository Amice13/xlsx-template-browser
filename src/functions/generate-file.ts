import { createDataStore } from './data-helpers'
import { processWorksheets } from './process-worksheets'
import { readXlsx } from './xml-helpers'
import { getWorkbook } from './get-workbook'

export const generateTemplate = async (buffer: ArrayBuffer, data: Record<string, unknown>): Promise<Blob> => {
  const xlsx = await readXlsx(buffer)
  const workbook = await getWorkbook(xlsx)
  const dataStore = createDataStore(data)
  await processWorksheets({
    workbook,
    dataStore
  })
  workbook.save()
  workbook.persons.save()
  const file = await xlsx.generateAsync({ type: 'blob' })
  return file
}
