// Refer to Excel OpenXML format:
// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cell?view=openxml-2.8.1

// Import JSZip library for working with ZIP files
import JSZip from 'jszip'
import {
  createCell,
  createSharedString,
  getExcelColumnIndex,
  guessDataType,
  valueToString
} from './excelHelpers'

import {
  getByNotation
} from './helpers'

// Regular expression for parsing a property accessor enclosed in ${}
const accessorRegex = /(?<=^\$\{)(?<table>table:)?(?<accessor>[^}]+)(?=}$)/

// Global regular expression for extracting all property accessors enclosed in ${}
const globalAccessorRegex = /\$\{[^}]+}/g

const replace = async (template, data) => {
  if (!template || !data) return template

  /* INITIALIZATION */

  // Create a new DOMParser instance for parsing XML. Refer to https://developer.mozilla.org/en-US/docs/Web/API/DOMParser
  const parser = new DOMParser()
  const serializer = new XMLSerializer()
  // Create a new instance of JSZip for ZIP file manipulation
  const new_zip = new JSZip()
  // Load xlsx (zipped file)
  let res = await new_zip.loadAsync(template)

  /* PROCESSING OF THE SHARED STRINGS */

  // Read all shared text values
  let xmlText = await res.file('xl/sharedStrings.xml').async('string')
  // Get the xml with shared shared strings
  const xml = parser.parseFromString(xmlText, 'application/xml')
  // Get the header of shared strings
  const sst = xml.querySelector('sst')
  // Array with new values
  const valuesToReplace = []
  // New list of shared string values
  let newSharedStrings = []

  // We need to replace only text values with the new ones, therefore we process shared strings first
  // Get all string items
  xml.querySelectorAll('si').forEach((si, i) => {
    // Get all rich format tags
    const r = si.querySelector('r')
    if (r) {
      const xmlString = si.innerHTML
      const newString = xmlString.replace(globalAccessorRegex, s => valueToString(getByNotation(data, s.replace(/\$\{|}/g, ''))))
      si.innerHTML = newString
      newSharedStrings.push(si)
      valuesToReplace[i] = { value: newSharedStrings.length - 1, cellType: 's' }
      return false
    }
    // Get all text tags
    const t = si.querySelector('t')
    if (!t) return false
    // Get text value
    const textValue = t.textContent
    // Get all accessors
    let match = textValue.match(accessorRegex)
    // If the cell does not contain only one placeholder, check for multiple ones
    if (!match) {
      // Check if the text contains any accessors at all
      let newText = textValue.replace(globalAccessorRegex, s => valueToString(getByNotation(data, s.replace(/\$\{|}/g, ''))))
      newSharedStrings.push(createSharedString(newText, si))
      valuesToReplace[i] = { value: newSharedStrings.length - 1, cellType: 's' }
      return false
    }
    // Process a value with an accessor only
    let { accessor, table } = match.groups
    // If we have table values, we need to process them in a separate way
    const isTable = typeof table === 'string'
    let value = getByNotation(data, accessor)
    if (typeof value === 'undefined') return false
    // Save tables for further processing
    if (isTable) {
      valuesToReplace[i] = { isTable, value: value.map(guessDataType) }
      return false
    }
    if (Array.isArray(value)) {
      valuesToReplace[i] = value.map(guessDataType)
      return false
    }
    valuesToReplace[i] = guessDataType(value)
  })

  /* PROCESSING OF WORKSHEETS */

  // To prevent string duplication, store unique strings (they will be added to shared strings)
  const newStrings = []

  // Adds a string to the collection, avoiding duplication, and returns its index.
  const addString = (value) => {
    let stringIndex = newStrings.indexOf(value)
    if (stringIndex === -1) {
      newStrings.push(value)
      stringIndex = newStrings.length - 1
    }
    value = stringIndex + newSharedStrings.length
    return value
  }

  // Get all worksheets
  const worksheets = Object.keys(res.files).filter(el => /xl\/worksheets\/[^/]+$/.test(el))
  for (let worksheet of worksheets) {
    // Unzip the file
    let xmlText = await res.file(worksheet).async('string')
    // Parse file to DOM
    let xml = parser.parseFromString(xmlText, 'application/xml')
    // Process rows
    let rows = xml.querySelectorAll('sheetData row')
    // Array with the list of the new rows
    const newRows = []
    // Process cells in rows
    let rowOffset = 0
    for (let row of rows) {
      // Since some rows can be skipped, we want to get an actual row number from the template
      let currentRow = parseInt(row.getAttribute('r'))
      // List of new cells
      let newCells = []
      // If we have an array of values to extend the list of cells, we should keep an offset to move static cells
      let cellOffset = 0
      // Process each cell
      let cells = row.querySelectorAll('c').forEach((c) => {
        // Current cell index (note that Excel has 1-based index)
        let index = getExcelColumnIndex(c.getAttribute('r'))
        let newIndex = index + cellOffset
        // Get the cell value tag
        let v = c.querySelector('v') || {}
        // Check if the cell contains formula, and skip
        let isFormula = c.querySelector('f')
        if (isFormula) {
          const value = c
          const cellType = 'f'
          newCells[newIndex] = Object.assign({}, { value, cellType }, { template: c })
          return false          
        }
        // Check if the cell contains string
        let isString = c.getAttribute('t')
        if (!isString || isString !== 's') {
          newCells[newIndex] = Object.assign({}, { template: c })
          return false
        }
        // Get the new value to replace
        let newValue = valuesToReplace[v.textContent] || {}
        if (!Array.isArray(newValue)) {
          let { value, cellType, isTable } = newValue
          // If the new value is an array from the table, then return an array of values
          if (isTable) {
            newCells[newIndex] = value.map(el => {
              if (!el) return { template: c }
              if (el.cellType === 's' && typeof el.value === 'string') el.value = addString(el.value)
              return { value: el.value, cellType: el.cellType, template: c }
            })
            return false
          }
          // Return a new value
          if (cellType === 's' && typeof value === 'string') value = addString(value)
          newCells[newIndex] = Object.assign({}, { value, cellType }, { template: c })
          return false
        }
        // If the value is an array (not from table), then extend the existing list of cells
        for (let i = 0; i < newValue.length; i++) {
          let { value, cellType } = newValue[i] || {}
          if (cellType === 's' && typeof value === 'string') value = addString(value)
          newCells[newIndex + i] = Object.assign({}, { value, cellType }, { template: c })
          if (i) cellOffset++
        }
      })

      /*  GENERATION OF THE NEW ROWS */

      // Check if the row contains arrays and the values should be duplicated
      let length = Math.max(...newCells.filter(el => Array.isArray(el)).map(el => el.length))

      // Create a row
      if (length <= 0) {
        let rowIndex = rowOffset + currentRow
        const rowValues = newCells.map((el, i) => {
          if (!el) return undefined
          return createCell(Object.assign({}, el, { row: rowIndex, column: i }))
        }).filter(Boolean)
        const newRow = row.cloneNode()
        newRow.setAttribute('r', rowIndex)
        rowValues.forEach(value => newRow.append(value))
        newRows.push(newRow)
        continue
      }

      // Create table rows
      for (let i = 0; i < length; i++) {
        let rowIndex = rowOffset + currentRow
        rowOffset++
        const rowValues = newCells.map((el, index) => {
          if (!el) return undefined
          if (Array.isArray(el)) {
            if (typeof el[i] === 'undefined') return undefined
            return createCell(Object.assign({}, el[i], { row: rowIndex, column: index }))
          }
          return createCell(Object.assign({}, el, { row: rowIndex, column: index }))
        })
        const newRow = row.cloneNode()
        newRow.setAttribute('r', rowIndex)
        rowValues.forEach(value => newRow.append(value))
        newRows.push(newRow)
      }
    }

    // Generate new sheet data
    const newSheetData = xml.querySelector('sheetData').cloneNode()
    newRows.forEach(row => newSheetData.append(row))
    xml.querySelector('sheetData').replaceWith(newSheetData)
    const newData = serializer.serializeToString(xml).replace(/ ?xmlns="http:\/\/www\.w3\.org\/1999\/xhtml"/g, '')
    // Save the data of the worksheet
    await new_zip.file(worksheet, newData)
  }

  /* FINALIZE NEW SHARED STRINGS */
  newSharedStrings = [...newSharedStrings, ...newStrings.map(el => createSharedString(el))]
  const newSst = sst.cloneNode()
  newSharedStrings.forEach(si => newSst.append(si))
  sst.replaceWith(newSst)
  const newData = serializer.serializeToString(xml).replace(/ ?xmlns="http:\/\/www\.w3\.org\/1999\/xhtml"/g, '')
  // Save new shared strings
  await new_zip.file('xl/sharedStrings.xml', newData)
  return await new_zip.generateAsync({ type: 'blob' })
}

export default replace
