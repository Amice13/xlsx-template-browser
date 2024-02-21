/**
 * Converts a value to a string.
 * @param {*} value - The value to be converted.
 * @returns {string} - The converted string representation of the value.
 */
export const valueToString = (value) => {
  if (!value) return ''
  if (Array.isArray(value)) return value.map(valueToString).toString()
  if (typeof value === 'object') value = JSON.stringify(value)
  return value.toString()
}

/**
 * Converts dates to Excel serial index.
 * @param {Date} value - The date to be converted.
 * @returns {number} - The Excel serial index representing the date.
 */
const dateToExcel = (value) => {
  let date = new Date(value)
  return 25569 + ((date.getTime() - (date.getTimezoneOffset() * 60 * 1000)) / (1000 * 60 * 60 * 24))
}

/**
 * Guesses the Excel data type for the given JavaScript primitive value.
 * @param {*} value - The JavaScript primitive value to be converted.
 * @returns {Object|undefined} - An object with 'cellType' and 'value' properties representing the Excel data type,
 *                               or undefined if the data type is not recognized.
 */
export const guessDataType = (value) => {
  if (typeof value === 'undefined') return undefined
  if (typeof value === 'number') {
    if (isFinite(value)) return { cellType: 'n', value }
    return undefined
  }
  if (value instanceof Date) {
    if (isNaN(value)) return undefined
    return { cellType: 'n', value: dateToExcel(value) }
  }
  if (typeof value === 'boolean') return { cellType: 'b', value: value + 0 }
  return { cellType: 's', value: valueToString(value )}
}

/**
 * Creates a new shared string for an Excel file.
 * @param {string} text - The text content of the shared string.
 * @param {Element} template - (Optional) An existing node with cell to clone for the shared string.
 * @returns {Element} - The newly created or cloned shared string element.
 */
export const createSharedString = (text, template) => {
  const si = template ? template.cloneNode() : document.createElement('si')
  let t = document.createElement('t')
  t.textContent = text
  si.appendChild(t)
  return si
}

/**
 * Creates a cell with the specified value.
 * @param {Object} options - An object containing parameters for cell creation.
 * @param {*} options.value - The value to be set in the cell.
 * @param {string|number} options.row - The row reference for the cell.
 * @param {string|number} options.column - The column reference (either string or numeric) for the cell.
 * @param {Element} options.template - (Optional) An existing template to clone for the cell.
 * @param {string} options.cellType - The type of the cell (e.g., 's', 'n', 'b', 'f') for Excel formatting.
 * @returns {Element} - The created cell element.
 */
export const createCell = ({ value, row, column, template, cellType }) => {
  // If the cell is formula, then skip
  if (cellType === 'f') {
    // Check precalculated value
    const v = value.querySelector('v')
    if (v) v.remove()
    return value
  }
  // Clone initial node with the style or create a new one
  const cell = template ? template.cloneNode() : document.createElement('c')
  // Create a cell reference
  if (row && column) {
    column = typeof column === 'number' ? getExcelColumnName(column) : column
    cell.setAttribute('r', column + row)
  }
  // If no value is provided, return cell
  if (typeof value === 'undefined') return cell
  const valueTag = document.createElement('v')
  cell.setAttribute('t', cellType)
  valueTag.textContent = value
  cell.appendChild(valueTag)
  return cell
}

/**
 * Converts a numeric index to an Excel column name.
 * @param {number} index - The numeric index to be converted.
 * @returns {string} - The Excel column name corresponding to the index.
 */
const getExcelColumnName = (index) => {
  let result = ''
  while (index > 0) {
    const remainder = (index - 1) % 26
    result = String.fromCharCode(65 + remainder) + result
    index = Math.floor((index - 1) / 26)
  }
  return result
}

/**
 * Converts an Excel column name to its numeric index.
 * @param {string} columnName - The Excel column name to be converted.
 * @returns {number} - The numeric index corresponding to the Excel column name.
 */
export const getExcelColumnIndex = (columnName) => {
  columnName = columnName.replace(/\d+/g, '')
  let index = 0
  for (let i = 0; i < columnName.length; i++) {
    const charCode = columnName.charCodeAt(i) - 64
    index = index * 26 + charCode
  }
  return index
}

