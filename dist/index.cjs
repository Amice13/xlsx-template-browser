"use strict";
var __create = Object.create;
var __defProp = Object.defineProperty;
var __getOwnPropDesc = Object.getOwnPropertyDescriptor;
var __getOwnPropNames = Object.getOwnPropertyNames;
var __getProtoOf = Object.getPrototypeOf;
var __hasOwnProp = Object.prototype.hasOwnProperty;
var __export = (target, all) => {
  for (var name in all)
    __defProp(target, name, { get: all[name], enumerable: true });
};
var __copyProps = (to, from, except, desc) => {
  if (from && typeof from === "object" || typeof from === "function") {
    for (let key of __getOwnPropNames(from))
      if (!__hasOwnProp.call(to, key) && key !== except)
        __defProp(to, key, { get: () => from[key], enumerable: !(desc = __getOwnPropDesc(from, key)) || desc.enumerable });
  }
  return to;
};
var __toESM = (mod, isNodeMode, target) => (target = mod != null ? __create(__getProtoOf(mod)) : {}, __copyProps(
  // If the importer is in node compatibility mode or this is not an ESM
  // file that has been converted to a CommonJS file using a Babel-
  // compatible transform (i.e. "__esModule" has not been set), then set
  // "default" to the CommonJS "module.exports" for node compatibility.
  isNodeMode || !mod || !mod.__esModule ? __defProp(target, "default", { value: mod, enumerable: true }) : target,
  mod
));
var __toCommonJS = (mod) => __copyProps(__defProp({}, "__esModule", { value: true }), mod);

// src/index.ts
var index_exports = {};
__export(index_exports, {
  downloadXlsx: () => downloadXlsx,
  generateXlsx: () => generateXlsx
});
module.exports = __toCommonJS(index_exports);

// src/functions/file-helpers.ts
var fetchFile = async (url) => {
  const response = await fetch(url);
  if (!response.ok) {
    throw new Error(
      `Failed to download file: ${response.status} ${response.statusText}`
    );
  }
  return await response.arrayBuffer();
};
var downloadURL = (data, fileName) => {
  const a = document.createElement("a");
  a.href = data;
  a.download = fileName;
  a.style = "display: none";
  document.body.append(a);
  a.click();
  a.remove();
};
var downloadFile = (data, fileName, mimeType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet") => {
  const blob = new Blob([data], { type: mimeType });
  const url = window.URL.createObjectURL(blob);
  downloadURL(url, fileName);
  setTimeout(() => {
    return window.URL.revokeObjectURL(url);
  }, 100);
};

// src/functions/data-helpers.ts
var isNumeric = (v) => {
  return /^\d+$/.test(v);
};
var parseNotation = (input) => {
  const result = [];
  let i = 0;
  while (i < input.length) {
    const char = input[i];
    if (char === ".") {
      i++;
      continue;
    }
    if (char === "[") {
      i++;
      const quote = input[i] === '"' || input[i] === "'" ? input[i++] : null;
      let value2 = "";
      while (i < input.length) {
        if (quote !== null) {
          if (input[i] === quote && input[i + 1] === "]") break;
        } else {
          if (input[i] === "]") break;
        }
        value2 += input[i++];
      }
      i += quote !== null ? 2 : 1;
      result.push(isNumeric(value2) ? Number(value2) : value2);
      continue;
    }
    let value = "";
    while (i < input.length) {
      const c = input[i];
      if (c === "." || c === "[") break;
      value += c;
      i++;
    }
    result.push(value);
  }
  return result;
};
var getDeep = (obj, accessors) => {
  for (let i = 0; i < accessors.length; i++) {
    if (obj === void 0) {
      return void 0;
    }
    if (obj === null) {
      return obj;
    }
    if (Array.isArray(obj) && typeof accessors[i] === "string") {
      return obj.map((el) => getDeep(el, accessors.slice(i)));
    }
    const accessor = accessors[i];
    if (accessor === void 0) {
      continue;
    }
    obj = obj[accessor];
  }
  return obj;
};
var getByNotation = (obj, props) => {
  const parsed = parseNotation(props);
  const value = getDeep(obj, parsed);
  return value ?? "";
};
var createDataStore = (data) => {
  const accessors = /* @__PURE__ */ new Map();
  const get = (accessor) => {
    accessor = accessor.trim();
    if (accessors.has(accessor)) {
      return accessors.get(accessor);
    }
    const value = getByNotation(data, accessor);
    accessors.set(accessor, value);
    return value;
  };
  return { get };
};
var isElement = (value) => {
  return value instanceof Element || value instanceof Node && value.nodeType === Node.ELEMENT_NODE;
};

// src/functions/constants.ts
var NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
var NS_PREFIX = "http://schemas.openxmlformats.org/officeDocument/2006";
var NS_TC = "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments";
var GLOBAL_RELS = "_rels/.rels";
var OFFICE_DOCUMENT_TYPE = `${NS_PREFIX}/relationships`;
var NS_REL = "http://schemas.openxmlformats.org/package/2006/relationships";
var TAGS_ORDER = [
  "sheetData",
  "mergeCells",
  "conditionalFormatting",
  "dataValidations",
  "pageMargins",
  "tableParts",
  "extLst",
  "legacyDrawing"
];

// src/functions/excel-helpers.ts
var toCellRef = (ref) => {
  let col = ref.column;
  if (col <= 0) throw new Error(`Invalid column: ${col}`);
  let column = "";
  while (col > 0) {
    const rem = (col - 1) % 26;
    column = String.fromCharCode(65 + rem) + column;
    col = Math.floor((col - 1) / 26);
  }
  const colPart = `${ref.columnAbsolute === true ? "$" : ""}${column}`;
  const rowPart = `${ref.rowAbsolute === true ? "$" : ""}${ref.row}`;
  if (ref.row <= 0) throw new Error(`Invalid row: ${ref.row}`);
  return colPart + rowPart;
};
var parseCellRef = (ref) => {
  const normalized = ref.trim().toUpperCase();
  let column = 0;
  let i = 0;
  let columnAbsolute = false;
  let rowAbsolute = false;
  if (normalized[i] === "$") {
    columnAbsolute = true;
    i++;
  }
  for (; i < normalized.length; i++) {
    const code = normalized.charCodeAt(i);
    if (code < 65 || code > 90) break;
    column = column * 26 + (code - 64);
  }
  if (normalized[i] === "$") {
    rowAbsolute = true;
    i++;
  }
  const rowStr = normalized.slice(i);
  const row = Number(rowStr);
  if (column === 0 || rowStr === "" || Number.isNaN(row)) {
    throw new Error(`Invalid cell reference: ${ref}`);
  }
  return { row, column, rowAbsolute, columnAbsolute };
};
var parseCellRange = (ref) => {
  const parts = ref.split(":");
  if (parts.length === 0 || parts.length > 2) {
    throw new Error(`Invalid cell range: ${ref}`);
  }
  if (parts.length === 1) parts.push(parts[0]);
  const [start, end] = parts;
  if (start === void 0 || end === void 0) throw new Error(`Cell ${ref} is broken`);
  const startRef = parseCellRef(start);
  const endRef = parseCellRef(end);
  return {
    rowStart: startRef.row,
    columnStart: startRef.column,
    rowAbsoluteStart: startRef.rowAbsolute,
    columnAbsoluteStart: startRef.columnAbsolute,
    rowEnd: endRef.row,
    columnEnd: endRef.column,
    rowAbsoluteEnd: endRef.rowAbsolute,
    columnAbsoluteEnd: endRef.columnAbsolute
  };
};
var toCellRange = (range) => {
  if (range.rowStart === void 0 && range.rowEnd === void 0) {
    if (range.columnStart === void 0 || range.columnEnd === void 0) {
      throw new Error("Range is not defined");
    }
    return [
      range.columnAbsoluteStart === true ? "$" : "",
      String(range.columnStart),
      ":",
      range.columnAbsoluteEnd === true ? "$" : "",
      String(range.columnEnd)
    ].join("");
  }
  if (range.columnStart === void 0 && range.columnEnd === void 0) {
    if (range.rowStart === void 0 || range.rowEnd === void 0) {
      throw new Error("Range is not defined");
    }
    return [
      range.rowAbsoluteStart === true ? "$" : "",
      getExcelColumnName(range.rowStart),
      ":",
      range.rowAbsoluteEnd === true ? "$" : "",
      getExcelColumnName(range.rowEnd)
    ].join("");
  }
  if (range.rowStart === void 0 || range.columnStart === void 0 || range.rowEnd === void 0 || range.columnEnd === void 0) {
    throw new Error("Range is not defined");
  }
  const start = toCellRef({
    row: range.rowStart,
    column: range.columnStart,
    rowAbsolute: range.rowAbsoluteStart,
    columnAbsolute: range.columnAbsoluteStart
  });
  if (range.rowEnd === 0 && range.columnEnd === 0) return start;
  const end = toCellRef({
    row: range.rowEnd,
    column: range.columnEnd,
    rowAbsolute: range.rowAbsoluteEnd,
    columnAbsolute: range.columnAbsoluteEnd
  });
  return `${start}:${end}`;
};
var getExcelColumnName = (index) => {
  let result = "";
  while (index > 0) {
    index--;
    result = String.fromCharCode(65 + index % 26) + result;
    index = Math.floor(index / 26);
  }
  return result;
};
var dateToExcel = (value) => {
  const date = new Date(value);
  const dateValue = 25569 + (date.getTime() - date.getTimezoneOffset() * 60 * 1e3) / (1e3 * 60 * 60 * 24);
  return String(dateValue);
};
var convertValueToExcel = ({ value, cellType }) => {
  if (cellType === null) {
    return null;
  }
  switch (cellType) {
    case "formula": {
      if (typeof value === "object") {
        value = String(value);
      }
      if (typeof value === "string") {
        return value.startsWith("=") ? value.slice(1) : value;
      }
      break;
    }
    case "date": {
      if (value instanceof Date || typeof value === "string") {
        return dateToExcel(value);
      }
      if (typeof value === "object") {
        return dateToExcel(String(value));
      }
      break;
    }
    case "boolean": {
      if (typeof value === "boolean") {
        return String(Number(value));
      }
      if (typeof value === "object") {
        return String(Number(/true/i.test(String(value))));
      }
      break;
    }
    case "number": {
      if (typeof value === "object") {
        return String(value);
      }
      if (typeof value === "number") {
        return value.toString();
      }
      break;
    }
    default: {
      if (typeof value === "object") {
        return JSON.stringify(value);
      }
    }
  }
  return String(value);
};
var urlPatten = String.raw`^(?!mailto:)(?:(?:http|https|ftp)://)(?:\S+(?::\S*)?@)?(?:(?:(?:[1-9]\d?|1\d\d|2[01]\d|22[0-3])(?:\.(?:1?\d{1,2}|2[0-4]\d|25[0-5])){2}(?:\.(?:[0-9]\d?|1\d\d|2[0-4]\d|25[0-4]))|(?:(?:[a-z\u00a1-\uffff0-9]+-?)*[a-z\u00a1-\uffff0-9]+)(?:\.(?:[a-z\u00a1-\uffff0-9]+-?)*[a-z\u00a1-\uffff0-9]+)*(?:\.(?:[a-z\u00a1-\uffff]{2,})))|localhost)(?::\d{2,5})?(?:(/|\?|#)[^\s]*)?$`;
var urlRegex = new RegExp(urlPatten, "i");
var guessDataType = (value) => {
  if (value === void 0 || value === null || value === "") {
    return null;
  }
  if (typeof value === "number") {
    return "number";
  }
  if (typeof value === "boolean") {
    return "boolean";
  }
  if (value instanceof Date) {
    return isNaN(value.getTime()) ? null : "date";
  }
  if (typeof value === "string") {
    if (value.startsWith("=")) return "formula";
    if (urlRegex.test(value)) return "url";
    return "string";
  }
  if (typeof value === "object") {
    if ("$cellType" in value) {
      return value.$cellType;
    }
    return "string";
  }
  return null;
};
var getRelationsPath = (partPath) => {
  const slash = partPath.lastIndexOf("/");
  const file = partPath.slice(slash + 1);
  return `_rels/${file}.rels`;
};
var isInRange = (c, r) => c.column >= r.columnStart && c.column <= r.columnEnd && c.row >= r.rowStart && c.row <= r.rowEnd;

// src/functions/get-offsets.ts
function getOffsets({
  cells,
  mergedCells,
  tables
}) {
  let rowOffset = 0;
  let initialRowOffset = 0;
  let columnOffset = 0;
  let currentRow = -1;
  const offsets = {
    row: {},
    col: {}
  };
  for (const cell of cells) {
    if (cell.inTable !== void 0) {
      const table = tables.get(cell.inTable);
      if (table === void 0) continue;
      if (cell.column === table.range.columnEnd && cell.row === table.range.rowStart) {
        const tableHeight = table.range.rowEnd - table.range.rowStart + 1;
        for (let i = 0; i < tableHeight; i++) {
          const row = cell.row + i;
          offsets.col[row] ??= {};
          offsets.col[row][table.range.columnEnd] = table.extension.cols;
        }
        table.finalize();
        continue;
      }
    }
    if (!Array.isArray(cell.newValue)) continue;
    if (currentRow !== cell.row) {
      initialRowOffset = rowOffset;
      currentRow = cell.row;
      columnOffset = 0;
    }
    if (offsets.col[currentRow] !== void 0) {
      for (const column of Object.keys(offsets.col[currentRow])) {
        const index = parseInt(column);
        if (index < cell.column) columnOffset = columnOffset = Math.max(columnOffset, offsets.col[currentRow][index]);
      }
    }
    let unit = 1;
    let rows = 1;
    if (cell.inMergedCell !== void 0) {
      const merge = mergedCells.get(cell.inMergedCell);
      if (merge !== void 0) {
        unit = "$isTable" in cell.newValue ? merge.height : merge.width;
        rows = merge.height;
      }
    }
    if ("$isTable" in cell.newValue) {
      const existingOffset = offsets.row[cell.row] ?? 0;
      const newOffset = initialRowOffset + cell.newValue.length * unit - unit;
      rowOffset = Math.max(newOffset, existingOffset ?? 0);
      offsets.row[cell.row] = rowOffset;
      continue;
    }
    const delta = cell.newValue.length * unit - unit;
    columnOffset = columnOffset + delta;
    for (let i = 0; i < rows; i++) {
      const row = cell.row + i;
      offsets.col[row] ??= {};
      offsets.col[row][cell.column] = columnOffset;
    }
  }
  return offsets;
}
var getRowShift = (cell, offsets) => {
  let shift = 0;
  for (const r in offsets.row) {
    const ri = Number(r);
    if (ri < cell.row) shift = offsets.row[ri];
  }
  return shift;
};
var getColumnShift = (cell, offsets) => {
  const rowOffsets = offsets.col;
  const colOffsets = rowOffsets[cell.row];
  if (colOffsets === void 0) return 0;
  let shift = 0;
  for (const c in colOffsets) {
    const ci = Number(c);
    if (ci < cell.column) shift = colOffsets[ci];
  }
  return shift;
};

// src/functions/xml-helpers.ts
var import_jszip = __toESM(require("jszip"), 1);
var readXlsx = async (buffer) => {
  const zip = new import_jszip.default();
  const content = await zip.loadAsync(buffer);
  return content;
};
var createXml = ({
  coreSchema,
  specificSchema,
  tagName
}) => {
  const doc = document.implementation.createDocument(specificSchema, tagName, null);
  const root = doc.documentElement;
  root.setAttributeNS("http://www.w3.org/2000/xmlns/", "xmlns", specificSchema);
  if (coreSchema !== void 0) {
    root.setAttributeNS("http://www.w3.org/2000/xmlns/", "xmlns:x", coreSchema);
  }
  return doc;
};
var serializer = new XMLSerializer();
var serializeXml = (doc) => {
  let serialized = serializer.serializeToString(doc);
  if (!serialized.startsWith("<?xml")) {
    serialized = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + serialized;
  }
  return serialized;
};
var valueToString = (value) => {
  if (value === void 0 || value === null) {
    return "";
  }
  if (Array.isArray(value)) {
    return value.map(valueToString).toString();
  }
  if (typeof value === "object") {
    if ("$cellType" in value) {
      return String(value);
    }
    return JSON.stringify(value);
  }
  if (typeof value === "number") {
    return value.toString();
  }
  return String(value);
};
var createSharedString = (text, doc = document) => {
  const si = doc.createElementNS(NS, "si");
  const t = doc.createElementNS(NS, "t");
  if (/^\s|\s$/.test(text)) {
    t.setAttributeNS(
      "http://www.w3.org/XML/1998/namespace",
      "xml:space",
      "preserve"
    );
  }
  t.textContent = text;
  si.append(t);
  return si;
};
var toType = {
  null: "s",
  error: "e",
  number: null,
  boolean: "b",
  string: "s",
  date: null,
  formula: null,
  url: "s"
};
var createCell = ({
  value,
  type,
  style,
  row,
  column,
  formula,
  xml
}) => {
  const c = xml.createElementNS(NS, "c");
  if (style !== null) c.setAttribute("s", style);
  const ref = toCellRef({ row, column });
  c.setAttribute("r", ref);
  if (type !== null) {
    const excelType = toType[type];
    if (excelType !== void 0 && excelType !== null) c.setAttribute("t", excelType);
  }
  if (formula !== null && formula !== void 0) {
    const f = xml.createElementNS(NS, "f");
    f.textContent = formula;
    c.append(f);
  }
  if (value !== null && value !== void 0) {
    const v = xml.createElementNS(NS, "v");
    v.textContent = String(value);
    c.append(v);
  }
  return c;
};

// src/functions/get-replacements.ts
var accessorRegex = /(?<=^\$\{)(?<table>\s*table:\s*)?(?<accessor>[^}]+?)(?=\s*}$)/;
var globalAccessorRegex = /\$\{(?<accessor>[^}]+)}/g;
var replaceAccessors = (text, dataStore) => {
  return text.replace(globalAccessorRegex, (_match, accessor) => {
    return valueToString(dataStore.get(accessor));
  });
};
var getReplacements = (oldValues, dataStore) => {
  const replacements = /* @__PURE__ */ new Map();
  let index = -1;
  for (const si of oldValues) {
    index++;
    if (si.children.length === 1 && si.children[0].tagName === "t") {
      const element = si.children[0];
      const text = element.textContent;
      const match = text.match(accessorRegex);
      if (match === null) {
        const value2 = replaceAccessors(text, dataStore);
        replacements.set(String(index), value2);
        continue;
      }
      if (match.groups === void 0) {
        replacements.set(String(index), text);
        continue;
      }
      const isTable = typeof match.groups.table === "string";
      const value = dataStore.get(match.groups.accessor);
      if (typeof value !== "object") {
        replacements.set(String(index), "");
        continue;
      }
      const cloned = structuredClone(value);
      if (isTable) Object.defineProperty(cloned, "$isTable", { value: true, enumerable: false });
      replacements.set(String(index), cloned);
      continue;
    }
    for (const t of Array.from(si.querySelectorAll("t"))) {
      const text = t.textContent ?? "";
      t.textContent = replaceAccessors(text, dataStore);
      replacements.set(String(index), si);
    }
  }
  return replacements;
};

// src/functions/process-worksheets.ts
var processWorksheets = async ({
  workbook,
  dataStore
}) => {
  const {
    styles,
    sharedStrings,
    dataValidations,
    pivotTables,
    extensionLists
  } = workbook;
  const oldValues = sharedStrings.oldValues;
  const replacements = getReplacements(oldValues, dataStore);
  for (const worksheet of workbook.sheets) {
    const {
      name,
      cells,
      mergedCells,
      hyperlinks,
      comments,
      conditionalFormattings,
      tables,
      xml,
      replaceCells
    } = worksheet;
    comments.changeComments(dataStore);
    const processedCells = cells.filter((cell) => {
      if (cell.inMergedCell !== void 0) {
        if ((cell.value === void 0 || cell.value === null) && (cell.formula === void 0 || cell.formula === null)) return false;
      }
      return true;
    }).map((cell) => {
      if (cell.type !== "string") return cell;
      if (cell.value === void 0) return cell;
      if (cell.value === null) return cell;
      cell.newValue = replacements.get(String(cell.value));
      if (cell.inTable !== void 0) {
        const table = tables.get(cell.inTable);
        table?.extend(cell);
      }
      if (cell.inConditionalFormatting !== void 0) {
        const formatting = conditionalFormattings.get(cell.inConditionalFormatting);
        formatting?.extend(cell);
      }
      if (cell.inDataValidation !== void 0) {
        const dataValidation = dataValidations.get(cell.inDataValidation);
        dataValidation?.extend(cell);
      }
      if (cell.inDataValidationList !== void 0) {
        const dataValidation = dataValidations.get(cell.inDataValidationList);
        dataValidation?.extendFormula(cell);
      }
      if (cell.inExtensionList !== void 0) {
        const extension = extensionLists.get(cell.inExtensionList);
        extension?.extend(cell);
      }
      if (cell.inExtensionListSource !== void 0) {
        const extension = extensionLists.get(cell.inExtensionListSource);
        extension?.extendFormula(cell);
      }
      if (cell.inPivotTableSource !== void 0) {
        const pivotTableSource = pivotTables.get(cell.inPivotTableSource);
        pivotTableSource?.extend(cell);
      }
      return cell;
    });
    processedCells.sort((cell1, cell2) => {
      if (cell1.row !== cell2.row) return cell1.row - cell2.row;
      return cell1.column - cell2.column;
    });
    const offsets = getOffsets({
      cells: processedCells,
      mergedCells,
      tables
    });
    const movedCells = processedCells.map((cell) => {
      const colOffset = getColumnShift(cell, offsets);
      const rowOffset = getRowShift(cell, offsets);
      cell.column = cell.column + colOffset;
      cell.row = cell.row + rowOffset;
      return cell;
    });
    const extendedCells = [];
    for (const cell of movedCells) {
      if (cell.type !== "string" || !Array.isArray(cell.newValue)) {
        extendedCells.push(cell);
        continue;
      }
      if (Array.isArray(cell.newValue)) {
        const mergedCell = mergedCells.get(cell.inMergedCell ?? "");
        let width = 1;
        let height = 1;
        if (mergedCell !== void 0) {
          width = mergedCell.width;
          height = mergedCell.height;
        }
        for (let i = 0; i < cell.newValue.length; i++) {
          const newCell = { ...cell };
          newCell.newValue = cell.newValue[i];
          newCell.column = newCell.column + ("$isTable" in cell.newValue ? 0 : width) * i;
          newCell.row = newCell.row + ("$isTable" in cell.newValue ? height : 0) * i;
          if (mergedCell !== void 0) {
            mergedCells.add({
              columnStart: newCell.column,
              rowStart: newCell.row,
              columnEnd: newCell.column + width - 1,
              rowEnd: newCell.row + height - 1
            });
          }
          extendedCells.push(newCell);
        }
      }
    }
    const newCells = extendedCells.map((cell) => {
      if (isElement(cell.newValue)) {
        const sharedValue = sharedStrings.get(cell.newValue);
        cell.value = sharedValue;
        return cell;
      }
      if (cell.type !== "string") {
        if (cell.type === "formula") cell.value = null;
        return cell;
      }
      const newValue = cell.newValue === void 0 ? cell.value : cell.newValue;
      const cellType = cell.type === "string" ? guessDataType(newValue) : cell.type;
      const value = convertValueToExcel({ value: newValue, cellType });
      cell.type = cellType;
      if ((cellType === "formula" || cellType === "error") && typeof value === "string") {
        if (cell.formula === void 0) cell.formula = value;
        cell.type = "formula";
        cell.value = null;
        return cell;
      }
      if (cellType === "string" && typeof value === "string") {
        const sharedValue = sharedStrings.get(value);
        cell.value = sharedValue;
        return cell;
      }
      if (cellType === "url" && typeof value === "string") {
        cell.style = styles.getHyperlinkFormat();
        const sharedValue = sharedStrings.get(value);
        cell.value = sharedValue;
        hyperlinks.add({ range: cell, url: value });
        return cell;
      }
      if (cellType === "date") {
        cell.style = styles.getDateFormat();
      }
      cell.value = value;
      return cell;
    });
    const rowsMap = /* @__PURE__ */ new Map();
    for (const cell of newCells) {
      let row = rowsMap.get(cell.row);
      if (row === void 0) {
        row = xml.createElementNS(NS, "row");
        row.setAttribute("r", String(cell.row));
        rowsMap.set(cell.row, row);
        row = rowsMap.get(cell.row);
      }
      row?.append(createCell({ ...cell, xml }));
    }
    const original = xml.querySelector("sheetData");
    if (original === null) throw new Error("Excel workbook is malformed");
    const newSheetData = original.cloneNode();
    if (newSheetData === void 0) throw new Error("Excel workbook is mailformed");
    for (const row of rowsMap.values()) {
      newSheetData.append(row);
    }
    const validations = dataStore.get("$validations");
    if (Array.isArray(validations)) {
      for (const validation of validations) {
        if (validation.sheet !== name) continue;
        dataValidations.add(xml, validation);
      }
    }
    const currentComments = dataStore.get("$comments");
    if (Array.isArray(currentComments)) {
      for (const comment of currentComments) {
        if (comment.sheet !== name) continue;
        comments.add(comment);
      }
    }
    replaceCells(newSheetData);
    comments.save();
    conditionalFormattings.save();
    tables.save();
    hyperlinks.save();
  }
  dataValidations.save();
  extensionLists.save();
  pivotTables.save();
  for (const worksheet of workbook.sheets) worksheet.save();
};

// src/functions/global-helpers.ts
var parser = new DOMParser();
var generateUUID = () => {
  return crypto.randomUUID();
};

// src/functions/get-styles.ts
var DATE_FORMAT = String.raw`yyyy\-mm\-dd;@`;
var serializer2 = new XMLSerializer();
var hasAtLeastTwoDateParts = (s) => {
  let count = 0;
  const lower = s.toLowerCase();
  for (const p of ["y", "m", "d", "h"]) {
    if (lower.includes(p) && ++count >= 2) return true;
  }
  return false;
};
var createStyles = async (zip, target) => {
  let numFmts;
  let numFmtId;
  let cachedDateFormatId = null;
  let cachedLinkFormatId = null;
  let isDirty = false;
  const xmlStyle = await zip.file(`xl/${target}`)?.async("string");
  if (xmlStyle === void 0) throw new Error("This Excel file is corrupted: styles.xml missing");
  const domStyle = parser.parseFromString(xmlStyle, "application/xml");
  const styleSheet = domStyle.querySelector("styleSheet");
  const cellXfs = domStyle.querySelector("cellXfs");
  if (styleSheet === null || cellXfs === null) throw new Error("Invalid styles.xml structure");
  numFmts = domStyle.querySelector("numFmts");
  if (numFmts === null) {
    numFmts = domStyle.createElementNS(NS, "numFmts");
    styleSheet.prepend(numFmts);
  }
  const existingIds = Array.from(
    domStyle.querySelectorAll("numFmt"),
    (el) => Number(el.getAttribute("numFmtId"))
  ).filter(Number.isFinite);
  numFmtId = Math.max(163, ...existingIds);
  const save = () => {
    if (!isDirty) return;
    const newXfData = serializer2.serializeToString(domStyle);
    zip.file(`xl/${target}`, newXfData);
  };
  const getDateFormat = () => {
    const existingDateFormat = Array.from(
      domStyle.querySelectorAll("cellXfs xf")
    ).findIndex((el) => {
      return ["14", "15", "16", "17", "18", "22"].includes(el.getAttribute("numFmtId") ?? "") || hasAtLeastTwoDateParts(el.getAttribute("formatCode") ?? "");
    });
    if (existingDateFormat !== -1) {
      cachedDateFormatId = String(existingDateFormat);
      return cachedDateFormatId;
    }
    numFmtId++;
    const numFmt = domStyle.createElementNS(NS, "numFmt");
    numFmt.setAttribute("numFmtId", String(numFmtId));
    numFmt.setAttribute("formatCode", DATE_FORMAT);
    if (numFmts === null || cellXfs === null) throw new Error("Number formats are mailformed");
    numFmts.append(numFmt);
    numFmts.setAttribute("count", String(numFmts.children.length));
    const xf = domStyle.createElementNS(NS, "xf");
    xf.setAttribute("numFmtId", String(numFmtId));
    xf.setAttribute("fontId", "0");
    xf.setAttribute("fillId", "0");
    xf.setAttribute("borderId", "0");
    xf.setAttribute("xfId", "0");
    xf.setAttribute("applyNumberFormat", "1");
    cellXfs.append(xf);
    cellXfs.setAttribute("count", String(cellXfs.children.length));
    isDirty = true;
    cachedDateFormatId = String(cellXfs.children.length - 1);
    return cachedDateFormatId;
  };
  const getHyperlinkFormat = () => {
    const fonts = domStyle.querySelector("fonts");
    if (fonts === null || cellXfs === null) {
      throw new Error("styles.xml malformed");
    }
    let fontId = Array.from(fonts.children).findIndex((font) => {
      return font.querySelector("u") !== void 0 && font.querySelector("color")?.getAttribute("theme") === "10";
    });
    if (fontId === -1) {
      const font = domStyle.createElementNS(NS, "font");
      const color = domStyle.createElementNS(NS, "color");
      color.setAttribute("theme", "10");
      const underline = domStyle.createElementNS(NS, "u");
      font.append(color, underline);
      fonts.append(font);
      fonts.setAttribute("count", String(fonts.children.length));
      fontId = fonts.children.length - 1;
    }
    const xfIndex = Array.from(cellXfs.children).findIndex(
      (xf2) => xf2.getAttribute("fontId") === String(fontId)
    );
    if (xfIndex !== -1) return String(xfIndex);
    const xf = domStyle.createElementNS(NS, "xf");
    xf.setAttribute("fontId", String(fontId));
    xf.setAttribute("fillId", "0");
    xf.setAttribute("borderId", "0");
    xf.setAttribute("xfId", "0");
    xf.setAttribute("applyFont", "1");
    cellXfs.append(xf);
    cellXfs.setAttribute("count", String(cellXfs.children.length));
    isDirty = true;
    cachedLinkFormatId = String(cellXfs.children.length - 1);
    return cachedLinkFormatId;
  };
  return { save, getDateFormat, getHyperlinkFormat };
};

// src/functions/get-shared-strings.ts
var getSharedStrings = async (xlsx, target) => {
  let isDirty = false;
  const filename = `xl/${target}`;
  const xmlText = await xlsx.file(filename)?.async("string");
  if (xmlText === void 0) {
    throw new Error("This Excel template has no strings");
  }
  const xml = parser.parseFromString(xmlText, "application/xml");
  const sst = xml.querySelector("sst");
  if (sst === null) throw new Error("This Excel template has no strings");
  const oldValues = Array.from(sst.children);
  const strings = /* @__PURE__ */ new Map();
  const elements = /* @__PURE__ */ new WeakMap();
  const list = [];
  const get = (el) => {
    if (typeof el === "string") {
      const value2 = strings.get(el);
      if (value2 !== void 0) {
        return value2;
      }
      const newString = createSharedString(el, xml);
      const index2 = String(list.length);
      strings.set(el, index2);
      list.push(newString);
      isDirty = true;
      return index2;
    }
    const value = elements.get(el);
    if (value !== void 0) {
      return value;
    }
    const index = String(list.length);
    elements.set(el, index);
    list.push(el);
    isDirty = true;
    return index;
  };
  const save = () => {
    if (!isDirty) return;
    sst?.replaceChildren(...list);
    const text = serializeXml(xml);
    xlsx.file(filename, text);
  };
  return {
    oldValues,
    get,
    save
  };
};

// src/functions/cells-helpers.ts
var typesDict = {
  f: "formula",
  e: "error",
  b: "boolean",
  s: "string",
  n: "number"
};
var cellToModel = (cell) => {
  const currentType = cell.getAttribute("t");
  const cellType = typesDict[currentType ?? "n"];
  const cellRef = cell.getAttribute("r");
  const cellStyle = cell.getAttribute("s");
  let value = cell.querySelector("v")?.textContent;
  if (cellType === "number" && typeof value === "string") value = parseFloat(value);
  const formula = cell.querySelector("f")?.textContent;
  if (cellRef === null) {
    throw new Error("Excel template is mailformed");
  }
  const { row, column } = parseCellRef(cellRef);
  return {
    ref: cellRef,
    row,
    column,
    style: cellStyle,
    type: cellType,
    ...formula !== void 0 ? { formula } : {},
    ...value !== void 0 ? { value } : {}
  };
};

// src/functions/get-cells.ts
var getCells = (xml) => {
  const sheetData = xml.querySelector("sheetData") ?? xml.createElementNS(NS, "sheetData");
  const cellsXml = sheetData.querySelectorAll("c");
  const cells = cellsXml === void 0 ? [] : Array.from(cellsXml).map((c) => cellToModel(c));
  return cells;
};

// src/functions/get-merged-cells.ts
var getMergedCells = (xml) => {
  const mergedCells = /* @__PURE__ */ new Map();
  const mergeCellsParent = xml.querySelector("mergeCells") ?? xml.createElementNS(NS, "mergeCells");
  const mergeCellsXml = mergeCellsParent.querySelectorAll("mergeCell");
  for (const mergedCell of Array.from(mergeCellsXml)) {
    const ref = mergedCell.getAttribute("ref");
    if (ref === null) continue;
    const range = parseCellRange(ref);
    mergedCells.set(ref, {
      range,
      width: range.columnEnd - range.columnStart + 1,
      height: range.rowEnd - range.rowStart + 1,
      ref,
      xml: mergedCell
    });
  }
  mergeCellsParent.replaceChildren();
  const add = (ref) => {
    const mergedCell = xml.createElementNS(NS, "mergeCell");
    mergedCell.setAttribute("ref", toCellRange(ref));
    mergeCellsParent.appendChild(mergedCell);
  };
  const get = (ref) => {
    return mergedCells.get(ref);
  };
  const findByCell = (cell) => {
    for (const [name, value] of mergedCells) {
      if (isInRange(cell, value.range)) return name;
    }
    return null;
  };
  return {
    add,
    get,
    findByCell
  };
};

// src/functions/get-hyperlinks.ts
var addHyperlink = (doc, hyperlinks, ref, rel) => {
  const id = generateUUID();
  const el = doc.createElementNS(NS, "hyperlink");
  el.setAttribute("ref", ref);
  el.setAttribute("r:id", rel);
  el.setAttribute("xr:uid", `{${id.toUpperCase()}}`);
  hyperlinks.appendChild(el);
};
var getHyperlinks = ({
  xml,
  relations
}) => {
  let exists = true;
  let hyperLinksParent = xml.querySelector("hyperlinks");
  if (hyperLinksParent === null) {
    hyperLinksParent = xml.createElementNS(NS, "hyperlinks");
    exists = false;
  }
  const hyperLinksXml = hyperLinksParent.querySelectorAll("hyperlink");
  const existingLinks = /* @__PURE__ */ new Map();
  for (const hyperlink of hyperLinksXml) {
    const ref = hyperlink.getAttribute("ref");
    const rId = hyperlink.getAttribute("r:id");
    const xrUid = hyperlink.getAttribute("xr:uid");
    if (ref === null || rId === null || xrUid === null) continue;
    existingLinks.set(ref, {
      ref,
      "r:id": rId,
      "xr:uid": xrUid,
      xml: hyperlink
    });
  }
  const save = () => {
    if (exists) return;
    const pageMargins = xml.querySelector("pageMargins");
    if (pageMargins !== null) pageMargins.parentNode?.insertBefore(hyperLinksParent, pageMargins);
  };
  const add = ({ range, url }) => {
    let id;
    const ref = toCellRef(range);
    const oldUrlElemnt = xml.querySelector(`Relationship[Target="${url}"]`);
    if (oldUrlElemnt !== null) {
      const oldUrl = oldUrlElemnt.getAttribute("Id");
      if (oldUrl !== null) id = oldUrl;
    } else {
      id = relations.add({
        type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        target: url,
        targetMode: "External"
      });
    }
    if (id === void 0) throw new Error("The URL is mailformed");
    addHyperlink(xml, hyperLinksParent, ref, id);
  };
  return {
    add,
    save
  };
};

// src/functions/get-comments.ts
var addThreadedComment = (doc, c) => {
  const id = crypto.randomUUID();
  const el = doc.createElementNS(NS_TC, "threadedComment");
  el.setAttribute("ref", c.ref);
  el.setAttribute("dT", (/* @__PURE__ */ new Date()).toISOString());
  el.setAttribute("personId", c.personId ?? "");
  el.setAttribute("id", "{" + id.toUpperCase() + "}");
  const text = doc.createElementNS(NS_TC, "text");
  text.textContent = c.text;
  el.appendChild(text);
  doc.documentElement.appendChild(el);
  return id;
};
var getComments = async ({
  xlsx,
  relations,
  workbook,
  id
}) => {
  let isDirty = false;
  const guessedFilename = `../threadedComments/threadedComment${id.replace("rId", "")}.xml`;
  const commentsRelations = relations.getElements("threadedComment");
  if (commentsRelations.length > 1) throw new Error("Comments are mailformed");
  let relationExists = commentsRelations.length > 0;
  const commentsFileName = (commentsRelations[0]?.target ?? guessedFilename).replace("..", "xl");
  const xmlText = await xlsx.file(commentsFileName)?.async("string");
  const xml = xmlText === void 0 ? createXml({ coreSchema: NS, specificSchema: NS_TC, tagName: "threadedComments" }) : parser.parseFromString(xmlText, "application/xml");
  if (xmlText === void 0) isDirty = true;
  if (xml.getElementsByTagName("parsererror").length > 0) {
    throw new Error(`Invalid relations file for ${commentsFileName}`);
  }
  const oldCommentsMap = /* @__PURE__ */ new Map();
  const commentsXml = xml.querySelectorAll("threadedComment");
  for (const comment of commentsXml) {
    const ref = comment.getAttribute("ref");
    const dT = comment.getAttribute("dT");
    const personId = comment.getAttribute("personId");
    const id2 = comment.getAttribute("id");
    const text = comment.querySelector("text")?.textContent;
    if (ref === null || dT === null || personId === null || id2 === null) continue;
    oldCommentsMap.set(ref, {
      ref,
      dT,
      personId,
      id: id2,
      text: text ?? "",
      xml: comment
    });
  }
  const changeComments = (datastore) => {
    for (const [, comment] of oldCommentsMap) {
      comment.text = replaceAccessors(comment.text, datastore);
      const text = comment.xml.querySelector("text");
      if (text === null) continue;
      text.textContent = comment.text;
      isDirty = true;
    }
  };
  let oldComments;
  let drawings;
  const setOldComments = (props) => {
    oldComments = props;
  };
  const setDrawings = (props) => {
    drawings = props;
  };
  const add = ({ ref, text, row, column }) => {
    addRelation();
    if (ref === void 0) {
      if (row === void 0 || column === void 0) {
        throw new Error("Comment position is undefined");
      }
      ref = toCellRef({ row, column });
    }
    if (row === void 0 && column === void 0) {
      if (ref === void 0) {
        throw new Error("Comment position is undefined");
      }
      const location = parseCellRef(ref);
      row = location.row;
      column = location.column;
    }
    const systemPersonId = "{00000000-0000-0000-0000-000000000000}";
    const id2 = addThreadedComment(xml, {
      personId: systemPersonId,
      ref,
      text
    });
    oldComments?.add({ id: `{${id2.toUpperCase()}}`, ref, text });
    drawings?.add({ row: String(row), column: String(column) });
    isDirty = true;
  };
  const getComment = (ref) => {
    return oldCommentsMap.get(ref)?.ref;
  };
  const addRelation = () => {
    if (relationExists) return;
    const commentsRelation = {
      type: "http://schemas.microsoft.com/office/2017/10/relationships/threadedComment",
      target: `../threadedComments/threadedComment${id.replace("rId", "")}.xml`
    };
    relations.add(commentsRelation);
    relationExists = true;
  };
  const save = () => {
    if (!isDirty) return;
    if (oldComments !== void 0) oldComments.save();
    if (drawings !== void 0) drawings.save();
    const text = serializeXml(xml);
    xlsx.file(commentsFileName, text);
  };
  return {
    add,
    changeComments,
    getComment,
    setDrawings,
    setOldComments,
    save
  };
};

// src/functions/get-conditional-formatting.ts
var getConditionalFormatting = (xml) => {
  const conditionalFormattingXml = xml.querySelectorAll("conditionalFormatting");
  const conditionalFormattings = /* @__PURE__ */ new Map();
  let index = -1;
  for (const formatting of conditionalFormattingXml) {
    const sqref = formatting.getAttribute("sqref");
    const formula = formatting.querySelector("formula")?.textContent;
    if (sqref === null) continue;
    index++;
    const range = parseCellRange(sqref);
    const extension = {
      rows: 0,
      cols: 0
    };
    let columnExtension = 0;
    let rowExtension = 0;
    let currentRow = 0;
    const extend = (cell) => {
      if (Array.isArray(cell.newValue)) {
        if (cell.row !== currentRow) {
          currentRow = cell.row;
          columnExtension = 0;
          rowExtension = 0;
        }
        const length = cell.newValue.length - 1;
        if ("$isTable" in cell.newValue) {
          rowExtension = Math.max(length, rowExtension) - Math.min(length, rowExtension);
          extension.rows = extension.rows + rowExtension;
        } else {
          columnExtension = columnExtension + length;
          extension.cols = Math.max(columnExtension, extension.cols);
        }
      }
    };
    const currentFormatting = {
      sqref,
      range,
      extension,
      extend,
      ...formula === void 0 ? {} : { formula },
      xml: formatting
    };
    conditionalFormattings.set(String(index), currentFormatting);
  }
  const save = () => {
    for (const formatting of conditionalFormattings.values()) {
      formatting.range.columnEnd = formatting.range.columnEnd + formatting.extension.cols;
      formatting.range.rowEnd = formatting.range.rowEnd + formatting.extension.rows;
      const ref = toCellRange(formatting.range);
      formatting.xml.setAttribute("sqref", ref);
    }
  };
  const get = (ref) => {
    return conditionalFormattings.get(ref);
  };
  const findByCell = (cell) => {
    for (const [name, value] of conditionalFormattings) {
      if (isInRange(cell, value.range)) return name;
    }
    return null;
  };
  return {
    get,
    save,
    findByCell
  };
};

// src/functions/get-tables.ts
var serializer3 = new XMLSerializer();
var getTables = async ({
  xlsx,
  relations
}) => {
  const tablesProps = relations.getElements("table");
  const tableFiles = tablesProps.map((el) => el.target);
  const tables = /* @__PURE__ */ new Map();
  for (let file of tableFiles) {
    file = file.replace(/^\.\./, "xl");
    const xmlText = await xlsx.file(file)?.async("string");
    if (xmlText === void 0) {
      continue;
    }
    const xml = parser.parseFromString(xmlText, "application/xml");
    const table = xml.querySelector("table");
    if (table === null) continue;
    const ref = table.getAttribute("ref");
    if (ref === null) continue;
    const range = parseCellRange(ref);
    const headers = Array.from(xml.querySelectorAll("tableColumn")).map((el) => el.getAttribute("name")).filter((el) => el !== null);
    const initialHeadersLength = headers.length;
    const extension = {
      rows: 0,
      cols: 0
    };
    const existingHeaders = new Set(headers);
    const addHeader = (name, index) => {
      let newName = name;
      let j = 0;
      while (existingHeaders.has(newName)) {
        j++;
        newName = name + String(j);
      }
      existingHeaders.add(newName);
      headers.splice(index, 0, newName);
      return newName;
    };
    const createHeaders = (names) => {
      return names.map((name, i) => {
        const h = xml.createElementNS(NS, "tableColumn");
        h.setAttribute("id", String(i + 1));
        h.setAttribute("name", name);
        h.setAttribute("xr3:uid", `{${generateUUID().toUpperCase()}}`);
        return h;
      });
    };
    let columnExtension = 0;
    let rowExtension = 0;
    let currentRow = 0;
    const extend = (cell) => {
      const t = tables.get(file);
      if (t === void 0) return;
      if (Array.isArray(cell.newValue)) {
        const isHeader = cell.row === range.rowStart;
        if (isHeader) {
          const index = cell.column - range.columnStart;
          const removed = headers.splice(index, 1);
          existingHeaders.delete(removed[0]);
          cell.newValue = cell.newValue.map((header, i) => {
            return addHeader(header, i + index);
          });
        }
      }
      if (Array.isArray(cell.newValue)) {
        if (cell.row !== currentRow) {
          currentRow = cell.row;
          columnExtension = 0;
          rowExtension = 0;
        }
        const length = cell.newValue.length - 1;
        if ("$isTable" in cell.newValue) {
          rowExtension = Math.max(length, rowExtension) - Math.min(length, rowExtension);
          t.extension.rows = t.extension.rows + rowExtension;
        } else {
          columnExtension = columnExtension + length;
          t.extension.cols = Math.max(columnExtension, t.extension.cols);
        }
      }
      t.isDirty = true;
    };
    const finalize = () => {
      const headersLength = headers.length;
      const headersDifference = range.columnEnd - range.columnStart + extension.cols + 1 - headersLength;
      if (headersDifference !== 0) {
        for (let i = 0; i < headersDifference; i++) {
          addHeader("Column", range.columnEnd + i);
        }
        if (tableInMap.lastHeaderCell !== void 0) {
          tableInMap.lastHeaderCell.newValue = headers.slice(headersLength - 1);
        }
      }
      const newHeaders = createHeaders(headers);
      xml.querySelector("tableColumns")?.replaceChildren(...newHeaders);
      range.columnEnd = range.columnEnd + extension.cols;
      range.rowEnd = range.rowEnd + extension.rows;
      const ref2 = toCellRange(range);
      table.setAttribute("ref", ref2);
      const autofilter = table.querySelector("autoFilter");
      if (autofilter !== null) autofilter.setAttribute("ref", ref2);
    };
    const save2 = () => {
      const newData = serializer3.serializeToString(xml);
      xlsx.file(file, newData);
    };
    const tableInMap = {
      xml,
      ref,
      range,
      extend,
      extension,
      finalize,
      save: save2,
      isDirty: false
    };
    tables.set(file, tableInMap);
  }
  const get = (tableName) => {
    return tables.get(tableName);
  };
  const save = () => {
    for (const value of tables.values()) {
      if (value.isDirty) value.save();
    }
  };
  const findByCell = (cell) => {
    for (const [name, table] of tables) {
      if (cell.column === table.range.columnEnd && cell.row === table.range.rowStart) {
        table.lastHeaderCell = cell;
      }
      if (isInRange(cell, table.range)) return name;
    }
    return null;
  };
  return {
    get,
    save,
    findByCell
  };
};

// src/functions/get-drawings.ts
var serializeXml2 = (content) => `
<xml xmlns:v="urn:schemas-microsoft-com:vml"
     xmlns:o="urn:schemas-microsoft-com:office:office"
     xmlns:x="urn:schemas-microsoft-com:office:excel">

  <o:shapelayout v:ext="edit">
    <o:idmap v:ext="edit" data="1"/>
  </o:shapelayout>

  <v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe">
    <v:stroke joinstyle="miter"/>
    <v:path gradientshapeok="t" o:connecttype="rect"/>
  </v:shapetype>
  ${content}
</xml>
`;
var createShape = ({
  id,
  row,
  column
}) => {
  return `<v:shape id="_x0000_s${String(1025 + id)}" type="#_x0000_t202"
  style="position:absolute;width:10pt;height:10pt;visibility:hidden"
  fillcolor="none" strokecolor="none">
  <v:fill color2="infoBackground [80]"/>
  <v:shadow color="none [81]" obscured="t"/>
  <v:path o:connecttype="none"/>
  <v:textbox style="mso-direction-alt:auto">
  <div style="text-align:left"/>
  </v:textbox>
  <x:ClientData ObjectType="Note">
  <x:MoveWithCells/>
  <x:SizeWithCells/>
  <x:Anchor>${column},0,${row},0,${column + 1},0,${row + 2},0</x:Anchor>
  <x:AutoFill>False</x:AutoFill>
  <x:Row>${row}</x:Row>
  <x:Column>${column}</x:Column>
  </x:ClientData>
  </v:shape>`;
};
var getDrawings = async ({
  id,
  relations,
  xlsx
}) => {
  const drawingRelation = relations.getElements("vmlDrawing");
  const guessedFilename = `../drawings/vmlDrawing${id.replace("rId", "")}.vml`;
  const noRelation = drawingRelation.length === 0;
  const filename = (drawingRelation[0]?.target ?? guessedFilename).replace("..", "xl");
  const unique = /* @__PURE__ */ new Set();
  let isDirty = false;
  const drawings = [];
  const xmlText = await xlsx.file(filename)?.async("string");
  if (xmlText !== void 0) {
    const vmlDoc = parser.parseFromString(xmlText, "application/xml");
    const shapes = vmlDoc.getElementsByTagName("v:shape");
    for (const shape of shapes) {
      const row = shape.getElementsByTagName("x:Row")[0]?.textContent;
      const column = shape.getElementsByTagName("x:Column")[0]?.textContent;
      if (row === void 0 || column === void 0) continue;
      const rowValue = String(parseInt(row) + 1);
      const columnValue = String(parseInt(column) + 1);
      drawings.push({ row: rowValue, column: columnValue });
    }
  }
  const add = ({ row, column }) => {
    const ref = `${row}:${column}`;
    if (unique.has(ref)) return;
    drawings.push({ row, column });
    unique.add(ref);
    isDirty = true;
  };
  const save = () => {
    if (!isDirty) return;
    if (noRelation) {
      relations.add({
        target: filename.replace("xl", ".."),
        type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing"
      });
    }
    const content = drawings.map((drawing, i) => {
      return createShape({
        id: i,
        row: parseInt(drawing.row) - 1,
        column: parseInt(drawing.column) - 1
      });
    }).join("");
    const data = serializeXml2(content);
    xlsx.file(filename, data);
  };
  return {
    add,
    save
  };
};

// src/functions/get-old-comments.ts
var EXCEL_NOTIFICATION = `[Threaded comment]

Your version of Excel allows you to read this threaded comment; however, any edits to it will get removed if the file is opened in a newer version of Excel. Learn more: https://go.microsoft.com/fwlink/?linkid=870924

Comment:
    `;
var NS_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";
var NS_XR = "http://schemas.microsoft.com/office/spreadsheetml/2014/revision";
var createComment = ({ xml, ref, id, text, author }) => {
  const comment = xml.createElementNS(NS, "comment");
  comment.setAttribute("ref", ref);
  comment.setAttribute("xr:uid", id);
  comment.setAttribute("authorId", author);
  comment.setAttribute("shapeId", "0");
  const textTag = xml.createElementNS(NS, "text");
  const tTag = xml.createElementNS(NS, "t");
  tTag.textContent = text;
  textTag.appendChild(tTag);
  comment.appendChild(textTag);
  return comment;
};
var createAuthor = ({ xml, id }) => {
  const author = xml.createElementNS(NS, "author");
  author.textContent = `tc=${id}`;
  return author;
};
var getOldComments = async ({
  id,
  relations,
  xlsx
}) => {
  const oldCommentsRelation = relations.getElements("comments");
  const guessedFilename = `../comments${id.replace("rId", "")}.xml`;
  const noRelation = oldCommentsRelation.length === 0;
  const filename = (oldCommentsRelation[0]?.target ?? guessedFilename).replace("..", "xl");
  let isDirty = false;
  const comments = [];
  const xmlText = await xlsx.file(filename)?.async("string");
  const xml = xmlText === void 0 ? createXml({ specificSchema: NS, tagName: "comments" }) : parser.parseFromString(xmlText, "application/xml");
  if (xmlText === void 0) {
    xml.querySelector("comments")?.setAttribute("xmlns:mc", NS_MC);
    xml.querySelector("comments")?.setAttribute("xmlns:xr", NS_XR);
    xml.querySelector("comments")?.setAttribute("mc:Ignorable", "xr");
    xml.querySelector("comments")?.appendChild(xml.createElementNS(NS, "authors"));
    xml.querySelector("comments")?.appendChild(xml.createElementNS(NS, "commentList"));
  }
  const parent = xml.querySelector("comments");
  const authors = xml.querySelector("authors");
  const commentList = xml.querySelector("commentList");
  if (parent === null || authors === null || commentList === null) throw new Error("Comments are mailformed");
  const uniqueRefs = /* @__PURE__ */ new Set();
  if (xmlText !== void 0) {
    const currentComments = commentList.getElementsByTagName("comment");
    for (const comment of currentComments) {
      const ref = comment.getAttribute("ref");
      const id2 = comment.getAttribute("xr:uid");
      let text = comment.querySelector("t")?.textContent ?? "";
      text = text.replace(EXCEL_NOTIFICATION, "");
      if (ref === null || id2 === null) continue;
      comments.push({ ref, id: id2, text });
    }
  }
  const add = ({ ref, id: id2, text }) => {
    if (uniqueRefs.has(ref)) return;
    uniqueRefs.add(ref);
    comments.push({ ref, id: id2, text });
    isDirty = true;
  };
  const save = () => {
    if (!isDirty) return;
    if (noRelation) {
      relations.add({
        target: filename.replace("xl", ".."),
        type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
      });
    }
    const newComments = [];
    const newAuthors = [];
    let index = 0;
    for (const comment of comments) {
      const newComment = createComment({ xml, author: String(index), ...comment });
      const author = createAuthor({ xml, id: comment.id });
      newComments.push(newComment);
      newAuthors.push(author);
      index++;
    }
    authors.replaceChildren(...newAuthors);
    commentList?.replaceChildren(...newComments);
    const data = serializeXml(xml);
    xlsx.file(filename, data);
  };
  return {
    add,
    save
  };
};

// src/functions/get-sheets.ts
var getSheet = async ({
  xlsx,
  xml,
  relations,
  workbook,
  target,
  id,
  name
}) => {
  const filename = `xl/${target}`;
  const {
    pivotTables,
    dataValidations,
    extensionLists
  } = workbook;
  const initialCells = getCells(xml);
  const mergedCells = getMergedCells(xml);
  const hyperlinks = getHyperlinks({ xml, relations });
  const comments = await getComments({ xlsx, workbook, relations, id });
  const conditionalFormattings = getConditionalFormatting(xml);
  const tables = await getTables({ xlsx, relations });
  const drawings = await getDrawings({ xlsx, id, relations });
  const oldComments = await getOldComments({ xlsx, id, relations });
  comments.setDrawings(drawings);
  comments.setOldComments(oldComments);
  const cells = initialCells.map((cell) => {
    const comment = comments.getComment(cell.ref);
    if (comment !== void 0) cell.hasComment = comment;
    const table = tables.findByCell(cell);
    if (table !== null) cell.inTable = table;
    const mergeCell = mergedCells.findByCell(cell);
    if (mergeCell !== null) cell.inMergedCell = mergeCell;
    const pivotTable = pivotTables.findByCell(name, cell);
    if (pivotTable !== null) cell.inPivotTableSource = pivotTable;
    const dataValidation = dataValidations.findByCell(name, cell);
    if (dataValidation !== null) cell.inDataValidation = dataValidation;
    const dataValidationList = dataValidations.findListByCell(name, cell);
    if (dataValidationList !== null) cell.inDataValidationList = dataValidationList;
    const conditionalFormatting = conditionalFormattings.findByCell(cell);
    if (conditionalFormatting !== null) cell.inConditionalFormatting = conditionalFormatting;
    const extension = conditionalFormattings.findByCell(cell);
    if (extension !== null) cell.inExtensionList = extension;
    const extensionListSource = extensionLists.findListByCell(name, cell);
    if (extensionListSource !== null) cell.inExtensionListSource = extensionListSource;
    return cell;
  });
  const replaceCells = (element) => {
    xml.querySelector("sheetData")?.replaceWith(element);
  };
  const cleanUp = () => {
    const hyperlinks2 = xml.querySelector("hyperlinks");
    if (hyperlinks2?.childNodes.length === 0) hyperlinks2.remove();
    const mergeCells = xml.querySelector("mergeCells");
    if (mergeCells?.childNodes.length === 0) mergeCells.remove();
  };
  const save = () => {
    cleanUp();
    const newData = serializeXml(xml);
    relations.save();
    xlsx.file(filename, newData);
  };
  return {
    xml,
    sheetId: id,
    relations,
    target,
    id,
    name,
    cells,
    mergedCells,
    hyperlinks,
    comments,
    dataValidations,
    extensionLists,
    conditionalFormattings,
    drawings,
    oldComments,
    tables,
    replaceCells,
    save
  };
};

// src/functions/get-relations.ts
var getRelations = async ({ xlsx, filename }) => {
  let isDirty = false;
  const relFilename = `xl/${filename}`;
  const xmlText = await xlsx.file(relFilename)?.async("string");
  const xml = xmlText === void 0 ? createXml({ specificSchema: NS_REL, tagName: "Relationships" }) : parser.parseFromString(xmlText, "application/xml");
  if (xmlText === void 0) isDirty = true;
  if (xml.getElementsByTagName("parsererror").length > 0) {
    throw new Error(`Invalid relations file for ${filename}`);
  }
  const realationsXml = Array.from(xml.querySelectorAll("Relationship"));
  let relations = [];
  for (const el of realationsXml) {
    const id = el.getAttribute("Id");
    const target = el.getAttribute("Target");
    const elementType = el.getAttribute("Type");
    const targetMode = el.getAttribute("TargetMode");
    if (id === null || target === null || elementType === null) continue;
    const realType = elementType.slice(elementType.lastIndexOf("/") + 1);
    relations.push({
      id,
      target,
      type: elementType,
      realType,
      ...targetMode === null ? {} : { targetMode }
    });
  }
  const getNextId = () => {
    let max = 0;
    xml.querySelectorAll("Relationship[Id^='rId']").forEach((el) => {
      const n = Number((el.getAttribute("Id") ?? "Id1").slice(3));
      if (!Number.isNaN(n)) max = Math.max(max, n);
    });
    return `rId${max + 1}`;
  };
  const add = (rel) => {
    const id = getNextId();
    const root = xml.querySelector("Relationships");
    if (root === null) return "NA";
    const el = xml.createElementNS(NS_REL, "Relationship");
    el.setAttribute("Id", id);
    el.setAttribute("Type", rel.type);
    el.setAttribute("Target", rel.target);
    if (rel.targetMode !== void 0) el.setAttribute("TargetMode", rel.targetMode);
    root.appendChild(el);
    relations.push({ id, type: rel.type, target: rel.target });
    isDirty = true;
    return id;
  };
  const get = ({ by, value }) => {
    return relations.filter((relation) => relation[by] === value);
  };
  const remove = ({ by, value }) => {
    const valuesToRemove = relations.filter((relation) => relation[by] === value);
    for (const value2 of valuesToRemove) {
      xml.querySelector(`Relationship[id="${String(value2)}"]`)?.remove();
    }
    relations = relations.filter((relation) => relation[by] !== value);
  };
  const save = () => {
    if (!isDirty) return;
    const newData = serializeXml(xml);
    xlsx.file(relFilename, newData);
  };
  const getElements = (element) => {
    return relations.filter((el) => el.realType === element);
  };
  return {
    add,
    remove,
    get,
    save,
    getElements
  };
};

// src/functions/get-pivot-tables.ts
var serializer4 = new XMLSerializer();
var getPivotTables = async (xlsx) => {
  const pivotTables = /* @__PURE__ */ new Map();
  const pivotFiles = Object.keys(xlsx.files).filter((f) => f.startsWith("xl/pivotTables/") && f.endsWith(".xml")).map((f) => ({
    definitionFilename: f,
    relsFilename: f.replace("/pivotTables", "/pivotTables/_rels") + ".rels"
  }));
  for (const { definitionFilename, relsFilename } of pivotFiles) {
    const tableDefinitionText = await xlsx.file(definitionFilename)?.async("string");
    if (tableDefinitionText === void 0) continue;
    const tableDefinitionXml = parser.parseFromString(tableDefinitionText, "application/xml");
    const tableDefinition = tableDefinitionXml.querySelector("pivotTableDefinition");
    if (tableDefinition === null) continue;
    const xmlRelsText = await xlsx.file(relsFilename)?.async("string");
    if (xmlRelsText === void 0) continue;
    const rels = parser.parseFromString(xmlRelsText, "application/xml");
    const cacheTarget = rels.querySelector(`Relationship[Type="${NS_PREFIX}/relationships/pivotCacheDefinition"]`)?.getAttribute("Target");
    if (cacheTarget === void 0 || cacheTarget === null) continue;
    const cacheFile = cacheTarget.replace("../", "xl/");
    const cacheText = await xlsx.file(cacheFile)?.async("string");
    if (cacheText === void 0) continue;
    const cache = parser.parseFromString(cacheText, "application/xml");
    const workSheetSource = cache.querySelector("worksheetSource");
    if (workSheetSource === null) continue;
    const ref = workSheetSource.getAttribute("ref");
    const sheet = workSheetSource.getAttribute("sheet");
    if (ref === null || sheet === null) continue;
    const range = parseCellRange(ref);
    const extension = {
      rows: 0,
      cols: 0
    };
    let columnExtension = 0;
    let rowExtension = 0;
    let currentRow = 0;
    const extend = (cell) => {
      if (Array.isArray(cell.newValue)) {
        if (cell.row !== currentRow) {
          currentRow = cell.row;
          columnExtension = 0;
          rowExtension = 0;
        }
        const length = cell.newValue.length - 1;
        if ("$isTable" in cell.newValue) {
          rowExtension = Math.max(length, rowExtension) - Math.min(length, rowExtension);
          extension.rows = extension.rows + rowExtension;
        } else {
          columnExtension = columnExtension + length;
          extension.cols = Math.max(columnExtension, extension.cols);
        }
      }
    };
    const save2 = () => {
      tableDefinition.setAttribute("refreshOnLoad", "1");
      tableDefinition.setAttribute("saveData", "0");
      const cacheDefinition = cache.querySelector("pivotCacheDefinition");
      if (cacheDefinition !== null) {
        cacheDefinition.setAttribute("refreshOnLoad", "1");
        cacheDefinition.setAttribute("saveData", "0");
        cacheDefinition.setAttribute("invalid", "1");
        cacheDefinition.setAttribute("recordCount", "1");
      }
      const newDefinition = serializer4.serializeToString(tableDefinition);
      xlsx.file(definitionFilename, newDefinition);
      range.columnEnd = range.columnEnd + extension.cols;
      range.rowEnd = range.rowEnd + extension.rows;
      const ref2 = toCellRange(range);
      workSheetSource.setAttribute("ref", ref2);
      const newCache = serializer4.serializeToString(cache);
      xlsx.file(cacheFile, newCache);
    };
    const pivotTable = {
      ref,
      sheet,
      range,
      definitionXml: tableDefinitionXml,
      cacheXml: cache,
      extension,
      extend,
      save: save2
    };
    pivotTables.set(definitionFilename, pivotTable);
  }
  const get = (tableName) => {
    return pivotTables.get(tableName);
  };
  const save = () => {
    for (const table of pivotTables.values()) {
      table.save();
    }
  };
  const findByCell = (sheet, cell) => {
    for (const [name, value] of pivotTables) {
      if (value.sheet !== sheet) continue;
      if (isInRange(cell, value.range)) return name;
    }
    return null;
  };
  return {
    get,
    save,
    findByCell
  };
};

// src/functions/get-formulas.ts
var sheetRegex = "(?:(?<sheet>(?:'(?:[^'!]|'')+'|[^'!()+/*-]+))!)";
var completeRefRegex = "\\$?[A-Z]+\\$?[0-9]+(?::\\$?[A-Z]+\\$?[0-9]+)?";
var partialRefRegex = "\\$?[A-Z]+(?::\\$?[A-Z]+)|\\$?[0-9]+(?::\\$?[0-9]+)";
var re = new RegExp(`(?<![A-Z0-9])(?:${sheetRegex})?(?<ref>${completeRefRegex}|${partialRefRegex})(?![A-Z0-9])`, "g");
var getFormulasFromReferences = (s) => [...s.matchAll(re)].map((m) => ({
  sheet: m.groups?.sheet?.replace(/^'|'$/g, "").replace(/''/g, "'"),
  ref: m.groups?.ref
}));

// src/functions/get-data-validations.ts
var createDataValidation = (doc, dataValidation) => {
  if (dataValidation.sqref === void 0) throw new Error("The validation area is not defined");
  if (dataValidation.formula === void 0) throw new Error("The data validation formula is not defined");
  const dataValidationTag = doc.createElementNS(NS, "dataValidation");
  const formulaTag = doc.createElementNS(NS, "formula1");
  dataValidationTag.setAttribute("allowBlank", "1");
  dataValidationTag.setAttribute("showInputMessage", "1");
  dataValidationTag.setAttribute("showErrorMessage", "1");
  dataValidationTag.setAttribute("sqref", dataValidation.sqref);
  dataValidationTag.setAttribute("type", dataValidation.type);
  formulaTag.textContent = dataValidation.formula;
  dataValidationTag.appendChild(formulaTag);
  return dataValidationTag;
};
var getDataValidations = (sheets) => {
  const dataValidations = /* @__PURE__ */ new Map();
  let index = -1;
  for (const sheet of sheets) {
    const dataValidationsParent = sheet.xml.querySelector("dataValidations");
    if (dataValidationsParent === null) continue;
    const dataValidationsXml = dataValidationsParent.querySelectorAll("dataValidation");
    for (const dataValidation of dataValidationsXml) {
      const sheetName = sheet.name;
      const extension = { rows: 0, cols: 0 };
      const formulaExtension = { rows: 0, cols: 0 };
      const sqref = dataValidation.getAttribute("sqref");
      const formula = dataValidation.querySelector("formula1")?.textContent;
      if (sqref === null) continue;
      index++;
      const range = parseCellRange(sqref);
      const validation = {
        sheet: sheetName,
        sqref,
        range,
        xml: dataValidation
      };
      if (formula !== void 0) {
        validation.formula = formula;
        const references = getFormulasFromReferences(formula);
        if (references.length !== 1) continue;
        const reference = references[0];
        if (reference.ref === void 0) continue;
        validation.listRange = {
          ref: reference.ref,
          sheet: reference.sheet ?? sheetName
        };
        if (reference.ref === formula) {
          validation.listRange.range = parseCellRange(reference.ref);
        }
      }
      let columnExtension = 0;
      let rowExtension = 0;
      let currentRow = 0;
      const extend = (cell) => {
        if (Array.isArray(cell.newValue)) {
          if (cell.row !== currentRow) {
            currentRow = cell.row;
            columnExtension = 0;
            rowExtension = 0;
          }
          const length = cell.newValue.length - 1;
          if ("$isTable" in cell.newValue) {
            rowExtension = Math.max(length, rowExtension) - Math.min(length, rowExtension);
            extension.rows = extension.rows + rowExtension;
          } else {
            columnExtension = columnExtension + length;
            extension.cols = Math.max(columnExtension, extension.cols);
          }
        }
      };
      const extendFormula = (cell) => {
        if (Array.isArray(cell.newValue)) {
          const length = cell.newValue.length - 1;
          if ("$isTable" in cell.newValue) {
            formulaExtension.rows = Math.max(length, formulaExtension.rows);
          } else {
            formulaExtension.cols = Math.max(length, formulaExtension.cols);
          }
        }
      };
      validation.extension = extension;
      validation.formulaExtension = formulaExtension;
      validation.extend = extend;
      validation.extendFormula = extendFormula;
      dataValidations.set(String(index), validation);
    }
  }
  const save = () => {
    for (const dataValidation of dataValidations.values()) {
      dataValidation.range.columnEnd = dataValidation.range.columnEnd + dataValidation.extension.cols;
      dataValidation.range.rowEnd = dataValidation.range.rowEnd + dataValidation.extension.rows;
      const ref = toCellRange(dataValidation.range);
      dataValidation.xml.setAttribute("sqref", ref);
      if (dataValidation.listRange?.range !== void 0) {
        dataValidation.listRange.range.columnEnd = dataValidation.listRange.range.columnEnd + dataValidation.formulaExtension.cols;
        dataValidation.listRange.range.rowEnd = dataValidation.listRange.range.rowEnd + dataValidation.formulaExtension.rows;
        const listRef = toCellRange(dataValidation.listRange.range);
        const formula = dataValidation.xml.querySelector("formula1");
        if (formula !== null) {
          formula.textContent = listRef;
        }
      }
    }
  };
  const add = (xml, dataValidation) => {
    if (dataValidation.sqref === void 0 && dataValidation.range === void 0) {
      throw new Error("The validation area is not defined");
    }
    if (dataValidation.type === "formula" && dataValidation.formula === void 0) {
      throw new Error("Data validation formula is not defined");
    }
    if (dataValidation.type === "list" && dataValidation.formula === void 0 && dataValidation.range === void 0) {
      throw new Error("Data validation list is not defined");
    }
    if (dataValidation.range !== void 0) {
      dataValidation.sqref = toCellRange(dataValidation.range);
    }
    if (dataValidation.formula === void 0) {
      if (dataValidation.listRange?.range === void 0) {
        throw new Error("The data validation list is not defined");
      }
      dataValidation.formula = toCellRange(dataValidation.listRange.range);
    }
    const dataValidationXml = createDataValidation(sheets[0].xml, dataValidation);
    const dataValidationsTag = xml.querySelector("dataValidations");
    if (dataValidationsTag === null) {
      const elements = TAGS_ORDER.slice(TAGS_ORDER.indexOf("dataValidations") + 1);
      let element = null;
      for (const elName of elements) {
        const foundElement = xml.querySelector(elName);
        if (foundElement !== null) {
          element = foundElement;
          break;
        }
      }
      xml.querySelector("worksheet")?.insertBefore(xml.createElementNS(NS, "dataValidations"), element);
    }
    xml.querySelector("dataValidations")?.appendChild(dataValidationXml);
  };
  const get = (ref) => {
    return dataValidations.get(ref);
  };
  const findByCell = (sheetName, cell) => {
    for (const [name, value] of dataValidations) {
      if (sheetName !== value.sheet) continue;
      if (isInRange(cell, value.range)) return name;
    }
    return null;
  };
  const findListByCell = (sheetName, cell) => {
    for (const [name, value] of dataValidations) {
      if (value.listRange?.range === void 0) continue;
      if (value.listRange.sheet !== sheetName) continue;
      if (isInRange(cell, value.listRange.range)) return name;
    }
    return null;
  };
  return {
    add,
    save,
    get,
    findByCell,
    findListByCell
  };
};

// src/functions/get-extension-lists.ts
var getExtensionLists = (sheets) => {
  const extensions = /* @__PURE__ */ new Map();
  let index = -1;
  for (const sheet of sheets) {
    const extenstionListsParent = sheet.xml.querySelector("extLst dataValidations") ?? sheet.xml.createElementNS(NS, "dataValidations");
    const extenstionListsXml = extenstionListsParent.querySelectorAll("dataValidation");
    for (const extensionList of extenstionListsXml) {
      const sheetName = sheet.name;
      const extension = {
        rows: 0,
        cols: 0
      };
      const formulaExtension = {
        rows: 0,
        cols: 0
      };
      const sqref = extensionList.querySelector("sqref")?.textContent;
      const formula = extensionList.querySelector("formula1 f")?.textContent;
      if (sqref === void 0 || formula === void 0) continue;
      index++;
      const range = parseCellRange(sqref);
      const currentExtension = {
        xml: extensionList,
        sheet: sheetName,
        sqref,
        range
      };
      if (formula !== void 0) {
        currentExtension.formula = formula;
        const references = getFormulasFromReferences(formula);
        if (references.length !== 1) continue;
        const reference = references[0];
        if (reference.ref === void 0) continue;
        currentExtension.listRange = {
          ref: reference.ref,
          sheet: reference.sheet ?? sheetName
        };
        if (`${String(reference.sheet)}!${String(reference.ref)}` === formula) {
          currentExtension.listRange.range = parseCellRange(reference.ref);
        }
      }
      const extend = (cell) => {
        if (Array.isArray(cell.newValue)) {
          const length = cell.newValue.length - 1;
          if ("$isTable" in cell.newValue) {
            extension.rows = Math.max(length, extension.rows);
          } else {
            extension.cols = Math.max(length, extension.cols);
          }
        }
      };
      const extendFormula = (cell) => {
        if (Array.isArray(cell.newValue)) {
          const length = cell.newValue.length - 1;
          if ("$isTable" in cell.newValue) {
            formulaExtension.rows = Math.max(length, formulaExtension.rows);
          } else {
            formulaExtension.cols = Math.max(length, formulaExtension.cols);
          }
        }
      };
      currentExtension.extension = extension;
      currentExtension.formulaExtension = formulaExtension;
      currentExtension.extend = extend;
      currentExtension.extendFormula = extendFormula;
      extensions.set(String(index), currentExtension);
    }
  }
  const save = () => {
    for (const extension of extensions.values()) {
      extension.range.columnEnd = extension.range.columnEnd + extension.extension.cols;
      extension.range.rowEnd = extension.range.rowEnd + extension.extension.rows;
      const ref = toCellRange(extension.range);
      const sqref = extension.xml.querySelector("sqref");
      if (sqref !== null) sqref.textContent = ref;
      extension.xml.setAttribute("sqref", ref);
      if (extension.listRange?.range !== void 0) {
        extension.listRange.range.columnEnd = extension.listRange.range.columnEnd + extension.formulaExtension.cols;
        extension.listRange.range.rowEnd = extension.listRange.range.rowEnd + extension.formulaExtension.rows;
        const listRef = toCellRange(extension.listRange.range);
        const formula = extension.xml.querySelector("formula1 f");
        if (formula !== null) {
          formula.textContent = String(extension.listRange.sheet) + "!" + listRef;
        }
      }
    }
  };
  const get = (ref) => {
    return extensions.get(ref);
  };
  const findByCell = (cell) => {
    for (const [name, value] of extensions) {
      if (isInRange(cell, value.range)) return name;
    }
    return null;
  };
  const findListByCell = (sheetName, cell) => {
    for (const [name, value] of extensions) {
      if (value.listRange?.range === void 0) continue;
      if (value.listRange.sheet !== sheetName) continue;
      if (isInRange(cell, value.listRange.range)) return name;
    }
    return null;
  };
  return {
    save,
    get,
    findByCell,
    findListByCell
  };
};

// src/functions/get-persons.ts
var addPerson = (doc, person) => {
  const el = doc.createElementNS(NS_TC, "person");
  el.setAttribute("displayName", person.displayName);
  el.setAttribute("id", person.id);
  el.setAttribute("userId", person.userId);
  el.setAttribute("providerId", person.providerId);
  doc.documentElement.appendChild(el);
};
var getPersons = async (xlsx, target) => {
  let isDirty = false;
  const xmlText = await xlsx.file(`xl/${target}`)?.async("string");
  if (xmlText === void 0) isDirty = true;
  const xml = xmlText === void 0 ? createXml({ coreSchema: NS, specificSchema: NS_TC, tagName: "personList" }) : parser.parseFromString(xmlText, "application/xml");
  if (xml.getElementsByTagName("parsererror").length > 0) {
    throw new Error("Invalid person.xml");
  }
  const getSystemUserId = () => {
    const existing = Array.from(xml.getElementsByTagNameNS(NS_TC, "person")).find((p) => p.getAttribute("displayName") === "System")?.getAttribute("id");
    if (existing !== null && existing !== void 0) return existing;
    const systemPerson = {
      id: `{${generateUUID().toUpperCase()}}`,
      displayName: "System",
      userId: "S::system@local::00000000-0000-0000-0000-000000000000",
      providerId: "AD"
    };
    addPerson(xml, systemPerson);
    isDirty = true;
    return systemPerson.id;
  };
  const save = () => {
    if (!isDirty) return;
    const newData = serializeXml(xml);
    xlsx.file(`xl/${target}`, newData);
  };
  return {
    getSystemUserId,
    save
  };
};

// src/functions/get-workbook.ts
var getWorkbook = async (xlsx) => {
  const relsXmlText = await xlsx.file(GLOBAL_RELS)?.async("string");
  if (relsXmlText === void 0) throw new Error("This Excel file has no relations");
  const relsXml = parser.parseFromString(relsXmlText, "application/xml");
  const workBookElement = relsXml.querySelector(`Relationship[Type="${OFFICE_DOCUMENT_TYPE}/officeDocument"]`);
  if (workBookElement === null) throw new Error("This Excel file has no workbook");
  const workbookPath = workBookElement.getAttribute("Target");
  if (workbookPath === null) throw new Error("This Excel file has no workbook");
  const workbookXmlText = await xlsx.file(workbookPath)?.async("string");
  if (workbookXmlText === void 0) throw new Error("This Excel file has no workbook");
  const workbookXml = parser.parseFromString(workbookXmlText, "application/xml");
  const sheetsParent = workbookXml.querySelector("sheets");
  if (sheetsParent === null) throw new Error("This Excel file has no sheets");
  const sheetsXml = Array.from(sheetsParent.children);
  const sheets = sheetsXml.map((el) => ({
    id: el.getAttribute("r:id") ?? void 0,
    name: el.getAttribute("name") ?? void 0,
    sheetId: el.getAttribute("sheetId") ?? void 0
  }));
  const relationsPath = getRelationsPath(workbookPath);
  const relations = await getRelations({ xlsx, filename: relationsPath });
  const workbook = {};
  workbook.relations = relations;
  const styleProps = relations.getElements("styles")[0];
  if (styleProps === void 0) throw new Error("Styles are mailformed");
  const styles = await createStyles(xlsx, styleProps.target);
  workbook.styles = styles;
  const sharedStringsProps = relations.getElements("sharedStrings")[0];
  if (sharedStringsProps === void 0) throw new Error("This file does not contain strings");
  const sharedStrings = await getSharedStrings(xlsx, sharedStringsProps.target);
  workbook.sharedStrings = sharedStrings;
  const persons = await getPersons(xlsx, "persons/person.xml");
  workbook.persons = persons;
  const pivotTables = await getPivotTables(xlsx);
  workbook.pivotTables = pivotTables;
  const sheetsRels = relations.getElements("worksheet");
  const sheetsData = await Promise.all(sheetsRels.map(async (s) => {
    if (s?.target === void 0 || s?.id === void 0) throw new Error("Excel file is mailformed");
    const currentSheet = sheets.find((sheet2) => sheet2.id === s.id);
    if (currentSheet === void 0 || currentSheet.name === void 0) throw new Error("Excel file is mailformed");
    const relsTarget = s.target.replace(/worksheets\//, "worksheets/_rels/") + ".rels";
    const relations2 = await getRelations({ xlsx, filename: relsTarget });
    const sheetXmlText = await xlsx.file(`xl/${s.target}`)?.async("string");
    if (sheetXmlText === void 0) throw new Error(`${s.target} sheet is mailformed`);
    const xml = parser.parseFromString(sheetXmlText, "application/xml");
    const sheet = {
      xml,
      relations: relations2,
      target: s.target,
      id: s.id,
      name: currentSheet.name
    };
    Object.assign(currentSheet, sheet);
    return sheet;
  }));
  const dataValidations = getDataValidations(sheetsData);
  workbook.dataValidations = dataValidations;
  const extensionLists = getExtensionLists(sheetsData);
  workbook.extensionLists = extensionLists;
  const enrichedSheets = await Promise.all(sheetsData.map(async (sheet) => await getSheet({
    xlsx,
    workbook,
    ...sheet
  })));
  workbook.sheets = enrichedSheets;
  workbook.save = () => {
    const calcChain = relations.get({ by: "realType", value: "calcChain" });
    for (const { target } of calcChain) {
      xlsx.remove(`xl/${target}`);
    }
    relations.remove({ by: "realType", value: "calcChain" });
    sharedStrings.save();
    styles.save();
  };
  return workbook;
};

// src/functions/generate-file.ts
var generateTemplate = async (buffer, data) => {
  const xlsx = await readXlsx(buffer);
  const workbook = await getWorkbook(xlsx);
  const dataStore = createDataStore(data);
  await processWorksheets({
    workbook,
    dataStore
  });
  workbook.save();
  workbook.persons.save();
  const file = await xlsx.generateAsync({ type: "blob" });
  return file;
};

// src/index.ts
var generateXlsx = async (template, data) => {
  if (template === void 0 || data === void 0) throw Error("No template or data provided");
  if (typeof template === "string") {
    if (!/^https|ftp?:/.test(template)) throw new Error("Template URL is invalid");
    const response = await fetchFile(template);
    const fileContent = await generateTemplate(response, data);
    return fileContent;
  }
  return await generateTemplate(template, data);
};
var downloadXlsx = async (template, data, fileName, mimeType) => {
  if (fileName === void 0) fileName = (/* @__PURE__ */ new Date()).toLocaleDateString("sv") + " - Report.xlsx";
  const fileContent = await generateXlsx(template, data);
  downloadFile(fileContent, fileName, mimeType);
};
// Annotate the CommonJS export names for ESM import in node:
0 && (module.exports = {
  downloadXlsx,
  generateXlsx
});
//# sourceMappingURL=index.cjs.map