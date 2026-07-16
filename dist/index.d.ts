declare const generateXlsx: (template: string | ArrayBuffer, data: Record<string, unknown>) => Promise<Blob>;
/**
 * Downloads an XLSX file by replacing data in a template file and initiates the download in the browser.
 * @param {string} templateURL - The URL of the template XLSX file.
 * @param {Object} data - The data object containing values to replace in the template.
 * @param {string} [fileName] - The name of the file to be downloaded. If not provided, a default name based on the current date will be used.
 * @returns {Promise<void>} A promise that resolves once the download is initiated.
 */
declare const downloadXlsx: (template: string | ArrayBuffer, data: Record<string, unknown>, fileName?: string, mimeType?: string) => Promise<void>;

export { downloadXlsx, generateXlsx };
