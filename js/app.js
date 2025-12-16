/**
 * XLSX Inventory Editor
 * Browser-based spreadsheet editor using ExcelJS
 */

// Global state
let workbook = null;
let fileName = '';

// Column configuration
const CONFIG = {
    skuColumn: 'B',
    quantityColumn: 'J',
    dateColumnsStart: 'L',  // Note: K contains other data, dates start at L
    dateColumnsEnd: 'T',
    headerCell: 'G1',
    headerSheet: 'WAREHOUSE',
    searchAllSheets: true,  // Toggle for multi-sheet search
    excludeSheets: []       // Sheet names to exclude from search (if any)
};

// DOM Elements
const fileInput = document.getElementById('file-input');
const downloadBtn = document.getElementById('download-btn');
const searchSection = document.getElementById('search-section');
const skuInput = document.getElementById('sku-input');
const searchBtn = document.getElementById('search-btn');
const resultsDiv = document.getElementById('results');
const headerSection = document.getElementById('header-section');
const nameInput = document.getElementById('name-input');
const updateHeaderBtn = document.getElementById('update-header-btn');

// Initialize event listeners
document.addEventListener('DOMContentLoaded', init);

function init() {
    fileInput.addEventListener('change', handleFileUpload);
    downloadBtn.addEventListener('click', handleDownload);
    searchBtn.addEventListener('click', handleSearch);
    updateHeaderBtn.addEventListener('click', handleUpdateHeader);

    // Allow Enter key to trigger search
    skuInput.addEventListener('keypress', (e) => {
        if (e.key === 'Enter') handleSearch();
    });

    // Handle multi-sheet toggle
    const searchAllSheetsCheckbox = document.getElementById('search-all-sheets');
    if (searchAllSheetsCheckbox) {
        searchAllSheetsCheckbox.addEventListener('change', (e) => {
            CONFIG.searchAllSheets = e.target.checked;
        });
    }
}

/**
 * Handle file upload and parse XLSX
 */
async function handleFileUpload(event) {
    const file = event.target.files[0];
    if (!file) return;

    fileName = file.name;

    try {
        const arrayBuffer = await file.arrayBuffer();
        workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(arrayBuffer);

        // Enable UI elements
        downloadBtn.disabled = false;
        searchSection.style.display = 'block';
        headerSection.style.display = 'block';

        showStatus(`Loaded: ${fileName} (${workbook.worksheets.length} sheets)`, 'success');
    } catch (error) {
        showStatus(`Error loading file: ${error.message}`, 'error');
        console.error(error);
    }
}

/**
 * Handle file download
 */
async function handleDownload() {
    if (!workbook) return;

    try {
        const buffer = await workbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], {
            type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        });

        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = fileName.replace('.xlsx', '_modified.xlsx');
        a.click();
        URL.revokeObjectURL(url);

        showStatus('File downloaded successfully', 'success');
    } catch (error) {
        showStatus(`Error downloading: ${error.message}`, 'error');
    }
}

/**
 * Show status message to user
 */
function showStatus(message, type) {
    const statusDiv = document.createElement('div');
    statusDiv.className = `status-message ${type}`;
    statusDiv.textContent = message;

    resultsDiv.insertBefore(statusDiv, resultsDiv.firstChild);

    // Remove after 3 seconds
    setTimeout(() => statusDiv.remove(), 3000);
}

/**
 * Search for SKU on sheets (single or all based on config)
 */
function handleSearch() {
    const sku = skuInput.value.trim();
    if (!sku) {
        showStatus('Please enter a SKU to search', 'error');
        return;
    }

    if (!workbook) {
        showStatus('Please load a file first', 'error');
        return;
    }

    // Use config to determine search scope
    const results = CONFIG.searchAllSheets
        ? searchAllSheets(sku)
        : searchWarehouse(sku);

    displayResults(results, sku);
}

/**
 * Search WAREHOUSE sheet for SKU matches
 * @param {string} sku - SKU to search for
 * @returns {Array} Array of match objects
 */
function searchWarehouse(sku) {
    const matches = [];
    const skuColIndex = columnLetterToIndex(CONFIG.skuColumn);
    const qtyColIndex = columnLetterToIndex(CONFIG.quantityColumn);

    const sheet = workbook.getWorksheet(CONFIG.headerSheet);
    if (!sheet) {
        showStatus(`Sheet "${CONFIG.headerSheet}" not found`, 'error');
        return matches;
    }

    sheet.eachRow((row, rowNumber) => {
        // Skip header row (row 1)
        if (rowNumber === 1) return;

        const cellValue = getCellValue(row.getCell(skuColIndex));

        if (cellValue && cellValue.toString().toLowerCase() === sku.toLowerCase()) {
            matches.push({
                sheetName: sheet.name,
                rowNumber: rowNumber,
                sku: cellValue,
                quantity: getCellValue(row.getCell(qtyColIndex)),
                row: row
            });
        }
    });

    return matches;
}

/**
 * Search all sheets for SKU matches
 * @param {string} sku - SKU to search for
 * @returns {Array} Array of match objects
 */
function searchAllSheets(sku) {
    const matches = [];
    const skuColIndex = columnLetterToIndex(CONFIG.skuColumn);
    const qtyColIndex = columnLetterToIndex(CONFIG.quantityColumn);

    workbook.worksheets.forEach(sheet => {
        // Skip excluded sheets
        if (CONFIG.excludeSheets.includes(sheet.name)) return;

        sheet.eachRow((row, rowNumber) => {
            // Skip header row (row 1)
            if (rowNumber === 1) return;

            const cellValue = getCellValue(row.getCell(skuColIndex));

            if (cellValue && cellValue.toString().toLowerCase() === sku.toLowerCase()) {
                matches.push({
                    sheetName: sheet.name,
                    rowNumber: rowNumber,
                    sku: cellValue,
                    quantity: getCellValue(row.getCell(qtyColIndex)),
                    row: row
                });
            }
        });
    });

    return matches;
}

/**
 * Display search results
 * @param {Array} results - Array of match objects
 * @param {string} searchTerm - Original search term
 */
function displayResults(results, searchTerm) {
    resultsDiv.innerHTML = '';

    if (results.length === 0) {
        resultsDiv.innerHTML = `<p>No results found for SKU: <strong>${escapeHtml(searchTerm)}</strong></p>`;
        return;
    }

    const heading = document.createElement('h3');
    heading.textContent = `Found ${results.length} match${results.length > 1 ? 'es' : ''} for "${searchTerm}"`;
    resultsDiv.appendChild(heading);

    results.forEach((result, index) => {
        const item = createResultItem(result, index);
        resultsDiv.appendChild(item);
    });
}

/**
 * Create a result item DOM element
 * @param {Object} result - Match object
 * @param {number} index - Index for unique identification
 * @returns {HTMLElement}
 */
function createResultItem(result, index) {
    const div = document.createElement('div');
    div.className = 'result-item';
    div.dataset.sheetName = result.sheetName;
    div.dataset.rowNumber = result.rowNumber;
    div.id = `result-${index}`;

    // Use multi-sheet highlight check when multi-sheet search is enabled
    const isHighlighted = CONFIG.searchAllSheets
        ? checkRowHighlightMulti(result.sheetName, result.rowNumber)
        : checkRowHighlight(result.rowNumber);
    if (isHighlighted) {
        div.classList.add('highlighted');
    }

    // Show sheet name if multi-sheet search is enabled
    const headerText = CONFIG.searchAllSheets
        ? `Sheet: ${escapeHtml(result.sheetName)} | Row: ${result.rowNumber}`
        : `Row: ${result.rowNumber}`;

    // Escape sheet name for onclick - replace quotes with HTML entities
    const safeSheetName = result.sheetName.replace(/&/g, '&amp;').replace(/'/g, '&#39;');

    div.innerHTML = `
        <h4>${headerText}</h4>
        <label>
            SKU: <strong>${escapeHtml(result.sku)}</strong>
        </label>
        <label>
            Quantity:
            <input type="number"
                   class="qty-input"
                   value="${result.quantity || 0}"
                   data-index="${index}"
                   min="0">
        </label>
        <div class="result-actions">
            <button class="save-btn" onclick="saveQuantityMulti(${index}, '${safeSheetName}', ${result.rowNumber})">
                Save Quantity
            </button>
            <button class="highlight-btn" onclick="toggleHighlightMulti(${index}, '${safeSheetName}', ${result.rowNumber})">
                ${isHighlighted ? 'Remove Highlight' : 'Highlight Low Stock'}
            </button>
        </div>
    `;

    return div;
}

/**
 * Check if a row has yellow background highlight (WAREHOUSE only)
 */
function checkRowHighlight(rowNumber) {
    const sheet = workbook.getWorksheet(CONFIG.headerSheet);
    if (!sheet) return false;

    const row = sheet.getRow(rowNumber);
    const cell = row.getCell(columnLetterToIndex(CONFIG.skuColumn));

    if (cell.fill && cell.fill.fgColor) {
        const color = cell.fill.fgColor.argb || '';
        // Check for yellow (FFFFFF00 or similar)
        return color.toUpperCase().includes('FFFF00');
    }
    return false;
}

/**
 * Check highlight on any sheet (multi-sheet version)
 * @param {string} sheetName - Name of the sheet
 * @param {number} rowNumber - Row number
 * @returns {boolean}
 */
function checkRowHighlightMulti(sheetName, rowNumber) {
    const sheet = workbook.getWorksheet(sheetName);
    if (!sheet) return false;

    const row = sheet.getRow(rowNumber);
    const cell = row.getCell(columnLetterToIndex(CONFIG.skuColumn));

    if (cell.fill && cell.fill.fgColor) {
        const color = cell.fill.fgColor.argb || '';
        return color.toUpperCase().includes('FFFF00');
    }
    return false;
}

// Utility functions

/**
 * Convert column letter to 1-based index
 * @param {string} letter - Column letter (e.g., 'A', 'B', 'AA')
 * @returns {number}
 */
function columnLetterToIndex(letter) {
    let index = 0;
    for (let i = 0; i < letter.length; i++) {
        index = index * 26 + (letter.charCodeAt(i) - 64);
    }
    return index;
}

/**
 * Get cell value handling different types
 * @param {Object} cell - ExcelJS cell object
 * @returns {any}
 */
function getCellValue(cell) {
    if (!cell || cell.value === null || cell.value === undefined) {
        return '';
    }

    // Handle rich text
    if (cell.value.richText) {
        return cell.value.richText.map(r => r.text).join('');
    }

    // Handle formula results
    if (cell.value.result !== undefined) {
        return cell.value.result;
    }

    return cell.value;
}

/**
 * Escape HTML to prevent XSS
 * @param {string} str - String to escape
 * @returns {string}
 */
function escapeHtml(str) {
    if (typeof str !== 'string') return str;
    return str
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#039;');
}

/**
 * Save quantity change and update date/clear cells
 * @param {number} index - Result index
 * @param {number} rowNumber - Row number
 */
function saveQuantity(index, rowNumber) {
    const resultItem = document.getElementById(`result-${index}`);
    const qtyInput = resultItem.querySelector('.qty-input');
    const newQuantity = parseInt(qtyInput.value, 10);

    if (isNaN(newQuantity) || newQuantity < 0) {
        showStatus('Please enter a valid quantity', 'error');
        return;
    }

    const sheet = workbook.getWorksheet(CONFIG.headerSheet);
    if (!sheet) {
        showStatus(`Sheet "${CONFIG.headerSheet}" not found`, 'error');
        return;
    }

    const row = sheet.getRow(rowNumber);
    const qtyColIndex = columnLetterToIndex(CONFIG.quantityColumn);

    // Get old quantity to calculate change
    const oldQuantity = parseInt(getCellValue(row.getCell(qtyColIndex)), 10) || 0;
    const qtyChange = newQuantity - oldQuantity;

    // Update quantity
    row.getCell(qtyColIndex).value = newQuantity;

    // Find next empty date column and add date with quantity change
    const dateColIndex = findNextEmptyDateColumn(row);
    if (dateColIndex) {
        const dateStr = formatDateWithQty(new Date(), qtyChange);
        row.getCell(dateColIndex).value = dateStr;

        // Clear the column after the date
        const clearColIndex = dateColIndex + 1;
        if (clearColIndex <= columnLetterToIndex(CONFIG.dateColumnsEnd) + 1) {
            row.getCell(clearColIndex).value = null;
        }

        showStatus(`Saved: Qty ${oldQuantity}→${newQuantity} (${qtyChange >= 0 ? '+' : ''}${qtyChange}), Date in column ${indexToColumnLetter(dateColIndex)}`, 'success');
    } else {
        showStatus(`Saved: Qty=${newQuantity} (No empty date column in L-T range)`, 'success');
    }
}

/**
 * Save quantity on any sheet (multi-sheet version)
 * @param {number} index - Result index
 * @param {string} sheetName - Name of the sheet
 * @param {number} rowNumber - Row number
 */
function saveQuantityMulti(index, sheetName, rowNumber) {
    console.log('saveQuantityMulti called:', { index, sheetName, rowNumber });

    const resultItem = document.getElementById(`result-${index}`);
    if (!resultItem) {
        console.error(`Result item not found: result-${index}`);
        showStatus('Error: Result item not found', 'error');
        return;
    }

    const qtyInput = resultItem.querySelector('.qty-input');
    if (!qtyInput) {
        console.error('Quantity input not found');
        showStatus('Error: Quantity input not found', 'error');
        return;
    }

    const newQuantity = parseInt(qtyInput.value, 10);

    if (isNaN(newQuantity) || newQuantity < 0) {
        showStatus('Please enter a valid quantity', 'error');
        return;
    }

    const sheet = workbook.getWorksheet(sheetName);
    if (!sheet) {
        showStatus(`Sheet "${sheetName}" not found`, 'error');
        return;
    }

    const row = sheet.getRow(rowNumber);
    const qtyColIndex = columnLetterToIndex(CONFIG.quantityColumn);

    // Get old quantity to calculate change
    const oldQuantity = parseInt(getCellValue(row.getCell(qtyColIndex)), 10) || 0;
    const qtyChange = newQuantity - oldQuantity;

    // Update quantity
    row.getCell(qtyColIndex).value = newQuantity;

    // Find next empty date column and add date with quantity change
    const dateColIndex = findNextEmptyDateColumn(row);
    if (dateColIndex) {
        const dateStr = formatDateWithQty(new Date(), qtyChange);
        row.getCell(dateColIndex).value = dateStr;

        // Clear the column after the date
        const clearColIndex = dateColIndex + 1;
        if (clearColIndex <= columnLetterToIndex(CONFIG.dateColumnsEnd) + 1) {
            row.getCell(clearColIndex).value = null;
        }

        showStatus(`Saved on ${sheetName}: Qty ${oldQuantity}→${newQuantity} (${qtyChange >= 0 ? '+' : ''}${qtyChange}), Date in column ${indexToColumnLetter(dateColIndex)}`, 'success');
    } else {
        showStatus(`Saved on ${sheetName}: Qty=${newQuantity} (No empty date column)`, 'success');
    }

}

/**
 * Find next empty column in date range (L-T)
 * @param {Object} row - ExcelJS row object
 * @returns {number|null} Column index or null if all full
 */
function findNextEmptyDateColumn(row) {
    const startIndex = columnLetterToIndex(CONFIG.dateColumnsStart);
    const endIndex = columnLetterToIndex(CONFIG.dateColumnsEnd);

    for (let i = startIndex; i <= endIndex; i++) {
        const cell = row.getCell(i);
        const value = getCellValue(cell);
        if (value === '' || value === null || value === undefined) {
            return i;
        }
    }

    return null; // All columns full
}

/**
 * Format date as D/M/YY-N (e.g., "16/12/25-3")
 * @param {Date} date - Date to format
 * @param {number} qtyChange - Quantity change (absolute value)
 * @returns {string}
 */
function formatDateWithQty(date, qtyChange) {
    const day = date.getDate();  // No padding - matches existing format
    const month = date.getMonth() + 1;  // No padding
    const year = String(date.getFullYear()).slice(-2);
    return `${day}/${month}/${year}-${Math.abs(qtyChange)}`;
}

/**
 * Format date as D/M/YY (for header)
 * @param {Date} date - Date to format
 * @returns {string}
 */
function formatDate(date) {
    const day = date.getDate();
    const month = date.getMonth() + 1;
    const year = String(date.getFullYear()).slice(-2);
    return `${day}/${month}/${year}`;
}

/**
 * Convert column index to letter
 * @param {number} index - 1-based column index
 * @returns {string}
 */
function indexToColumnLetter(index) {
    let letter = '';
    while (index > 0) {
        const remainder = (index - 1) % 26;
        letter = String.fromCharCode(65 + remainder) + letter;
        index = Math.floor((index - 1) / 26);
    }
    return letter;
}

/**
 * Toggle yellow highlight on a row
 * @param {number} index - Result index
 * @param {number} rowNumber - Row number
 */
function toggleHighlight(index, rowNumber) {
    const sheet = workbook.getWorksheet(CONFIG.headerSheet);
    if (!sheet) {
        showStatus(`Sheet "${CONFIG.headerSheet}" not found`, 'error');
        return;
    }

    const resultItem = document.getElementById(`result-${index}`);
    const isCurrentlyHighlighted = resultItem.classList.contains('highlighted');

    const startCol = columnLetterToIndex('A');
    const endCol = columnLetterToIndex(CONFIG.dateColumnsEnd) + 1;

    // Apply fill to each cell individually
    for (let col = startCol; col <= endCol; col++) {
        const cell = sheet.getCell(rowNumber, col);
        if (isCurrentlyHighlighted) {
            cell.fill = { type: 'pattern', pattern: 'none' };
        } else {
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFFFFF00' }
            };
        }
    }

    // Update UI
    if (isCurrentlyHighlighted) {
        resultItem.classList.remove('highlighted');
        resultItem.querySelector('.highlight-btn').textContent = 'Highlight Low Stock';
        showStatus('Highlight removed', 'success');
    } else {
        resultItem.classList.add('highlighted');
        resultItem.querySelector('.highlight-btn').textContent = 'Remove Highlight';
        showStatus('Row highlighted yellow for low stock', 'success');
    }
}

/**
 * Toggle highlight on any sheet (multi-sheet version)
 * @param {number} index - Result index
 * @param {string} sheetName - Name of the sheet
 * @param {number} rowNumber - Row number
 */
function toggleHighlightMulti(index, sheetName, rowNumber) {
    console.log('toggleHighlightMulti called:', { index, sheetName, rowNumber });

    const sheet = workbook.getWorksheet(sheetName);
    if (!sheet) {
        showStatus(`Sheet "${sheetName}" not found`, 'error');
        return;
    }

    const resultItem = document.getElementById(`result-${index}`);
    if (!resultItem) {
        console.error(`Result item not found: result-${index}`);
        showStatus('Error: Result item not found', 'error');
        return;
    }
    const isCurrentlyHighlighted = resultItem.classList.contains('highlighted');

    const startCol = columnLetterToIndex('A');
    const endCol = columnLetterToIndex(CONFIG.dateColumnsEnd) + 1;

    // Apply fill to each cell individually to avoid shared reference issues
    for (let col = startCol; col <= endCol; col++) {
        const cell = sheet.getCell(rowNumber, col);
        if (isCurrentlyHighlighted) {
            cell.fill = { type: 'pattern', pattern: 'none' };
        } else {
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFFFFF00' }
            };
        }
    }

    if (isCurrentlyHighlighted) {
        resultItem.classList.remove('highlighted');
        resultItem.querySelector('.highlight-btn').textContent = 'Highlight Low Stock';
        showStatus(`Highlight removed from row ${rowNumber} on ${sheetName}`, 'success');
    } else {
        resultItem.classList.add('highlighted');
        resultItem.querySelector('.highlight-btn').textContent = 'Remove Highlight';
        showStatus(`Row ${rowNumber} highlighted on ${sheetName}`, 'success');
    }
}

/**
 * Update header cell (G1 on WAREHOUSE sheet) with name and date
 */
function handleUpdateHeader() {
    const name = nameInput.value.trim();
    if (!name) {
        showStatus('Please enter your name', 'error');
        return;
    }

    if (!workbook) {
        showStatus('Please load a file first', 'error');
        return;
    }

    const sheet = workbook.getWorksheet(CONFIG.headerSheet);
    if (!sheet) {
        showStatus(`Sheet "${CONFIG.headerSheet}" not found`, 'error');
        return;
    }

    const dateStr = formatDate(new Date());
    const headerText = `Date Changed - ${dateStr} ${name}`;

    const cell = sheet.getCell(CONFIG.headerCell);
    cell.value = headerText;

    showStatus(`Header updated: "${headerText}"`, 'success');
}

// Export functions to global scope (required for inline onclick handlers)
window.columnLetterToIndex = columnLetterToIndex;
window.indexToColumnLetter = indexToColumnLetter;
window.formatDate = formatDate;
window.formatDateWithQty = formatDateWithQty;
window.escapeHtml = escapeHtml;
window.getCellValue = getCellValue;
window.searchWarehouse = searchWarehouse;
window.findNextEmptyDateColumn = findNextEmptyDateColumn;
window.checkRowHighlight = checkRowHighlight;
window.saveQuantity = saveQuantity;
window.toggleHighlight = toggleHighlight;
// Multi-sheet functions
window.searchAllSheets = searchAllSheets;
window.checkRowHighlightMulti = checkRowHighlightMulti;
window.saveQuantityMulti = saveQuantityMulti;
window.toggleHighlightMulti = toggleHighlightMulti;
window.CONFIG = CONFIG;
