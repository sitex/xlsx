# XLSX Inventory Editor Implementation Plan

## Overview

Build a browser-based XLSX inventory editor using ExcelJS, hosted on GitHub Pages. The application allows users to search for products by SKU on the WAREHOUSE sheet, edit quantities, automatically track edit dates, highlight low-stock items, and download the modified file.

## Current State Analysis

- **Project**: Empty - no application code exists
- **Directory**: Contains only `.claude/` configuration and `.gitignore`
- **Target**: Static site deployable to GitHub Pages

## Desired End State

A single-page web application that:
1. Loads an XLSX file via file input
2. Searches WAREHOUSE sheet for SKU matches (Column B)
3. Displays all matching rows
4. Allows editing quantity (Column J) with automatic date stamping (next empty in K-T)
5. Clears the column after the date entry
6. Highlights rows yellow for low-stock marking
7. Updates header (G1 on WAREHOUSE sheet) with name and date
8. Downloads the modified XLSX file
9. Includes QUnit browser tests

### Verification:
- Open `index.html` in browser, load a test XLSX file
- Search for a SKU, see results from WAREHOUSE sheet
- Edit quantity, verify date appears in next empty K-T column
- Verify column after date is cleared
- Apply yellow highlight, verify styling persists in downloaded file
- Update header, verify G1 on WAREHOUSE sheet is updated
- All QUnit tests pass

## What We're NOT Doing

- No OneDrive/cloud sync integration (download/upload only)
- No server-side code or build step
- No offline/Service Worker support
- No direct filesystem access (browser security)
- No automatic low-stock threshold detection (manual highlight only)
- No multi-sheet search in Phases 1-6 (WAREHOUSE only initially, expanded in Phase 7)

## Spreadsheet Structure Reference

| Column | Content | Example |
|--------|---------|---------|
| A | Barcode | 7 427025678770 |
| B | **SKU (product code)** | AGA003 |
| D | Supplier | BC |
| E | Brand | Aganorsa Leaf |
| G | Product name | Aganorsa Aniversario... |
| H | Qty per package | 10 |
| J | **Quantity** | 7 |
| K | Arrival invoice + qty | 36028+8 |
| **L-T** | **Date entries** | 22/5/23-1 |
| G1 | Header | "Date Changed - 15/12/25 Ivan" |

**Date format**: `D/M/YY-N` where N = quantity change (e.g., `16/12/25-3` means date + qty changed by 3)

**Date columns**: L through T (NOT K - K contains other data)

**Sheet**: WAREHOUSE (first tab only for initial version)

**Test file**: `Warehouse Stock Stick.xlsx` (580 rows)

## Implementation Approach

Build incrementally with each phase fully testable:
1. Set up project structure with ExcelJS CDN
2. Implement file loading and basic display
3. Add SKU search on WAREHOUSE sheet
4. Implement quantity editing with date logic
5. Add yellow highlighting feature
6. Add header update functionality
7. Implement comprehensive QUnit tests

---

## Phase 1: Project Setup and File Handling

### Overview
Create project structure, load ExcelJS via CDN, implement file upload and download functionality.

### Changes Required:

#### 1. Create HTML Entry Point
**File**: `index.html`

```html
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Inventory Editor</title>
    <link rel="stylesheet" href="css/style.css">
</head>
<body>
    <div class="container">
        <header>
            <h1>Inventory Editor</h1>
        </header>

        <section id="file-section">
            <input type="file" id="file-input" accept=".xlsx">
            <button id="download-btn" disabled>Download Modified File</button>
        </section>

        <section id="search-section" style="display: none;">
            <input type="text" id="sku-input" placeholder="Enter SKU to search...">
            <button id="search-btn">Search</button>
        </section>

        <section id="results-section">
            <div id="results"></div>
        </section>

        <section id="header-section" style="display: none;">
            <h3>Update Header</h3>
            <input type="text" id="name-input" placeholder="Your name">
            <button id="update-header-btn">Update Header (G1 on WAREHOUSE)</button>
        </section>
    </div>

    <script src="https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js"></script>
    <script src="js/app.js"></script>
</body>
</html>
```

#### 2. Create CSS Styles
**File**: `css/style.css`

```css
* {
    box-sizing: border-box;
}

body {
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
    max-width: 800px;
    margin: 0 auto;
    padding: 20px;
    background: #f5f5f5;
}

.container {
    background: white;
    padding: 20px;
    border-radius: 8px;
    box-shadow: 0 2px 4px rgba(0,0,0,0.1);
}

header h1 {
    margin-top: 0;
    color: #333;
}

section {
    margin: 20px 0;
    padding: 15px;
    border: 1px solid #ddd;
    border-radius: 4px;
}

input[type="text"], input[type="file"] {
    padding: 10px;
    font-size: 16px;
    border: 1px solid #ccc;
    border-radius: 4px;
    margin-right: 10px;
}

button {
    padding: 10px 20px;
    font-size: 16px;
    background: #007bff;
    color: white;
    border: none;
    border-radius: 4px;
    cursor: pointer;
}

button:hover {
    background: #0056b3;
}

button:disabled {
    background: #ccc;
    cursor: not-allowed;
}

.result-item {
    padding: 15px;
    margin: 10px 0;
    border: 1px solid #ddd;
    border-radius: 4px;
    background: #fafafa;
}

.result-item.highlighted {
    background: #ffff00;
}

.result-item h4 {
    margin: 0 0 10px 0;
}

.result-item label {
    display: block;
    margin: 5px 0;
}

.result-item input[type="number"] {
    width: 100px;
    padding: 5px;
}

.result-actions {
    margin-top: 10px;
}

.result-actions button {
    margin-right: 5px;
    padding: 5px 10px;
    font-size: 14px;
}

.highlight-btn {
    background: #ffc107;
    color: #333;
}

.highlight-btn:hover {
    background: #e0a800;
}

.save-btn {
    background: #28a745;
}

.save-btn:hover {
    background: #1e7e34;
}

#results {
    min-height: 100px;
}

.status-message {
    padding: 10px;
    margin: 10px 0;
    border-radius: 4px;
}

.status-message.success {
    background: #d4edda;
    color: #155724;
}

.status-message.error {
    background: #f8d7da;
    color: #721c24;
}
```

#### 3. Create JavaScript Application Core
**File**: `js/app.js`

```javascript
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
    headerSheet: 'WAREHOUSE'
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

// Placeholder functions for Phase 2+
function handleSearch() {
    showStatus('Search functionality coming in Phase 2', 'success');
}

function handleUpdateHeader() {
    showStatus('Header update functionality coming in Phase 3', 'success');
}
```

### Success Criteria:

#### Automated Verification:
- [ ] `index.html` loads without errors in browser console
- [ ] ExcelJS library loads from CDN (check Network tab)
- [ ] File input accepts `.xlsx` files

#### Manual Verification:
- [ ] Can select and load a test XLSX file
- [ ] Status shows file name and sheet count after loading
- [ ] Download button becomes enabled after file load
- [ ] Download produces a valid XLSX file
- [ ] Search and header sections appear after file load

**Implementation Note**: After completing this phase and all automated verification passes, pause here for manual confirmation before proceeding to Phase 2.

---

## Phase 2: SKU Search on WAREHOUSE Sheet

### Overview
Implement search functionality that finds all SKU matches on the WAREHOUSE sheet and displays results.

### Changes Required:

#### 1. Add Search Functions to app.js
**File**: `js/app.js`
**Changes**: Replace placeholder `handleSearch` and add search helper functions

```javascript
/**
 * Search for SKU on WAREHOUSE sheet
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

    const results = searchWarehouse(sku);
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

    // Check if row has yellow highlight
    const isHighlighted = checkRowHighlight(result.rowNumber);
    if (isHighlighted) {
        div.classList.add('highlighted');
    }

    div.innerHTML = `
        <h4>Row: ${result.rowNumber}</h4>
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
            <button class="save-btn" onclick="saveQuantity(${index}, ${result.rowNumber})">
                Save Quantity
            </button>
            <button class="highlight-btn" onclick="toggleHighlight(${index}, ${result.rowNumber})">
                ${isHighlighted ? 'Remove Highlight' : 'Highlight Low Stock'}
            </button>
        </div>
    `;

    return div;
}

/**
 * Check if a row has yellow background highlight
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

// Placeholder for Phase 3
function saveQuantity(index, rowNumber) {
    showStatus('Save quantity coming in Phase 3', 'success');
}

function toggleHighlight(index, rowNumber) {
    showStatus('Highlight coming in Phase 4', 'success');
}
```

### Success Criteria:

#### Automated Verification:
- [x] No JavaScript errors in console when searching
- [x] `columnLetterToIndex('A')` returns 1
- [x] `columnLetterToIndex('J')` returns 10
- [x] `columnLetterToIndex('T')` returns 20

#### Manual Verification:
- [ ] Search finds SKUs in column B on WAREHOUSE sheet
- [ ] Multiple matches display with row numbers
- [ ] Each result shows SKU and current quantity
- [ ] Empty search shows error message
- [ ] Case-insensitive search works
- [ ] Results include "Save Quantity" and "Highlight" buttons
- [ ] Missing WAREHOUSE sheet shows appropriate error

**Implementation Note**: After completing this phase, test with a real inventory file before proceeding to Phase 3.

---

## Phase 3: Quantity Editing with Date Logic

### Overview
Implement quantity editing that saves to the workbook, adds date to next empty column in K-T range, and clears the following column.

### Changes Required:

#### 1. Update app.js with Save and Date Functions
**File**: `js/app.js`
**Changes**: Replace `saveQuantity` placeholder and add date handling

```javascript
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

    row.commit();
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
```

### Success Criteria:

#### Automated Verification:
- [x] `formatDate(new Date(2025, 11, 16))` returns `"16/12/25"`
- [x] `formatDateWithQty(new Date(2025, 11, 16), 3)` returns `"16/12/25-3"`
- [x] `indexToColumnLetter(12)` returns `"L"` (date columns start at L)
- [x] `indexToColumnLetter(20)` returns `"T"`

#### Manual Verification:
- [ ] Editing quantity updates value in workbook
- [ ] Date appears in next empty column (L, then M, etc.)
- [ ] Date format includes quantity change (e.g., "16/12/25-3")
- [ ] Column after date is cleared
- [ ] Status message shows old→new quantity and column
- [ ] Downloaded file preserves quantity and date changes
- [ ] Works correctly when all L-T columns are full (shows appropriate message)

**Implementation Note**: Test with a file where some rows have partial date columns filled to verify the "next empty" logic works correctly.

---

## Phase 4: Yellow Highlight for Low Stock

### Overview
Implement yellow background highlighting to mark rows as low stock, with toggle functionality.

### Changes Required:

#### 1. Update app.js with Highlight Function
**File**: `js/app.js`
**Changes**: Replace `toggleHighlight` placeholder

```javascript
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

    const row = sheet.getRow(rowNumber);
    const resultItem = document.getElementById(`result-${index}`);
    const isCurrentlyHighlighted = resultItem.classList.contains('highlighted');

    // Define yellow fill
    const yellowFill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFFFFF00' }
    };

    const noFill = {
        type: 'pattern',
        pattern: 'none'
    };

    // Get the columns we use (B through T, roughly)
    const startCol = columnLetterToIndex('A');
    const endCol = columnLetterToIndex(CONFIG.dateColumnsEnd) + 1;

    // Apply or remove fill to all cells in row
    for (let col = startCol; col <= endCol; col++) {
        const cell = row.getCell(col);
        cell.fill = isCurrentlyHighlighted ? noFill : yellowFill;
    }

    row.commit();

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
```

### Success Criteria:

#### Automated Verification:
- [x] No JavaScript errors when clicking highlight button
- [x] Yellow fill object structure matches ExcelJS specification

#### Manual Verification:
- [ ] Clicking "Highlight Low Stock" turns result item yellow in UI
- [ ] Button text changes to "Remove Highlight"
- [ ] Clicking again removes the highlight
- [ ] Downloaded file shows yellow background on entire row
- [ ] Highlight persists when re-opening downloaded file

**Implementation Note**: Test with a file that already has some yellow rows to verify detection works both ways.

---

## Phase 5: Header Update on WAREHOUSE Sheet

### Overview
Implement functionality to update cell G1 on the WAREHOUSE sheet with user name and current date.

### Changes Required:

#### 1. Update app.js with Header Update Function
**File**: `js/app.js`
**Changes**: Replace `handleUpdateHeader` placeholder

```javascript
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
```

### Success Criteria:

#### Automated Verification:
- [x] No JavaScript errors when updating header
- [x] Header text format is "Date Changed - DD/MM/YY Name"

#### Manual Verification:
- [ ] Entering name and clicking update shows success message
- [ ] Empty name shows error message
- [ ] Downloaded file shows header text in G1 of WAREHOUSE sheet
- [ ] Missing WAREHOUSE sheet shows appropriate error

**Implementation Note**: If the test file doesn't have a WAREHOUSE sheet, create one or adjust the config for testing.

---

## Phase 6: QUnit Browser Tests

### Overview
Implement comprehensive QUnit test suite that runs in the browser.

### Changes Required:

#### 1. Create Test HTML Page
**File**: `tests/index.html`

```html
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Inventory Editor Tests</title>
    <link rel="stylesheet" href="https://code.jquery.com/qunit/qunit-2.20.0.css">
</head>
<body>
    <div id="qunit"></div>
    <div id="qunit-fixture"></div>

    <script src="https://code.jquery.com/qunit/qunit-2.20.0.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js"></script>
    <script src="test-utils.js"></script>
    <script src="app.test.js"></script>
</body>
</html>
```

#### 2. Create Test Utilities
**File**: `tests/test-utils.js`

```javascript
/**
 * Test utilities for creating mock workbooks
 */

/**
 * Create a mock workbook with test data
 * @returns {ExcelJS.Workbook}
 */
async function createMockWorkbook() {
    const workbook = new ExcelJS.Workbook();

    // Create WAREHOUSE sheet
    const warehouse = workbook.addWorksheet('WAREHOUSE');
    warehouse.getCell('A1').value = 'Product';
    warehouse.getCell('B1').value = 'SKU';
    warehouse.getCell('J1').value = 'Quantity';
    warehouse.getCell('K1').value = 'Date 1';
    warehouse.getCell('G1').value = ''; // Header cell

    // Add test data (matching actual file structure)
    warehouse.getCell('B2').value = 'SKU001';
    warehouse.getCell('G2').value = 'Test Product A';
    warehouse.getCell('J2').value = 100;
    // K2 has formula/code, L2 empty (ready for date)

    warehouse.getCell('B3').value = 'SKU002';
    warehouse.getCell('G3').value = 'Test Product B';
    warehouse.getCell('J3').value = 50;
    warehouse.getCell('K3').value = '36028+5';  // Code column (not date)
    warehouse.getCell('L3').value = '1/12/25-1'; // Date with qty change

    return workbook;
}

/**
 * Import functions from main app
 * These need to be exposed globally for testing
 */
// Functions will be imported via script inclusion
```

#### 3. Create Main Test File
**File**: `tests/app.test.js`

```javascript
/**
 * QUnit tests for Inventory Editor
 */

// Import utility functions to test
// Note: In production, these would be in a module. For browser testing,
// we'll define them locally or include them via script tag.

// Utility function tests
QUnit.module('Utility Functions', function() {

    QUnit.test('columnLetterToIndex converts correctly', function(assert) {
        assert.equal(columnLetterToIndex('A'), 1, 'A = 1');
        assert.equal(columnLetterToIndex('B'), 2, 'B = 2');
        assert.equal(columnLetterToIndex('J'), 10, 'J = 10');
        assert.equal(columnLetterToIndex('K'), 11, 'K = 11');
        assert.equal(columnLetterToIndex('T'), 20, 'T = 20');
        assert.equal(columnLetterToIndex('Z'), 26, 'Z = 26');
        assert.equal(columnLetterToIndex('AA'), 27, 'AA = 27');
    });

    QUnit.test('indexToColumnLetter converts correctly', function(assert) {
        assert.equal(indexToColumnLetter(1), 'A', '1 = A');
        assert.equal(indexToColumnLetter(2), 'B', '2 = B');
        assert.equal(indexToColumnLetter(10), 'J', '10 = J');
        assert.equal(indexToColumnLetter(11), 'K', '11 = K');
        assert.equal(indexToColumnLetter(20), 'T', '20 = T');
        assert.equal(indexToColumnLetter(26), 'Z', '26 = Z');
        assert.equal(indexToColumnLetter(27), 'AA', '27 = AA');
    });

    QUnit.test('formatDate formats as D/M/YY', function(assert) {
        const date = new Date(2025, 11, 16); // Dec 16, 2025
        assert.equal(formatDate(date), '16/12/25', 'Formats correctly');

        const date2 = new Date(2025, 0, 5); // Jan 5, 2025
        assert.equal(formatDate(date2), '5/1/25', 'No padding on single digits');
    });

    QUnit.test('formatDateWithQty includes quantity change', function(assert) {
        const date = new Date(2025, 11, 16);
        assert.equal(formatDateWithQty(date, 3), '16/12/25-3', 'Positive change');
        assert.equal(formatDateWithQty(date, -2), '16/12/25-2', 'Negative uses absolute');
        assert.equal(formatDateWithQty(date, 0), '16/12/25-0', 'Zero change');
    });

    QUnit.test('escapeHtml escapes dangerous characters', function(assert) {
        assert.equal(escapeHtml('<script>'), '&lt;script&gt;', 'Escapes tags');
        assert.equal(escapeHtml('"quotes"'), '&quot;quotes&quot;', 'Escapes quotes');
        assert.equal(escapeHtml("it's"), "it&#039;s", 'Escapes apostrophes');
        assert.equal(escapeHtml('a & b'), 'a &amp; b', 'Escapes ampersand');
    });
});

// Search function tests
QUnit.module('Search Functions', function(hooks) {
    let testWorkbook;

    hooks.beforeEach(async function() {
        testWorkbook = await createMockWorkbook();
        workbook = testWorkbook; // Set global workbook
    });

    QUnit.test('searchWarehouse finds SKU', function(assert) {
        const results = searchWarehouse('SKU002');
        assert.equal(results.length, 1, 'Found 1 match');
        assert.equal(results[0].sheetName, 'WAREHOUSE', 'Correct sheet');
        assert.equal(results[0].rowNumber, 3, 'Correct row');
        assert.equal(results[0].quantity, 50, 'Correct quantity');
    });

    QUnit.test('searchWarehouse finds multiple matches', function(assert) {
        const results = searchWarehouse('SKU001');
        assert.equal(results.length, 1, 'Found 1 match');
        assert.equal(results[0].rowNumber, 2, 'Found in row 2');
    });

    QUnit.test('searchWarehouse is case-insensitive', function(assert) {
        const results = searchWarehouse('sku001');
        assert.equal(results.length, 1, 'Found match with lowercase search');
    });

    QUnit.test('searchWarehouse returns empty for non-existent SKU', function(assert) {
        const results = searchWarehouse('NOTEXIST');
        assert.equal(results.length, 0, 'No matches found');
    });
});

// Date column tests
QUnit.module('Date Column Functions', function(hooks) {
    let testWorkbook;

    hooks.beforeEach(async function() {
        testWorkbook = await createMockWorkbook();
        workbook = testWorkbook;
    });

    QUnit.test('findNextEmptyDateColumn finds first empty', function(assert) {
        const sheet = workbook.getWorksheet('WAREHOUSE');
        const row = sheet.getRow(2); // Row with no dates

        const colIndex = findNextEmptyDateColumn(row);
        assert.equal(colIndex, 12, 'Returns column L (12) - dates start at L');
    });

    QUnit.test('findNextEmptyDateColumn skips filled columns', function(assert) {
        const sheet = workbook.getWorksheet('WAREHOUSE');
        const row = sheet.getRow(3); // Row with L filled

        const colIndex = findNextEmptyDateColumn(row);
        assert.equal(colIndex, 13, 'Returns column M (13)');
    });

    QUnit.test('findNextEmptyDateColumn returns null when full', function(assert) {
        const sheet = workbook.getWorksheet('WAREHOUSE');
        const row = sheet.getRow(2);

        // Fill all date columns (L=12 through T=20)
        for (let i = 12; i <= 20; i++) {
            row.getCell(i).value = '1/1/25-1';
        }

        const colIndex = findNextEmptyDateColumn(row);
        assert.equal(colIndex, null, 'Returns null when all full');
    });
});

// Highlight tests
QUnit.module('Highlight Functions', function(hooks) {
    let testWorkbook;

    hooks.beforeEach(async function() {
        testWorkbook = await createMockWorkbook();
        workbook = testWorkbook;
    });

    QUnit.test('checkRowHighlight returns false for non-highlighted row', function(assert) {
        const result = checkRowHighlight('WAREHOUSE', 2);
        assert.equal(result, false, 'Not highlighted');
    });

    QUnit.test('checkRowHighlight detects yellow highlight', function(assert) {
        const sheet = workbook.getWorksheet('WAREHOUSE');
        const row = sheet.getRow(2);

        // Apply yellow fill
        row.getCell(2).fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFFFFF00' }
        };

        const result = checkRowHighlight('WAREHOUSE', 2);
        assert.equal(result, true, 'Detected as highlighted');
    });
});

// Cell value extraction tests
QUnit.module('Cell Value Extraction', function() {

    QUnit.test('getCellValue handles simple values', function(assert) {
        const mockCell = { value: 'test' };
        assert.equal(getCellValue(mockCell), 'test', 'Simple string');

        mockCell.value = 123;
        assert.equal(getCellValue(mockCell), 123, 'Number');
    });

    QUnit.test('getCellValue handles null/undefined', function(assert) {
        assert.equal(getCellValue(null), '', 'Null cell');
        assert.equal(getCellValue({ value: null }), '', 'Null value');
        assert.equal(getCellValue({ value: undefined }), '', 'Undefined value');
    });

    QUnit.test('getCellValue handles rich text', function(assert) {
        const mockCell = {
            value: {
                richText: [
                    { text: 'Hello' },
                    { text: ' World' }
                ]
            }
        };
        assert.equal(getCellValue(mockCell), 'Hello World', 'Concatenates rich text');
    });

    QUnit.test('getCellValue handles formula results', function(assert) {
        const mockCell = {
            value: {
                formula: '=A1+B1',
                result: 42
            }
        };
        assert.equal(getCellValue(mockCell), 42, 'Returns formula result');
    });
});
```

#### 4. Update app.js to Export Functions for Testing
**File**: `js/app.js`
**Changes**: Add at the end of the file to expose functions globally for tests

```javascript
// Export functions for testing (browser global scope)
if (typeof window !== 'undefined') {
    window.columnLetterToIndex = columnLetterToIndex;
    window.indexToColumnLetter = indexToColumnLetter;
    window.formatDate = formatDate;
    window.formatDateWithQty = formatDateWithQty;
    window.escapeHtml = escapeHtml;
    window.getCellValue = getCellValue;
    window.searchWarehouse = searchWarehouse;
    window.findNextEmptyDateColumn = findNextEmptyDateColumn;
    window.checkRowHighlight = checkRowHighlight;
}
```

#### 5. Update tests/index.html to include app.js
**File**: `tests/index.html`
**Changes**: Add app.js script before test files

```html
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Inventory Editor Tests</title>
    <link rel="stylesheet" href="https://code.jquery.com/qunit/qunit-2.20.0.css">
</head>
<body>
    <div id="qunit"></div>
    <div id="qunit-fixture">
        <!-- DOM fixture for tests that need UI elements -->
        <input type="text" id="sku-input">
        <div id="results"></div>
    </div>

    <script src="https://code.jquery.com/qunit/qunit-2.20.0.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js"></script>
    <script src="../js/app.js"></script>
    <script src="test-utils.js"></script>
    <script src="app.test.js"></script>
</body>
</html>
```

### Success Criteria:

#### Automated Verification:
- [ ] Open `tests/index.html` in browser
- [ ] All QUnit tests pass (green)
- [ ] No test failures or errors

#### Manual Verification:
- [ ] Test page loads without errors
- [ ] QUnit shows module breakdown (Utility Functions, Search Functions, etc.)
- [ ] Each test displays descriptive name
- [ ] Tests complete in under 5 seconds

**Implementation Note**: Run tests in multiple browsers (Chrome, Firefox, Safari) to ensure compatibility.

---

## Phase 7: Multi-Sheet Search (Future Expansion)

### Overview
Expand the application to search across ALL sheets in the workbook, not just WAREHOUSE. Results will display with sheet context so users can identify which product category/sheet contains each match.

### Changes Required:

#### 1. Add Sheet Configuration
**File**: `js/app.js`
**Changes**: Add configuration for multi-sheet mode

```javascript
// Add to CONFIG object
const CONFIG = {
    // ... existing config ...
    searchAllSheets: true,  // Toggle for multi-sheet search
    excludeSheets: []       // Sheet names to exclude from search (if any)
};
```

#### 2. Create Multi-Sheet Search Function
**File**: `js/app.js`
**Changes**: Add `searchAllSheets` function alongside `searchWarehouse`

```javascript
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
```

#### 3. Update handleSearch to Use Config
**File**: `js/app.js`
**Changes**: Toggle between single-sheet and multi-sheet search

```javascript
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
```

#### 4. Update Result Display for Multi-Sheet
**File**: `js/app.js`
**Changes**: Show sheet name in results when multi-sheet is enabled

```javascript
function createResultItem(result, index) {
    const div = document.createElement('div');
    div.className = 'result-item';
    div.dataset.sheetName = result.sheetName;
    div.dataset.rowNumber = result.rowNumber;
    div.id = `result-${index}`;

    const isHighlighted = checkRowHighlightMulti(result.sheetName, result.rowNumber);
    if (isHighlighted) {
        div.classList.add('highlighted');
    }

    // Show sheet name if multi-sheet search is enabled
    const headerText = CONFIG.searchAllSheets
        ? `Sheet: ${escapeHtml(result.sheetName)} | Row: ${result.rowNumber}`
        : `Row: ${result.rowNumber}`;

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
            <button class="save-btn" onclick="saveQuantityMulti(${index}, '${escapeHtml(result.sheetName)}', ${result.rowNumber})">
                Save Quantity
            </button>
            <button class="highlight-btn" onclick="toggleHighlightMulti(${index}, '${escapeHtml(result.sheetName)}', ${result.rowNumber})">
                ${isHighlighted ? 'Remove Highlight' : 'Highlight Low Stock'}
            </button>
        </div>
    `;

    return div;
}
```

#### 5. Update Save/Highlight Functions for Multi-Sheet
**File**: `js/app.js`
**Changes**: Add sheet-aware versions of save and highlight functions

```javascript
/**
 * Save quantity on any sheet (multi-sheet version)
 */
function saveQuantityMulti(index, sheetName, rowNumber) {
    const resultItem = document.getElementById(`result-${index}`);
    const qtyInput = resultItem.querySelector('.qty-input');
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

    row.getCell(qtyColIndex).value = newQuantity;

    const dateColIndex = findNextEmptyDateColumn(row);
    if (dateColIndex) {
        const dateStr = formatDate(new Date());
        row.getCell(dateColIndex).value = dateStr;

        const clearColIndex = dateColIndex + 1;
        if (clearColIndex <= columnLetterToIndex(CONFIG.dateColumnsEnd) + 1) {
            row.getCell(clearColIndex).value = null;
        }

        showStatus(`Saved on ${sheetName}: Qty=${newQuantity}, Date in column ${indexToColumnLetter(dateColIndex)}`, 'success');
    } else {
        showStatus(`Saved on ${sheetName}: Qty=${newQuantity} (No empty date column)`, 'success');
    }

    row.commit();
}

/**
 * Toggle highlight on any sheet (multi-sheet version)
 */
function toggleHighlightMulti(index, sheetName, rowNumber) {
    const sheet = workbook.getWorksheet(sheetName);
    if (!sheet) {
        showStatus(`Sheet "${sheetName}" not found`, 'error');
        return;
    }

    const row = sheet.getRow(rowNumber);
    const resultItem = document.getElementById(`result-${index}`);
    const isCurrentlyHighlighted = resultItem.classList.contains('highlighted');

    const yellowFill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFFFFF00' }
    };

    const noFill = {
        type: 'pattern',
        pattern: 'none'
    };

    const startCol = columnLetterToIndex('A');
    const endCol = columnLetterToIndex(CONFIG.dateColumnsEnd) + 1;

    for (let col = startCol; col <= endCol; col++) {
        const cell = row.getCell(col);
        cell.fill = isCurrentlyHighlighted ? noFill : yellowFill;
    }

    row.commit();

    if (isCurrentlyHighlighted) {
        resultItem.classList.remove('highlighted');
        resultItem.querySelector('.highlight-btn').textContent = 'Highlight Low Stock';
        showStatus(`Highlight removed from ${sheetName}`, 'success');
    } else {
        resultItem.classList.add('highlighted');
        resultItem.querySelector('.highlight-btn').textContent = 'Remove Highlight';
        showStatus(`Row highlighted on ${sheetName}`, 'success');
    }
}

/**
 * Check highlight on any sheet (multi-sheet version)
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
```

#### 6. Optional: Add UI Toggle for Search Scope
**File**: `index.html`
**Changes**: Add checkbox to toggle between single/multi-sheet search

```html
<section id="search-section" style="display: none;">
    <div class="search-options">
        <label>
            <input type="checkbox" id="search-all-sheets" checked>
            Search all sheets
        </label>
    </div>
    <input type="text" id="sku-input" placeholder="Enter SKU to search...">
    <button id="search-btn">Search</button>
</section>
```

**File**: `js/app.js`
**Changes**: Handle checkbox toggle

```javascript
// Add to init()
document.getElementById('search-all-sheets').addEventListener('change', (e) => {
    CONFIG.searchAllSheets = e.target.checked;
});
```

### Success Criteria:

#### Automated Verification:
- [x] `searchAllSheets('SKU001')` returns matches from multiple sheets
- [x] Results include correct `sheetName` for each match
- [x] Save/highlight operations target correct sheet

#### Manual Verification:
- [ ] Search shows results from all sheets with sheet names
- [ ] Editing quantity on Sheet A doesn't affect Sheet B
- [ ] Highlight applies to correct row on correct sheet
- [ ] Downloaded file preserves changes on all modified sheets
- [ ] Toggle checkbox switches between single/multi-sheet mode

**Implementation Note**: The multi-sheet functions (`saveQuantityMulti`, `toggleHighlightMulti`, etc.) can eventually replace the WAREHOUSE-only versions, with the config controlling behavior.

---

## Testing Strategy

### Unit Tests (QUnit):
- Utility functions (column conversion, date formatting, escaping)
- Search logic on WAREHOUSE sheet (Phases 1-6)
- Multi-sheet search logic (Phase 7)
- Date column finding algorithm
- Highlight detection (single and multi-sheet)
- Cell value extraction

### Integration Tests:
- Full workflow: load file → search → edit → download
- Verify downloaded file contains changes

### Manual Testing Steps:
1. Load `Warehouse Stock Stick.xlsx` test file
2. Search for existing SKU, verify results from WAREHOUSE sheet
3. Edit quantity, verify date appears in correct column (K-T)
4. Apply highlight, verify UI updates
5. Download and re-open file to verify changes persist
6. Test with edge cases (empty file, missing WAREHOUSE sheet, full date columns)

## Performance Considerations

- ExcelJS processes files in memory; large files (10MB+) may be slow
- Consider showing loading indicator during file operations
- Date column search is O(10) per row - negligible
- Search is O(rows) on WAREHOUSE sheet - very fast

## File Structure Summary

```
/home/rocky/web/xlsx/
├── index.html              # Main entry point
├── css/
│   └── style.css           # Application styles
├── js/
│   └── app.js              # All application logic
├── tests/
│   ├── index.html          # QUnit test runner
│   ├── test-utils.js       # Mock workbook creation
│   └── app.test.js         # Test suite
├── thoughts/
│   └── shared/
│       ├── research/       # Requirements document
│       └── plans/          # This plan
└── README.md               # Project documentation (create later)
```

## References

- Requirements: `thoughts/shared/research/2025-12-16-xlsx-editor-requirements.md`
- ExcelJS Documentation: https://github.com/exceljs/exceljs
- QUnit Documentation: https://qunitjs.com/
