/**
 * QUnit tests for Inventory Editor
 */

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
        const result = checkRowHighlight(2);
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

        const result = checkRowHighlight(2);
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

// Multi-sheet search tests
QUnit.module('Multi-Sheet Search Functions', function(hooks) {
    let testWorkbook;

    hooks.beforeEach(async function() {
        testWorkbook = await createMockWorkbook();
        workbook = testWorkbook; // Set global workbook
    });

    QUnit.test('searchAllSheets finds SKU across all sheets', function(assert) {
        // SKU001 exists in both WAREHOUSE and CIGARS sheets
        const results = searchAllSheets('SKU001');
        assert.equal(results.length, 2, 'Found 2 matches across sheets');

        const sheetNames = results.map(r => r.sheetName).sort();
        assert.deepEqual(sheetNames, ['CIGARS', 'WAREHOUSE'], 'Found in correct sheets');
    });

    QUnit.test('searchAllSheets finds SKU in single sheet only', function(assert) {
        // SKU003 only exists in CIGARS sheet
        const results = searchAllSheets('SKU003');
        assert.equal(results.length, 1, 'Found 1 match');
        assert.equal(results[0].sheetName, 'CIGARS', 'Found in CIGARS sheet');
        assert.equal(results[0].quantity, 25, 'Correct quantity');
    });

    QUnit.test('searchAllSheets returns correct sheet context', function(assert) {
        const results = searchAllSheets('SKU001');

        const warehouseMatch = results.find(r => r.sheetName === 'WAREHOUSE');
        const cigarsMatch = results.find(r => r.sheetName === 'CIGARS');

        assert.ok(warehouseMatch, 'Found WAREHOUSE match');
        assert.equal(warehouseMatch.quantity, 100, 'WAREHOUSE quantity correct');
        assert.equal(warehouseMatch.rowNumber, 2, 'WAREHOUSE row correct');

        assert.ok(cigarsMatch, 'Found CIGARS match');
        assert.equal(cigarsMatch.quantity, 15, 'CIGARS quantity correct');
        assert.equal(cigarsMatch.rowNumber, 3, 'CIGARS row correct');
    });

    QUnit.test('searchAllSheets is case-insensitive', function(assert) {
        const results = searchAllSheets('sku003');
        assert.equal(results.length, 1, 'Found match with lowercase');
    });

    QUnit.test('searchAllSheets returns empty for non-existent SKU', function(assert) {
        const results = searchAllSheets('NOTEXIST');
        assert.equal(results.length, 0, 'No matches found');
    });

    QUnit.test('searchAllSheets respects excludeSheets config', function(assert) {
        CONFIG.excludeSheets = ['CIGARS'];
        const results = searchAllSheets('SKU001');
        assert.equal(results.length, 1, 'Only 1 match after excluding CIGARS');
        assert.equal(results[0].sheetName, 'WAREHOUSE', 'Match from WAREHOUSE only');
        CONFIG.excludeSheets = []; // Reset
    });
});

// Multi-sheet highlight tests
QUnit.module('Multi-Sheet Highlight Functions', function(hooks) {
    let testWorkbook;

    hooks.beforeEach(async function() {
        testWorkbook = await createMockWorkbook();
        workbook = testWorkbook;
    });

    QUnit.test('checkRowHighlightMulti returns false for non-highlighted row', function(assert) {
        const result = checkRowHighlightMulti('CIGARS', 2);
        assert.equal(result, false, 'Not highlighted');
    });

    QUnit.test('checkRowHighlightMulti detects yellow highlight on any sheet', function(assert) {
        const sheet = workbook.getWorksheet('CIGARS');
        const row = sheet.getRow(2);

        // Apply yellow fill
        row.getCell(2).fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFFFFF00' }
        };

        const result = checkRowHighlightMulti('CIGARS', 2);
        assert.equal(result, true, 'Detected as highlighted');

        // Verify WAREHOUSE is unaffected
        const warehouseResult = checkRowHighlightMulti('WAREHOUSE', 2);
        assert.equal(warehouseResult, false, 'WAREHOUSE not affected');
    });

    QUnit.test('checkRowHighlightMulti returns false for non-existent sheet', function(assert) {
        const result = checkRowHighlightMulti('NONEXISTENT', 2);
        assert.equal(result, false, 'Returns false for missing sheet');
    });
});
