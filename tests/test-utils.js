/**
 * Test utilities for creating mock workbooks
 */

/**
 * Create a mock workbook with test data
 * @returns {ExcelJS.Workbook}
 */
async function createMockWorkbook() {
    const wb = new ExcelJS.Workbook();

    // Create WAREHOUSE sheet
    const warehouse = wb.addWorksheet('WAREHOUSE');
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

    // Create a second sheet (CIGARS) for multi-sheet testing
    const cigars = wb.addWorksheet('CIGARS');
    cigars.getCell('A1').value = 'Product';
    cigars.getCell('B1').value = 'SKU';
    cigars.getCell('J1').value = 'Quantity';
    cigars.getCell('G1').value = '';

    // Add test data to second sheet
    cigars.getCell('B2').value = 'SKU003';
    cigars.getCell('G2').value = 'Premium Cigar';
    cigars.getCell('J2').value = 25;

    cigars.getCell('B3').value = 'SKU001';  // Same SKU as WAREHOUSE for cross-sheet test
    cigars.getCell('G3').value = 'Special Cigar';
    cigars.getCell('J3').value = 15;

    // Create a third sheet (ACCESSORIES) for broader testing
    const accessories = wb.addWorksheet('ACCESSORIES');
    accessories.getCell('A1').value = 'Product';
    accessories.getCell('B1').value = 'SKU';
    accessories.getCell('J1').value = 'Quantity';

    accessories.getCell('B2').value = 'ACC001';
    accessories.getCell('G2').value = 'Lighter';
    accessories.getCell('J2').value = 200;

    return wb;
}
