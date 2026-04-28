const ExcelJS = require('exceljs');
const xlsx = require('xlsx');

async function testBurhanuddinFix() {
    try {
        // Read the existing workbook
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile('expenses.xlsx');
        
        // Check if Burhanuddin sheet exists
        const burhanuddinSheet = workbook.getWorksheet('Burhanuddin');
        
        if (burhanuddinSheet) {
            console.log('Burhanuddin sheet found');
            
            // Check if sheet has data
            const rowCount = burhanuddinSheet.rowCount;
            console.log('Row count:', rowCount);
            
            if (rowCount <= 1) { // Only header row or completely empty
                console.log('Sheet is empty, adding sample data...');
                
                // Add sample data to test filters
                burhanuddinSheet.addRow(['Sr No', 'Date', 'Given To', 'Amount', 'Mode', 'Description', 'Fund', 'Status']);
                burhanuddinSheet.addRow([1, '2026-04-28', 'Test Person 1', '1000.00', 'Cash', 'Test Description 1', 'Test Fund 1', 'Pending']);
                burhanuddinSheet.addRow([2, '2026-04-28', 'Test Person 2', '2000.00', 'Bank', 'Test Description 2', 'Test Fund 2', 'Paid']);
                
                // Save the workbook
                await workbook.xlsx.writeFile('expenses.xlsx');
                console.log('Sample data added to Burhanuddin sheet');
            } else {
                console.log('Sheet already has data');
            }
        } else {
            console.log('Burhanuddin sheet not found');
        }
        
        // Test with xlsx as well to verify
        const wb = xlsx.readFile('expenses.xlsx');
        const ws = wb.Sheets['Burhanuddin'];
        if (ws) {
            const data = xlsx.utils.sheet_to_json(ws);
            console.log('Final Burhanuddin sheet data:', data.length, 'rows');
            if (data.length > 0) {
                console.log('Headers:', Object.keys(data[0]));
                console.log('Sample data:', data.slice(0, 2));
            }
        }
        
    } catch (error) {
        console.error('Error:', error.message);
    }
}

testBurhanuddinFix();
