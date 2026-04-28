const xlsx = require('xlsx');

try {
    const wb = xlsx.readFile('expenses.xlsx');
    console.log('Available sheets:', wb.SheetNames);
    
    // Check if Ilyas sheet exists
    const ilyasSheet = wb.Sheets['Ilyas'];
    if (ilyasSheet) {
        const data = xlsx.utils.sheet_to_json(ilyasSheet);
        console.log('\nIlyas sheet data:', data.length, 'rows');
        if (data.length > 0) {
            console.log('Sample row:', data[0]);
            console.log('Headers:', Object.keys(data[0]));
        }
    } else {
        console.log('\nIlyas sheet not found');
    }
    
    // Check Expenses sheet structure
    const expensesSheet = wb.Sheets['Expenses'];
    if (expensesSheet) {
        const data = xlsx.utils.sheet_to_json(expensesSheet);
        console.log('\nExpenses sheet data:', data.length, 'rows');
        if (data.length > 0) {
            console.log('Sample row:', data[0]);
            console.log('Headers:', Object.keys(data[0]));
        }
    }
    
} catch (error) {
    console.error('Error reading Excel file:', error.message);
}
