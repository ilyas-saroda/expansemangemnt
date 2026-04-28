// Test script to verify dropdown functionality for 'Given' column
const xlsx = require('xlsx');

console.log('Testing dropdown functionality for "Given" column...');

// Read the Excel file
const workbook = xlsx.readFile('expenses.xlsx');
const worksheet = workbook.Sheets['Hadiya rumal'];
const data = xlsx.utils.sheet_to_json(worksheet);

console.log('Hadiya rumal sheet data:');
console.log('Columns:', Object.keys(data[0] || {}));
console.log('Number of rows:', data.length);

// Check if 'Given' column exists
const hasGivenColumn = Object.keys(data[0] || {}).includes('Given');
console.log('Has "Given" column:', hasGivenColumn);

// Show sample data from 'Given' column
if (hasGivenColumn) {
    console.log('\nSample "Given" column values:');
    data.slice(0, 5).forEach((row, i) => {
        console.log(`Row ${i+1}: Given = "${row.Given}"`);
    });
    
    // Get unique values for dropdown
    const uniqueGivenValues = [...new Set(data.map(row => row.Given).filter(val => val && val.trim() !== ''))];
    console.log('\nUnique "Given" values for dropdown:', uniqueGivenValues);
    console.log('Number of unique values:', uniqueGivenValues.length);
}

console.log('\nTest completed successfully!');
