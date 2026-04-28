const xlsx = require('xlsx');

try {
    const wb = xlsx.readFile('expenses.xlsx');
    const ws = wb.Sheets['Ilyas'];
    const data = xlsx.utils.sheet_to_json(ws);

    // Simulate the generateDynamicFilters function
    const headers = Object.keys(data[0] || {});
    const excludedColumns = ['Sr No', 'srNo', 'Date', 'date', 'Amount', 'amount', 'S No', 'AccNo', 'Mobile', 'ITS '];
    const filterHeaders = headers.filter(header => !excludedColumns.some(excluded => header.toLowerCase().includes(excluded.toLowerCase())));

    console.log('Ilyas sheet headers:', headers);
    console.log('Filters that should be generated:', filterHeaders);

    // Get unique values for each filterable column
    filterHeaders.forEach(header => {
        const uniqueValues = [...new Set(data.map(row => row[header]).filter(val => val && val.toString().trim()))];
        console.log(`\n${header} unique values (${uniqueValues.length}):`, uniqueValues.slice(0, 10));
        if (uniqueValues.length > 10) {
            console.log('... and', uniqueValues.length - 10, 'more');
        }
    });

} catch (error) {
    console.error('Error:', error.message);
}
