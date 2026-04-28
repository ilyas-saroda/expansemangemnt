const xlsx = require('xlsx');

try {
    const wb = xlsx.readFile('expenses.xlsx');
    console.log('Available sheets:', wb.SheetNames);
    
    // Check Burhanuddin sheet
    const burhanuddinSheet = wb.Sheets['Burhanuddin'];
    if (burhanuddinSheet) {
        const data = xlsx.utils.sheet_to_json(burhanuddinSheet);
        console.log('\nBurhanuddin sheet data:', data.length, 'rows');
        if (data.length > 0) {
            console.log('Sample row:', data[0]);
            console.log('Headers:', Object.keys(data[0]));
            
            // Check what filters should be generated
            const headers = Object.keys(data[0]);
            const excludedColumns = ['Sr No', 'srNo', 'Date', 'date', 'Amount', 'amount', 'S No', 'AccNo', 'Mobile', 'ITS '];
            const filterHeaders = headers.filter(header => !excludedColumns.some(excluded => header.toLowerCase().includes(excluded.toLowerCase())));
            
            console.log('Filters that should be generated:', filterHeaders);
            
            // Get unique values for each filterable column
            filterHeaders.forEach(header => {
                const uniqueValues = [...new Set(data.map(row => row[header]).filter(val => val && val.toString().trim()))];
                console.log(`\n${header} unique values (${uniqueValues.length}):`, uniqueValues.slice(0, 10));
            });
        } else {
            console.log('Burhanuddin sheet is empty!');
        }
    } else {
        console.log('\nBurhanuddin sheet not found');
    }
    
} catch (error) {
    console.error('Error reading Excel file:', error.message);
}
