const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const ExcelJS = require('exceljs');
const cors = require('cors');
const path = require('path');
const fs = require('fs');

const app = express();
const PORT = 3000;

// Data storage files
const EXCEL_FILE = 'expenses.xlsx';
const JSON_FILE = 'expenses.json';

// Global variable to store current data
let currentExpenses = [];

// Middleware
app.use(cors());
app.use(express.json());
app.use(express.static('public'));

// Configure multer for file uploads
const storage = multer.diskStorage({
  destination: function (req, file, cb) {
    cb(null, 'uploads/');
  },
  filename: function (req, file, cb) {
    cb(null, Date.now() + '-' + file.originalname);
  }
});

const upload = multer({
  storage: storage,
  limits: {
    fileSize: 5 * 1024 * 1024 // 5MB limit
  },
  fileFilter: function (req, file, cb) {
    // Accept only Excel files
    if (file.mimetype === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' || 
        file.mimetype === 'application/vnd.ms-excel') {
      cb(null, true);
    } else {
      cb(new Error('Invalid file type. Only Excel files are allowed.'));
    }
  }
});

// Create uploads directory if it doesn't exist
if (!fs.existsSync('uploads')) {
  fs.mkdirSync('uploads');
}

// Excel file functions
function createExcelFile(expenses) {
  try {
    const worksheet = xlsx.utils.json_to_sheet(expenses);
    const workbook = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(workbook, worksheet, 'Expenses');
    xlsx.writeFile(workbook, EXCEL_FILE);
    console.log(`Excel file created/updated: ${EXCEL_FILE}`);
    return true;
  } catch (error) {
    console.error('Error creating Excel file:', error);
    return false;
  }
}

function updateExcelFile(expenses) {
  return createExcelFile(expenses);
}

// Multi-sheet Excel management functions
function createMultiSheetWorkbook() {
  try {
    if (fs.existsSync(EXCEL_FILE)) {
      // Check for lock file
      const lockFile = '~$' + EXCEL_FILE.split('/').pop();
      const lockFilePath = EXCEL_FILE.replace(/[^/]+$/, lockFile);
      
      if (fs.existsSync(lockFilePath)) {
        console.warn('Excel file is locked, attempting to read anyway...');
        try {
          // Try to read the file despite lock
          return xlsx.readFile(EXCEL_FILE);
        } catch (lockError) {
          console.error('Cannot read locked file, creating new workbook:', lockError.message);
          // Create new workbook as fallback
          const workbook = xlsx.utils.book_new();
          const emptyWorksheet = xlsx.utils.json_to_sheet([]);
          xlsx.utils.book_append_sheet(workbook, emptyWorksheet, 'Expenses');
          return workbook;
        }
      } else {
        // Load existing workbook
        return xlsx.readFile(EXCEL_FILE);
      }
    } else {
      // Create new workbook with default Expenses sheet
      const workbook = xlsx.utils.book_new();
      const emptyWorksheet = xlsx.utils.json_to_sheet([]);
      xlsx.utils.book_append_sheet(workbook, emptyWorksheet, 'Expenses');
      return workbook;
    }
  } catch (error) {
    console.error('Error creating/loading workbook:', error);
    // Return new workbook as fallback
    const workbook = xlsx.utils.book_new();
    const emptyWorksheet = xlsx.utils.json_to_sheet([]);
    xlsx.utils.book_append_sheet(workbook, emptyWorksheet, 'Expenses');
    return workbook;
  }
}

function addWorksheetToWorkbook(workbook, sheetName, sheetData = []) {
  try {
    // Check if sheet already exists
    if (workbook.SheetNames.includes(sheetName)) {
      throw new Error(`Sheet "${sheetName}" already exists`);
    }

    // Create new worksheet
    const worksheet = xlsx.utils.json_to_sheet(sheetData);
    xlsx.utils.book_append_sheet(workbook, worksheet, sheetName);
    
    return workbook;
  } catch (error) {
    console.error('Error adding worksheet:', error);
    throw error;
  }
}

function saveWorkbook(workbook) {
  try {
    // Check if file is locked by another program
    const tempFile = EXCEL_FILE.replace('.xlsx', '.tmp');
    
    // Try to save to temporary file first
    xlsx.writeFile(workbook, tempFile);
    
    // If successful, replace the original file
    if (fs.existsSync(tempFile)) {
      // Remove original file if it exists
      if (fs.existsSync(EXCEL_FILE)) {
        fs.unlinkSync(EXCEL_FILE);
      }
      
      // Rename temp file to original
      fs.renameSync(tempFile, EXCEL_FILE);
      console.log(`Workbook saved: ${EXCEL_FILE}`);
      return true;
    }
    
    return false;
  } catch (error) {
    console.error('Error saving workbook:', error);
    
    // Try alternative approach - save with timestamp
    try {
      const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
      const backupFile = EXCEL_FILE.replace('.xlsx', `_${timestamp}.xlsx`);
      xlsx.writeFile(workbook, backupFile);
      console.log(`Workbook saved as backup: ${backupFile}`);
      
      // Try to replace original after a short delay
      setTimeout(() => {
        try {
          if (fs.existsSync(backupFile) && fs.existsSync(EXCEL_FILE)) {
            fs.unlinkSync(EXCEL_FILE);
          }
          if (fs.existsSync(backupFile)) {
            fs.renameSync(backupFile, EXCEL_FILE);
            console.log(`Workbook restored: ${EXCEL_FILE}`);
          }
        } catch (restoreError) {
          console.error('Error restoring workbook:', restoreError);
        }
      }, 1000);
      
      return true;
    } catch (backupError) {
      console.error('Backup save also failed:', backupError);
      return false;
    }
  }
}

async function getWorkbookSheets() {
  try {
    if (!fs.existsSync(EXCEL_FILE)) {
      return [];
    }
    
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(EXCEL_FILE);
    
    return workbook.worksheets.map(worksheet => ({
      name: worksheet.name,
      isDefault: worksheet.name === 'Expenses'
    }));
  } catch (error) {
    console.error('Error reading workbook sheets:', error);
    return [];
  }
}

function saveToJSON(expenses) {
  try {
    fs.writeFileSync(JSON_FILE, JSON.stringify(expenses, null, 2));
    return true;
  } catch (error) {
    console.error('Error saving to JSON:', error);
    return false;
  }
}

// Route to handle Excel file upload
app.post('/upload-excel', upload.single('excelFile'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }

    const { sheetName = 'Expenses' } = req.body;
    const filePath = req.file.path;
    
    // Read the Excel file
    const workbook = xlsx.readFile(filePath);
    const sourceSheetName = workbook.SheetNames[0];
    const sourceWorksheet = workbook.Sheets[sourceSheetName];
    
    // Get the range to determine if first row is headers
    const range = xlsx.utils.decode_range(sourceWorksheet['!ref'] || 'A1:A1');
    const hasHeaders = range.s.r < range.e.r; // If there's more than one row, assume first is headers
    
    let processedData;
    let headers = [];
    
    if (hasHeaders) {
      // Read first row as headers
      const headerData = xlsx.utils.sheet_to_json(sourceWorksheet, { header: 1, range: 0 });
      headers = headerData[0] || [];
      
      // Read data rows starting from second row
      const dataRows = xlsx.utils.sheet_to_json(sourceWorksheet, { header: headers, range: 1 });
      
      // Process data dynamically based on headers
      processedData = dataRows.map((row, index) => {
        const processedRow = { srNo: index + 1 };
        
        headers.forEach(header => {
          if (header && header.trim()) {
            const key = header.trim();
            let value = row[key] || '';
            
            // Special formatting for common fields
            if (key.toLowerCase().includes('date') && value) {
              value = formatExcelDate(value);
            } else if (key.toLowerCase().includes('amount') || key.toLowerCase().includes('price')) {
              value = parseFloat(value || 0).toFixed(2);
            }
            
            processedRow[key] = value;
          }
        });
        
        return processedRow;
      });
    } else {
      // No headers - use default expense structure
      const data = xlsx.utils.sheet_to_json(sourceWorksheet);
      processedData = data.map((row, index) => ({
        srNo: index + 1,
        date: formatExcelDate(row.Date || row.date || new Date()),
        givenTo: row['Given To'] || row.givenTo || row['Given to'] || '',
        amount: parseFloat(row.Amount || row.amount || 0).toFixed(2),
        mode: row.Mode || row.mode || '',
        description: row.Description || row.description || '',
        fund: row.Fund || row.fund || '',
        status: row.Status || row.status || 'Pending'
      }));
      
      headers = ['Sr No', 'Date', 'Given To', 'Amount', 'Mode', 'Description', 'Fund', 'Status'];
    }
    
    // Delete the uploaded file after processing
    fs.unlinkSync(filePath);
    
    // Load the main workbook to import to specific sheet
    let mainWorkbook;
    if (fs.existsSync(EXCEL_FILE)) {
      mainWorkbook = new ExcelJS.Workbook();
      await mainWorkbook.xlsx.readFile(EXCEL_FILE);
    } else {
      mainWorkbook = new ExcelJS.Workbook();
      mainWorkbook.addWorksheet('Expenses');
    }
    
    // Get or create the target worksheet
    let targetWorksheet = mainWorkbook.getWorksheet(sheetName);
    if (!targetWorksheet) {
      targetWorksheet = mainWorkbook.addWorksheet(sheetName);
    }
    
    // Clear existing data (keep headers if they exist)
    targetWorksheet.eachRow((row, rowNumber) => {
      if (rowNumber > 1) {
        targetWorksheet.spliceRows(rowNumber, 1);
      }
    });
    
    // Add headers if sheet is empty or has different structure
    if (targetWorksheet.rowCount <= 1) {
      const headerColumns = headers.map(header => ({
        header: header,
        key: header.toLowerCase().replace(/[^a-z0-9]/g, '_'),
        width: 15
      }));
      targetWorksheet.columns = headerColumns;
    }
    
    // Add processed data to the worksheet
    processedData.forEach(row => {
      const rowData = headers.map(header => row[header] || '');
      targetWorksheet.addRow(rowData);
    });
    
    // Save the workbook
    await mainWorkbook.xlsx.writeFile(EXCEL_FILE);
    
    // Update global data if it's the current sheet
    if (sheetName === currentSheet) {
      currentExpenses = processedData;
    }
    
    res.json({
      success: true,
      data: processedData,
      headers: headers,
      sheetName: sheetName,
      message: `Successfully imported ${processedData.length} records to sheet "${sheetName}"`
    });
    
  } catch (error) {
    console.error('Error processing file:', error);
    
    // Delete file if it exists and there was an error
    if (req.file && fs.existsSync(req.file.path)) {
      fs.unlinkSync(req.file.path);
    }
    
    res.status(500).json({ 
      error: 'Error processing file: ' + error.message 
    });
  }
});

// Helper function to format Excel dates
function formatExcelDate(dateValue) {
  if (!dateValue) return new Date().toISOString().split('T')[0];
  
  // Handle Excel serial number dates
  if (typeof dateValue === 'number') {
    const excelDate = new Date((dateValue - 25569) * 86400 * 1000);
    return excelDate.toISOString().split('T')[0];
  }
  
  // Handle string dates
  const date = new Date(dateValue);
  if (isNaN(date.getTime())) {
    return new Date().toISOString().split('T')[0];
  }
  
  return date.toISOString().split('T')[0];
}

// API endpoints for data management

// GET /data - Fetch current expenses
app.get('/data', async (req, res) => {
  try {
    const { sheetName } = req.query;
    
    // Get available sheets
    const availableSheets = await getWorkbookSheets();
    
    // If no sheet specified, use default 'Expenses' or first available
    const selectedSheet = sheetName || 'Expenses';
    
    let data = [];
    let headers = [];
    
    if (fs.existsSync(EXCEL_FILE)) {
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(EXCEL_FILE);
      
      const worksheet = workbook.getWorksheet(selectedSheet);
      if (worksheet) {
        // Get headers from the first row
        headers = [];
        const headerRow = worksheet.getRow(1);
        headerRow.eachCell((cell, colNumber) => {
          headers.push(cell.value || `Column${colNumber}`);
        });
        
        // Read data rows dynamically based on headers
        data = [];
        worksheet.eachRow((row, rowNumber) => {
          if (rowNumber > 1) { // Skip header row
            const rowData = { srNo: rowNumber - 1 };
            
            headers.forEach((header, colIndex) => {
              const cellValue = row.getCell(colIndex + 1).value;
              rowData[header] = cellValue || '';
            });
            
            data.push(rowData);
          }
        });
      }
    }
    
    res.json({
      success: true,
      data: data,
      headers: headers,
      count: data.length,
      currentSheet: selectedSheet,
      availableSheets: availableSheets
    });
  } catch (error) {
    res.status(500).json({
      success: false,
      error: 'Failed to fetch data: ' + error.message
    });
  }
});

// POST /update-row - Update a specific row in a sheet (targeted update)
app.post('/update-row', async (req, res) => {
  try {
    const { sheetName = 'Expenses', rowIndex, rowData, uniqueId } = req.body;
    
    if (!rowIndex || rowIndex < 2) {
      return res.status(400).json({
        success: false,
        error: 'Invalid row index. Row index must be >= 2 (row 1 is headers)'
      });
    }
    
    if (!fs.existsSync(EXCEL_FILE)) {
      return res.status(404).json({
        success: false,
        error: 'Excel file not found'
      });
    }

    // Load the existing workbook with ExcelJS to preserve all sheets
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(EXCEL_FILE);
    
    // Get the target worksheet
    const worksheet = workbook.getWorksheet(sheetName);
    if (!worksheet) {
      return res.status(404).json({
        success: false,
        error: `Sheet "${sheetName}" not found`
      });
    }

    // Get headers from the first row dynamically
    const headers = [];
    const headerRow = worksheet.getRow(1);
    headerRow.eachCell((cell, colNumber) => {
      headers.push(cell.value || `Column${colNumber}`);
    });

    // Get the specific row to update
    const targetRow = worksheet.getRow(rowIndex);
    if (!targetRow) {
      return res.status(404).json({
        success: false,
        error: `Row ${rowIndex} not found in sheet "${sheetName}"`
      });
    }

    // Update the specific row with new data
    headers.forEach((header, colIndex) => {
      const cellValue = rowData[header] || rowData[header.toLowerCase()] || '';
      targetRow.getCell(colIndex + 1).value = cellValue;
    });

    // Save the workbook
    await workbook.xlsx.writeFile(EXCEL_FILE);
    
    res.json({
      success: true,
      message: `Successfully updated row ${rowIndex} in sheet "${sheetName}"`,
      sheetName: sheetName,
      rowIndex: rowIndex,
      headers: headers
    });
  } catch (error) {
    console.error('Error updating row:', error);
    res.status(500).json({
      success: false,
      error: 'Failed to update row: ' + error.message
    });
  }
});

// POST /update - Update expense data for specific sheet (legacy - for bulk operations)
app.post('/update', async (req, res) => {
  try {
    const { sheetName = 'Expenses', updatedData } = req.body;
    
    if (!Array.isArray(updatedData)) {
      return res.status(400).json({
        success: false,
        error: 'Invalid data format. Expected array of expenses'
      });
    }
    
    if (!fs.existsSync(EXCEL_FILE)) {
      return res.status(404).json({
        success: false,
        error: 'Excel file not found'
      });
    }

    // Load the existing workbook with ExcelJS to preserve all sheets
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(EXCEL_FILE);
    
    // Get the target worksheet
    const worksheet = workbook.getWorksheet(sheetName);
    if (!worksheet) {
      return res.status(404).json({
        success: false,
        error: `Sheet "${sheetName}" not found`
      });
    }

    // Get headers from the first row dynamically
    const headers = [];
    const headerRow = worksheet.getRow(1);
    headerRow.eachCell((cell, colNumber) => {
      headers.push(cell.value || `Column${colNumber}`);
    });

    // Clear existing data in the worksheet (except headers)
    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber > 1) { // Keep header row, remove data rows
        worksheet.spliceRows(rowNumber, 1);
      }
    });

    // Add updated data back to the worksheet dynamically
    updatedData.forEach(expense => {
      const rowData = headers.map(header => {
        // Handle special fields with fallbacks
        if (header === 'Sr No' || header === 'srNo') {
          return expense.srNo || expense[header] || '';
        } else if (header === 'Date' || header === 'date') {
          return expense.date || expense[header] || new Date().toISOString().split('T')[0];
        } else if (header === 'Amount' || header === 'amount') {
          return expense.amount || expense[header] || '0.00';
        } else if (header === 'Status' || header === 'status') {
          return expense.status || expense[header] || 'Pending';
        } else {
          return expense[header] || '';
        }
      });
      worksheet.addRow(rowData);
    });

    // Save the workbook - this preserves all other sheets
    await workbook.xlsx.writeFile(EXCEL_FILE);
    
    // Update global data if it's the current sheet
    if (sheetName === currentSheet) {
      currentExpenses = updatedData;
    }
    
    res.json({
      success: true,
      message: `Successfully updated ${updatedData.length} records in sheet "${sheetName}"`,
      sheetName: sheetName,
      recordsUpdated: updatedData.length,
      headers: headers
    });
  } catch (error) {
    console.error('Error updating data:', error);
    res.status(500).json({
      success: false,
      error: 'Failed to update data: ' + error.message
    });
  }
});

// POST /add-record - Add a dynamic record to any sheet
app.post('/add-record', async (req, res) => {
  try {
    const { sheetName = 'Expenses', recordData } = req.body;
    
    if (!recordData || typeof recordData !== 'object') {
      return res.status(400).json({
        success: false,
        error: 'Invalid record data. Expected object with field values'
      });
    }
    
    if (!fs.existsSync(EXCEL_FILE)) {
      return res.status(404).json({
        success: false,
        error: 'Excel file not found'
      });
    }

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(EXCEL_FILE);
    
    const worksheet = workbook.getWorksheet(sheetName);
    if (!worksheet) {
      return res.status(404).json({
        success: false,
        error: `Sheet "${sheetName}" not found`
      });
    }

    // Get headers from the first row dynamically
    const headers = [];
    const headerRow = worksheet.getRow(1);
    headerRow.eachCell((cell, colNumber) => {
      headers.push(cell.value || `Column${colNumber}`);
    });

    // Create row data in the same order as headers
    const rowData = headers.map(header => {
      const value = recordData[header] || recordData[header.toLowerCase()] || '';
      
      // Special handling for different field types
      if (header.toLowerCase().includes('amount') || header.toLowerCase().includes('price')) {
        return parseFloat(value) || 0;
      } else if (header.toLowerCase().includes('date')) {
        return value || new Date().toISOString().split('T')[0];
      } else if (header.toLowerCase().includes('status')) {
        return value || 'Pending';
      }
      
      return value;
    });

    // Add new row with dynamic data
    worksheet.addRow(rowData);

    // Save the workbook
    await workbook.xlsx.writeFile(EXCEL_FILE);
    
    res.json({
      success: true,
      message: `Record added successfully to sheet "${sheetName}"`,
      recordData: recordData,
      headers: headers
    });
  } catch (error) {
    console.error('Error adding record:', error);
    res.status(500).json({
      success: false,
      error: 'Failed to add record: ' + error.message
    });
  }
});

// DELETE /delete-row - Delete a specific row in a sheet (targeted deletion)
app.delete('/delete-row', async (req, res) => {
  try {
    const { sheetName = 'Expenses', rowIndex, uniqueId } = req.body;
    
    if (!rowIndex || rowIndex < 2) {
      return res.status(400).json({
        success: false,
        error: 'Invalid row index. Row index must be >= 2 (row 1 is headers)'
      });
    }
    
    if (!fs.existsSync(EXCEL_FILE)) {
      return res.status(404).json({
        success: false,
        error: 'Excel file not found'
      });
    }

    // Load the existing workbook with ExcelJS to preserve all sheets
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(EXCEL_FILE);
    
    // Get the target worksheet
    const worksheet = workbook.getWorksheet(sheetName);
    if (!worksheet) {
      return res.status(404).json({
        success: false,
        error: `Sheet "${sheetName}" not found`
      });
    }

    // Check if the row exists
    const targetRow = worksheet.getRow(rowIndex);
    if (!targetRow || targetRow.number > worksheet.rowCount) {
      return res.status(404).json({
        success: false,
        error: `Row ${rowIndex} not found in sheet "${sheetName}"`
      });
    }

    // Delete the specific row using spliceRows
    worksheet.spliceRows(rowIndex, 1);

    // Save the workbook with proper file lock handling
    await workbook.xlsx.writeFile(EXCEL_FILE);
    
    res.json({
      success: true,
      message: `Successfully deleted row ${rowIndex} from sheet "${sheetName}"`,
      sheetName: sheetName,
      rowIndex: rowIndex
    });
  } catch (error) {
    console.error('Error deleting row:', error);
    res.status(500).json({
      success: false,
      error: 'Failed to delete row: ' + error.message
    });
  }
});

// DELETE /delete/:srNo - Delete expense (legacy - for compatibility)
app.delete('/delete/:srNo', (req, res) => {
  try {
    const srNo = parseInt(req.params.srNo);
    
    // Find and remove expense
    const index = currentExpenses.findIndex(exp => exp.srNo === srNo);
    
    if (index === -1) {
      return res.status(404).json({
        success: false,
        error: 'Expense not found'
      });
    }
    
    const deletedExpense = currentExpenses.splice(index, 1)[0];
    
    // Re-number remaining expenses
    currentExpenses.forEach((exp, i) => {
      exp.srNo = i + 1;
    });
    
    // Update Excel file and JSON
    const excelSuccess = updateExcelFile(currentExpenses);
    const jsonSuccess = saveToJSON(currentExpenses);
    
    if (excelSuccess && jsonSuccess) {
      res.json({
        success: true,
        data: deletedExpense,
        message: 'Expense deleted and Excel file updated'
      });
    } else {
      res.status(500).json({
        success: false,
        error: 'Failed to update files'
      });
    }
  } catch (error) {
    res.status(500).json({
      success: false,
      error: 'Failed to delete expense: ' + error.message
    });
  }
});

// API endpoints for sheet management

// GET /sheets - Get all sheets in the workbook
app.get('/sheets', async (req, res) => {
  try {
    const sheets = await getWorkbookSheets();
    res.json({
      success: true,
      sheets: sheets,
      totalSheets: sheets.length
    });
  } catch (error) {
    res.status(500).json({
      success: false,
      error: 'Failed to get sheets: ' + error.message
    });
  }
});

// POST /sheets - Create a new sheet
app.post('/sheets', async (req, res) => {
  try {
    const { sheetName, description, initialData = [] } = req.body;
    
    if (!sheetName || sheetName.trim() === '') {
      return res.status(400).json({
        success: false,
        error: 'Sheet name is required'
      });
    }

    // Sanitize sheet name (remove invalid characters)
    const sanitizedSheetName = sheetName.trim().replace(/[\\/*?:[\]]/g, '_').substring(0, 31);
    
    if (sanitizedSheetName !== sheetName.trim()) {
      return res.status(400).json({
        success: false,
        error: 'Sheet name contains invalid characters or is too long (max 31 chars)'
      });
    }

    // Create or load workbook
    const workbook = createMultiSheetWorkbook();
    
    // Check if sheet already exists
    if (workbook.SheetNames.includes(sanitizedSheetName)) {
      return res.status(400).json({
        success: false,
        error: `Sheet "${sanitizedSheetName}" already exists`
      });
    }

    // Add worksheet with initial data (if provided)
    const worksheetData = initialData.length > 0 ? initialData : [{
      srNo: 1,
      date: new Date().toISOString().split('T')[0],
      givenTo: '',
      amount: '0.00',
      mode: 'Cash',
      description: '',
      fund: '',
      status: 'Pending'
    }];
    
    addWorksheetToWorkbook(workbook, sanitizedSheetName, worksheetData);
    
    // Save workbook
    const saved = saveWorkbook(workbook);
    
    if (saved) {
      res.json({
        success: true,
        message: `Sheet "${sanitizedSheetName}" created successfully`,
        sheet: {
          name: sanitizedSheetName,
          description: description || '',
          createdAt: new Date().toISOString(),
          initialRecords: worksheetData.length
        },
        allSheets: await getWorkbookSheets()
      });
    } else {
      res.status(500).json({
        success: false,
        error: 'Failed to save workbook'
      });
    }
  } catch (error) {
    res.status(500).json({
      success: false,
      error: 'Failed to create sheet: ' + error.message
    });
  }
});

// DELETE /sheets/:sheetName - Delete a sheet
app.delete('/sheets/:sheetName', async (req, res) => {
  try {
    const sheetName = req.params.sheetName;
    
    if (sheetName === 'Expenses') {
      return res.status(400).json({
        success: false,
        error: 'Cannot delete the default Expenses sheet'
      });
    }

    const workbook = createMultiSheetWorkbook();
    
    if (!workbook.SheetNames.includes(sheetName)) {
      return res.status(404).json({
        success: false,
        error: `Sheet "${sheetName}" not found`
      });
    }

    // Remove the sheet
    const sheetIndex = workbook.SheetNames.indexOf(sheetName);
    workbook.SheetNames.splice(sheetIndex, 1);
    delete workbook.Sheets[sheetName];

    // Save workbook
    const saved = saveWorkbook(workbook);
    
    if (saved) {
      res.json({
        success: true,
        message: `Sheet "${sheetName}" deleted successfully`,
        remainingSheets: await getWorkbookSheets()
      });
    } else {
      res.status(500).json({
        success: false,
        error: 'Failed to save workbook'
      });
    }
  } catch (error) {
    res.status(500).json({
      success: false,
      error: 'Failed to delete sheet: ' + error.message
    });
  }
});

// GET /sheets/:sheetName/data - Get data from a specific sheet
app.get('/sheets/:sheetName/data', (req, res) => {
  try {
    const sheetName = req.params.sheetName;
    
    if (!fs.existsSync(EXCEL_FILE)) {
      return res.status(404).json({
        success: false,
        error: 'Excel file not found'
      });
    }

    const workbook = xlsx.readFile(EXCEL_FILE);
    
    if (!workbook.SheetNames.includes(sheetName)) {
      return res.status(404).json({
        success: false,
        error: `Sheet "${sheetName}" not found`
      });
    }

    const worksheet = workbook.Sheets[sheetName];
    const data = xlsx.utils.sheet_to_json(worksheet);
    
    res.json({
      success: true,
      sheetName: sheetName,
      data: data,
      recordCount: data.length
    });
  } catch (error) {
    res.status(500).json({
      success: false,
      error: 'Failed to get sheet data: ' + error.message
    });
  }
});

// POST /sheets/:sheetName/data - Update data in a specific sheet
app.post('/sheets/:sheetName/data', (req, res) => {
  try {
    const sheetName = req.params.sheetName;
    const updatedData = req.body;
    
    if (!Array.isArray(updatedData)) {
      return res.status(400).json({
        success: false,
        error: 'Invalid data format. Expected array of expenses'
      });
    }

    const workbook = createMultiSheetWorkbook();
    
    if (!workbook.SheetNames.includes(sheetName)) {
      return res.status(404).json({
        success: false,
        error: `Sheet "${sheetName}" not found`
      });
    }

    // Update the worksheet
    const worksheet = xlsx.utils.json_to_sheet(updatedData);
    workbook.Sheets[sheetName] = worksheet;

    // Save workbook
    const saved = saveWorkbook(workbook);
    
    if (saved) {
      res.json({
        success: true,
        message: `Sheet "${sheetName}" updated successfully`,
        recordCount: updatedData.length
      });
    } else {
      res.status(500).json({
        success: false,
        error: 'Failed to save workbook'
      });
    }
  } catch (error) {
    res.status(500).json({
      success: false,
      error: 'Failed to update sheet data: ' + error.message
    });
  }
});

// API endpoint for adding new worksheet using exceljs
app.post('/api/add-sheet', async (req, res) => {
  try {
    const { sheetName, customHeaders } = req.body;
    
    if (!sheetName || sheetName.trim() === '') {
      return res.status(400).json({
        success: false,
        error: 'Sheet name is required'
      });
    }

    // Sanitize sheet name
    const sanitizedSheetName = sheetName.trim().replace(/[\\/*?:[\]]/g, '_').substring(0, 31);
    
    if (sanitizedSheetName !== sheetName.trim()) {
      return res.status(400).json({
        success: false,
        error: 'Sheet name contains invalid characters or is too long (max 31 chars)'
      });
    }

    const workbook = new ExcelJS.Workbook();
    
    // Check if Excel file exists and load it
    if (fs.existsSync(EXCEL_FILE)) {
      await workbook.xlsx.readFile(EXCEL_FILE);
    } else {
      // Create new workbook with default sheet if file doesn't exist
      workbook.addWorksheet('Expenses');
      await workbook.xlsx.writeFile(EXCEL_FILE);
    }

    // Check if sheet already exists
    if (workbook.worksheets.some(ws => ws.name === sanitizedSheetName)) {
      return res.status(400).json({
        success: false,
        error: `Sheet "${sanitizedSheetName}" already exists`
      });
    }

    // Add new worksheet
    const newWorksheet = workbook.addWorksheet(sanitizedSheetName);
    
    // Determine headers to use
    let headers = [];
    if (customHeaders && Array.isArray(customHeaders) && customHeaders.length > 0) {
      // Use custom headers provided by user
      headers = customHeaders.map(header => ({
        header: header,
        key: header.toLowerCase().replace(/[^a-z0-9]/g, '_'),
        width: 15
      }));
    } else {
      // Use default expense headers
      headers = [
        { header: 'Sr No', key: 'srNo', width: 10 },
        { header: 'Date', key: 'date', width: 15 },
        { header: 'Given To', key: 'givenTo', width: 20 },
        { header: 'Amount', key: 'amount', width: 15 },
        { header: 'Mode', key: 'mode', width: 15 },
        { header: 'Description', key: 'description', width: 30 },
        { header: 'Fund', key: 'fund', width: 15 },
        { header: 'Status', key: 'status', width: 15 }
      ];
    }

    // Add headers to the new sheet
    newWorksheet.columns = headers;

    // Save the workbook
    await workbook.xlsx.writeFile(EXCEL_FILE);
    
    res.json({
      success: true,
      message: `Sheet "${sanitizedSheetName}" created successfully`,
      sheet: {
        name: sanitizedSheetName,
        createdAt: new Date().toISOString(),
        totalSheets: workbook.worksheets.length,
        headers: headers.map(h => h.header),
        hasCustomHeaders: customHeaders && customHeaders.length > 0
      }
    });

  } catch (error) {
    console.error('Error creating sheet:', error);
    res.status(500).json({
      success: false,
      error: 'Failed to create sheet: ' + error.message
    });
  }
});

// POST /delete-sheet - Delete a sheet (frontend endpoint)
app.post('/delete-sheet', async (req, res) => {
  try {
    const { sheetName } = req.body;
    
    if (!sheetName || sheetName.trim() === '') {
      return res.status(400).json({
        success: false,
        error: 'Sheet name is required'
      });
    }
    
    if (sheetName === 'Expenses') {
      return res.status(400).json({
        success: false,
        error: 'Cannot delete default Expenses sheet'
      });
    }

    // Load the workbook using ExcelJS to preserve all sheets
    const workbook = new ExcelJS.Workbook();
    
    if (fs.existsSync(EXCEL_FILE)) {
      await workbook.xlsx.readFile(EXCEL_FILE);
    } else {
      return res.status(404).json({
        success: false,
        error: 'Excel file not found'
      });
    }

    // Check if sheet exists
    const worksheet = workbook.getWorksheet(sheetName);
    if (!worksheet) {
      return res.status(404).json({
        success: false,
        error: `Sheet "${sheetName}" not found`
      });
    }

    // Remove the worksheet
    workbook.removeWorksheet(worksheet.id);

    // Save the workbook with error handling for file locks
    try {
      await workbook.xlsx.writeFile(EXCEL_FILE);
    } catch (writeError) {
      console.error('Error writing workbook:', writeError);
      
      // Handle file lock scenarios
      if (writeError.message.includes('EBUSY') || writeError.message.includes('locked') || writeError.message.includes('in use')) {
        return res.status(423).json({
          success: false,
          error: 'File is currently in use. Please close the Excel file and try again.'
        });
      }
      
      return res.status(500).json({
        success: false,
        error: 'Failed to save workbook: ' + writeError.message
      });
    }
    
    // Get updated list of sheets
    const remainingSheets = await getWorkbookSheets();
    
    res.json({
      success: true,
      message: 'Sheet deleted successfully',
      sheetName: sheetName,
      remainingSheets: remainingSheets
    });
    
  } catch (error) {
    console.error('Error deleting sheet:', error);
    
    // Enhanced error handling for specific scenarios
    let errorMessage = error.message || 'Failed to delete sheet';
    
    if (errorMessage.includes('EBUSY') || errorMessage.includes('locked') || errorMessage.includes('in use')) {
      errorMessage = 'File is currently in use. Please close the Excel file and try again.';
    } else if (errorMessage.includes('not found')) {
      errorMessage = 'Sheet not found. It may have already been deleted.';
    } else if (errorMessage.includes('permission') || errorMessage.includes('access')) {
      errorMessage = 'Permission denied. Please check file permissions and try again.';
    }
    
    res.status(500).json({
      success: false,
      error: errorMessage
    });
  }
});

// Health check endpoint
app.get('/health', async (req, res) => {
  res.json({ 
    status: 'Server is running', 
    timestamp: new Date().toISOString(),
    excelFile: fs.existsSync(EXCEL_FILE),
    jsonFile: fs.existsSync(JSON_FILE),
    totalRecords: currentExpenses.length,
    sheets: await getWorkbookSheets()
  });
});

// Start server
app.listen(PORT, () => {
  console.log(`Expense Management System server running on http://localhost:${PORT}`);
  console.log('Upload directory: uploads/');
  console.log('Static files served from: public/');
});
