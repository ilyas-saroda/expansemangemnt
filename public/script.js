// Simple and Optimized JavaScript for Fast Performance

// Global variables - minimized
let expenseData = [];
let filteredData = [];
let currentVoucherData = null;
let currentVoucherIndex = null;
let sortOrder = {}; // Track sort order for each column
let selectedVouchers = new Set(); // Track selected vouchers for bulk operations
let currentSheet = 'Expenses'; // Track currently selected sheet
let availableSheets = []; // Track available sheets

// DOM elements cache
const elements = {
    importBtn: document.getElementById('importBtn'),
    addExpenseBtn: document.getElementById('addExpenseBtn'),
    addSheetBtn: document.getElementById('addSheetBtn'),
    exportBtn: document.getElementById('exportBtn'),
    fileInput: document.getElementById('fileInput'),
    tableBody: document.getElementById('tableBody'),
    emptyState: document.getElementById('emptyState'),
    loadingOverlay: document.getElementById('loadingOverlay'),
    toast: document.getElementById('toast'),
    toastMessage: document.getElementById('toastMessage'),
    clearBtn: document.getElementById('clearBtn'),
    voucherModal: document.getElementById('voucherModal'),
    addExpenseModal: document.getElementById('addExpenseModal'),
    addSheetModal: document.getElementById('addSheetModal'),
    closeVoucherBtn: document.getElementById('closeVoucherBtn'),
    closeModal: document.getElementById('closeModal'),
    closeAddModal: document.getElementById('closeAddModal'),
    closeAddSheetModal: document.getElementById('closeAddSheetModal'),
    searchInput: document.getElementById('searchInput'),
    statusFilter: document.getElementById('statusFilter'),
    modeFilter: document.getElementById('modeFilter'),
    givenToFilter: document.getElementById('givenToFilter'),
    dateFilter: document.getElementById('dateFilter'),
    clearFilters: document.getElementById('clearFilters'),
    refreshBtn: document.getElementById('refreshBtn'),
        bulkVoucherBtn: document.getElementById('bulkVoucherBtn'),
    bulkVoucherModal: document.getElementById('bulkVoucherModal'),
    closeBulkModal: document.getElementById('closeBulkModal'),
    closeBulkVoucherBtn: document.getElementById('closeBulkVoucherBtn'),
    selectAllCheckbox: document.getElementById('selectAllCheckbox'),
    sheetNameInput: document.getElementById('sheetNameInput'),
    createSheetBtn: document.getElementById('createSheetBtn'),
    cancelCreateSheetBtn: document.getElementById('cancelCreateSheetBtn'),
    sheetSelect: document.getElementById('sheetSelect')
};

// Initialize event listeners - simplified
function initializeEventListeners() {
    // Main buttons
    elements.addExpenseBtn?.addEventListener('click', openAddExpenseModal);
    elements.addSheetBtn?.addEventListener('click', openAddSheetModal);
    elements.importBtn?.addEventListener('click', () => elements.fileInput?.click());
    elements.clearBtn?.addEventListener('click', clearAllData);
    elements.refreshBtn?.addEventListener('click', refreshData);
    
    // Sheet selection
    elements.sheetSelect?.addEventListener('change', handleSheetChange);
        elements.closeVoucherBtn?.addEventListener('click', () => {
        elements.voucherModal?.classList.remove('active');
    });
    elements.closeModal?.addEventListener('click', () => {
        elements.voucherModal?.classList.remove('active');
    });
    elements.closeAddModal?.addEventListener('click', () => {
        elements.addExpenseModal?.classList.remove('active');
    });
    elements.closeAddSheetModal?.addEventListener('click', () => {
        elements.addSheetModal?.classList.remove('active');
    });
    elements.cancelCreateSheetBtn?.addEventListener('click', () => {
        elements.addSheetModal?.classList.remove('active');
    });
    elements.createSheetBtn?.addEventListener('click', createNewSheet);
        elements.bulkVoucherBtn?.addEventListener('click', showBulkVoucherModal);
    
    // Bulk voucher modal listeners
    elements.closeBulkModal?.addEventListener('click', () => {
        elements.bulkVoucherModal?.classList.remove('active');
    });
    elements.closeBulkVoucherBtn?.addEventListener('click', () => {
        elements.bulkVoucherModal?.classList.remove('active');
    });
    
    // File input
    elements.fileInput?.addEventListener('change', handleFileUpload);
    
    // Search and filters
    elements.searchInput?.addEventListener('input', applyFilters);
    elements.statusFilter?.addEventListener('change', applyFilters);
    elements.modeFilter?.addEventListener('change', applyFilters);
    elements.givenToFilter?.addEventListener('change', applyFilters);
    elements.dateFilter?.addEventListener('change', applyFilters);
    elements.clearFilters?.addEventListener('click', clearAllFilters);
    
    // Export dropdown - simplified
    elements.exportBtn?.addEventListener('click', (e) => {
        e.stopPropagation();
        const dropdown = elements.exportBtn.closest('.dropdown');
        dropdown?.classList.toggle('active');
    });
    
    // Close dropdowns on outside click
    document.addEventListener('click', () => {
        document.querySelectorAll('.dropdown').forEach(d => d.classList.remove('active'));
    });
    
    // Modal close events
    document.addEventListener('keydown', (e) => {
        if (e.key === 'Escape') {
            closeAllModals();
        }
    });
}

// Initialize app
document.addEventListener('DOMContentLoaded', function() {
    initializeEventListeners();
    loadExistingData();
    updateStatistics();
});



// Save data to current sheet
async function saveDataToCurrentSheet() {
    try {
        const response = await fetch('/update', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                sheetName: currentSheet,
                updatedData: expenseData
            })
        });
        
        const result = await response.json();
        
        if (result.success) {
            console.log(`Saved ${expenseData.length} records to sheet "${currentSheet}"`);
        } else {
            console.error(`Failed to save to sheet "${currentSheet}":`, result.error);
        }
    } catch (error) {
        console.error('Error saving data:', error);
    }
}

// Utility functions - simplified
function showLoading(message = 'Processing...') {
    if (elements.loadingOverlay) {
        elements.loadingOverlay.classList.add('active');
        const loadingMessage = document.getElementById('loadingMessage');
        if (loadingMessage) loadingMessage.textContent = message;
    }
}

function hideLoading() {
    elements.loadingOverlay?.classList.remove('active');
}

function showToast(message, type = 'success') {
    if (elements.toast && elements.toastMessage) {
        elements.toastMessage.textContent = message;
        elements.toast.className = 'toast show';
        if (type !== 'success') elements.toast.classList.add(type);
        
        setTimeout(() => {
            elements.toast.classList.remove('show');
        }, 3000);
    }
}

function closeAllModals() {
    elements.voucherModal?.classList.remove('active');
    elements.addExpenseModal?.classList.remove('active');
    elements.addSheetModal?.classList.remove('active');
    document.querySelectorAll('.dropdown').forEach(d => d.classList.remove('active'));
}

// Simple currency formatting
function formatCurrency(amount) {
    return '₹' + parseFloat(amount || 0).toLocaleString('en-IN', {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2
    });
}

// Enhanced number to words for Indian system
function numberToWords(num) {
    if (num === 0) return 'Zero Rupees';
    
    const ones = ['', 'One', 'Two', 'Three', 'Four', 'Five', 'Six', 'Seven', 'Eight', 'Nine'];
    const teens = ['Ten', 'Eleven', 'Twelve', 'Thirteen', 'Fourteen', 'Fifteen', 'Sixteen', 'Seventeen', 'Eighteen', 'Nineteen'];
    const tens = ['', '', 'Twenty', 'Thirty', 'Forty', 'Fifty', 'Sixty', 'Seventy', 'Eighty', 'Ninety'];
    
    // Convert to integer and handle decimal part
    const integerPart = Math.floor(num);
    const decimalPart = Math.round((num - integerPart) * 100);
    
    let words = '';
    
    // Convert integer part to words
    if (integerPart > 0) {
        words = convertIntegerToWords(integerPart) + ' Rupees';
    }
    
    // Add decimal part if exists
    if (decimalPart > 0) {
        if (words) words += ' and ';
        words += convertIntegerToWords(decimalPart) + ' Paise';
    }
    
    return words || 'Zero Rupees';
}

function convertIntegerToWords(num) {
    if (num === 0) return '';
    
    const ones = ['', 'One', 'Two', 'Three', 'Four', 'Five', 'Six', 'Seven', 'Eight', 'Nine'];
    const teens = ['Ten', 'Eleven', 'Twelve', 'Thirteen', 'Fourteen', 'Fifteen', 'Sixteen', 'Seventeen', 'Eighteen', 'Nineteen'];
    const tens = ['', '', 'Twenty', 'Thirty', 'Forty', 'Fifty', 'Sixty', 'Seventy', 'Eighty', 'Ninety'];
    
    let words = '';
    
    if (num < 10) {
        words = ones[num];
    } else if (num < 20) {
        words = teens[num - 10];
    } else if (num < 100) {
        words = tens[Math.floor(num / 10)] + ' ' + ones[num % 10];
    } else if (num < 1000) {
        words = ones[Math.floor(num / 100)] + ' Hundred ' + convertIntegerToWords(num % 100);
    } else if (num < 100000) {
        words = convertIntegerToWords(Math.floor(num / 1000)) + ' Thousand ' + convertIntegerToWords(num % 1000);
    } else if (num < 10000000) {
        words = convertIntegerToWords(Math.floor(num / 100000)) + ' Lakh ' + convertIntegerToWords(num % 100000);
    } else {
        words = convertIntegerToWords(Math.floor(num / 10000000)) + ' Crore ' + convertIntegerToWords(num % 10000000);
    }
    
    return words.trim();
}

// Refresh data function
function refreshData() {
    showLoading('Refreshing data...');
    loadExistingData();
}

// Table sorting function
function sortTable(column) {
    const dataToSort = filteredData.length > 0 ? filteredData : expenseData;
    
    // Toggle sort order
    if (!sortOrder[column]) {
        sortOrder[column] = 'asc';
    } else if (sortOrder[column] === 'asc') {
        sortOrder[column] = 'desc';
    } else {
        sortOrder[column] = 'asc';
    }
    
    // Update sort icons
    updateSortIcons(column);
    
    // Sort the data
    dataToSort.sort((a, b) => {
        let valueA = a[column] || '';
        let valueB = b[column] || '';
        
        // Handle different data types
        if (column === 'amount' || column === 'srNo') {
            valueA = parseFloat(valueA) || 0;
            valueB = parseFloat(valueB) || 0;
        } else if (column === 'date') {
            valueA = new Date(valueA) || new Date(0);
            valueB = new Date(valueB) || new Date(0);
        } else {
            valueA = valueA.toString().toLowerCase();
            valueB = valueB.toString().toLowerCase();
        }
        
        if (sortOrder[column] === 'asc') {
            return valueA > valueB ? 1 : valueA < valueB ? -1 : 0;
        } else {
            return valueA < valueB ? 1 : valueA > valueB ? -1 : 0;
        }
    });
    
    // Re-render table
    renderTable();
}

// Update sort icons
function updateSortIcons(activeColumn) {
    const headers = document.querySelectorAll('th');
    headers.forEach(header => {
        const icon = header.querySelector('i');
        if (icon) {
            icon.className = 'fas fa-sort';
            icon.style.color = '#666';
        }
    });
    
    // Update active column icon
    const activeHeader = document.querySelector(`th[onclick*="${activeColumn}"]`);
    if (activeHeader) {
        const icon = activeHeader.querySelector('i');
        if (icon) {
            icon.className = sortOrder[activeColumn] === 'asc' ? 'fas fa-sort-up' : 'fas fa-sort-down';
            icon.style.color = '#2196F3';
        }
    }
}

// Load existing data from server
async function loadExistingData() {
    try {
        showLoading('Loading existing data...');
        
        const response = await fetch(`/data?sheetName=${encodeURIComponent(currentSheet)}`);
        const result = await response.json();
        
        if (result.success) {
            expenseData = result.data || [];
            
            // Update available sheets
            if (result.availableSheets) {
                availableSheets = result.availableSheets;
                updateSheetSelector();
            }
            
            // Update current sheet if server returned a different one
            if (result.currentSheet && result.currentSheet !== currentSheet) {
                currentSheet = result.currentSheet;
            }
            
            renderTable();
            updateStatistics();
            populateGivenToFilter();
            populateModeFilter();
            
            showToast(`Loaded ${result.data.length} records from "${currentSheet}"`, 'success');
        } else {
            showToast('No existing data found', 'info');
        }
    } catch (error) {
        console.error('Error loading existing data:', error);
        showToast('Failed to load existing data', 'error');
    } finally {
        hideLoading();
    }
}

// Handle sheet selection change
function handleSheetChange() {
    const selectedSheet = elements.sheetSelect.value;
    if (selectedSheet && selectedSheet !== currentSheet) {
        currentSheet = selectedSheet;
        loadExistingData(); // Reload data for the selected sheet
    }
}

// Update sheet selector dropdown
function updateSheetSelector() {
    if (!elements.sheetSelect || !availableSheets.length) return;
    
    // Clear existing options
    elements.sheetSelect.innerHTML = '';
    
    // Add available sheets to dropdown
    availableSheets.forEach(sheet => {
        const option = document.createElement('option');
        option.value = sheet.name;
        option.textContent = sheet.name + (sheet.isDefault ? ' (Default)' : '');
        option.selected = sheet.name === currentSheet;
        elements.sheetSelect.appendChild(option);
    });
}

// Populate Given To filter dropdown
function populateGivenToFilter() {
    const uniqueGivenTo = getUniqueValues('givenTo');
    
    // Clear existing options except the first one
    if (elements.givenToFilter) {
        elements.givenToFilter.innerHTML = '<option value="">All Persons</option>';
        
        // Add unique names to dropdown
        uniqueGivenTo.forEach(name => {
            const option = document.createElement('option');
            option.value = name;
            option.textContent = name;
            elements.givenToFilter.appendChild(option);
        });
    }
}

// Populate mode filter dropdown dynamically
function populateModeFilter() {
    const uniqueModes = getUniqueValues('mode');
    
    // Clear existing options except the first one
    if (elements.modeFilter) {
        elements.modeFilter.innerHTML = '<option value="">All Modes</option>';
        
        // Add unique modes to dropdown
        uniqueModes.forEach(mode => {
            const option = document.createElement('option');
            option.value = mode;
            option.textContent = mode;
            elements.modeFilter.appendChild(option);
        });
    }
}

// Get unique values from expense data
function getUniqueValues(field) {
    const values = new Set();
    expenseData.forEach(expense => {
        if (expense[field] && expense[field].trim() !== '') {
            values.add(expense[field].trim());
        }
    });
    return Array.from(values).sort();
}

// Simplified table rendering
function renderTable() {
    const dataToRender = filteredData.length > 0 ? filteredData : expenseData;
    
    if (!dataToRender || dataToRender.length === 0) {
        showEmptyState();
        return;
    }

    hideEmptyState();
    elements.tableBody.innerHTML = '';

    dataToRender.forEach((expense, index) => {
        const originalIndex = expenseData.findIndex(exp => exp.srNo === expense.srNo);
        const row = createTableRow(expense, originalIndex);
        elements.tableBody.appendChild(row);
    });
    
    // Update filters after rendering
    populateModeFilter();
    populateGivenToFilter();
}

// Simplified table row creation
function createTableRow(expense, index) {
    const amount = parseFloat(expense.amount || 0);
    const row = document.createElement('tr');
    row.innerHTML = `
        <td>${expense.srNo || index + 1}</td>
        <td contenteditable="true" data-field="date" data-index="${index}">${formatDate(expense.date)}</td>
        <td contenteditable="true" data-field="givenTo" data-index="${index}">${expense.givenTo || '-'}</td>
        <td contenteditable="true" data-field="amount" data-index="${index}">${formatCurrency(amount)}</td>
        <td>${numberToWords(amount)}</td>
        <td contenteditable="true" data-field="mode" data-index="${index}">${expense.mode || '-'}</td>
        <td contenteditable="true" data-field="description" data-index="${index}">${expense.description || '-'}</td>
        <td contenteditable="true" data-field="fund" data-index="${index}">${expense.fund || '-'}</td>
        <td>${createEditableStatusBadge(expense.status, index)}</td>
        <td>${createActionButtons(expense, index)}</td>
    `;
    
    addInlineEditListeners(row);
    return row;
}

// Helper functions - simplified
function showEmptyState() {
    elements.emptyState.style.display = 'block';
    elements.tableBody.parentElement.style.display = 'none';
}

function hideEmptyState() {
    elements.emptyState.style.display = 'none';
    elements.tableBody.parentElement.style.display = 'table';
}

function formatDate(dateString) {
    if (!dateString) return '-';
    try {
        const date = new Date(dateString);
        if (isNaN(date.getTime())) return dateString;
        return date.toLocaleDateString('en-US', {
            year: 'numeric',
            month: 'short',
            day: 'numeric'
        });
    } catch (error) {
        return dateString;
    }
}

// Simplified statistics
function updateStatistics() {
    const totalAmount = expenseData.reduce((sum, expense) => sum + parseFloat(expense.amount || 0), 0);
    const paidCount = expenseData.filter(expense => expense.status === 'Paid').length;
    const pendingCount = expenseData.filter(expense => expense.status === 'Pending').length;
    const totalCount = expenseData.length;

    const totalAmountEl = document.getElementById('totalAmount');
    const totalAmountInWordsEl = document.getElementById('totalAmountInWords');
    const paidCountEl = document.getElementById('paidCount');
    const pendingCountEl = document.getElementById('pendingCount');
    const totalCountEl = document.getElementById('totalCount');

    if (totalAmountEl) totalAmountEl.textContent = formatCurrency(totalAmount);
    if (totalAmountInWordsEl) totalAmountInWordsEl.textContent = numberToWords(totalAmount);
    if (paidCountEl) paidCountEl.textContent = paidCount;
    if (pendingCountEl) pendingCountEl.textContent = pendingCount;
    if (totalCountEl) totalCountEl.textContent = totalCount;
}

// Simplified filter functions
function applyFilters() {
    const searchTerm = elements.searchInput?.value.toLowerCase().trim() || '';
    const statusValue = elements.statusFilter?.value || '';
    const modeValue = elements.modeFilter?.value || '';
    const givenToValue = elements.givenToFilter?.value || '';
    const dateValue = elements.dateFilter?.value || '';
    
    filteredData = expenseData.filter(expense => {
        const matchesSearch = !searchTerm || 
            expense.givenTo?.toLowerCase().includes(searchTerm) ||
            expense.description?.toLowerCase().includes(searchTerm) ||
            expense.fund?.toLowerCase().includes(searchTerm) ||
            expense.mode?.toLowerCase().includes(searchTerm);
        
        const matchesStatus = !statusValue || expense.status === statusValue;
        const matchesMode = !modeValue || expense.mode === modeValue;
        const matchesGivenTo = !givenToValue || expense.givenTo === givenToValue;
        
        let matchesDate = !dateValue;
        if (dateValue && expense.date) {
            const expenseDate = new Date(expense.date).toISOString().split('T')[0];
            matchesDate = expenseDate === dateValue;
        }
        
        return matchesSearch && matchesStatus && matchesMode && matchesGivenTo && matchesDate;
    });
    
    renderTable();
    updateFilteredStatistics();
}

function clearAllFilters() {
    elements.searchInput.value = '';
    elements.statusFilter.value = '';
    elements.modeFilter.value = '';
    elements.givenToFilter.value = '';
    elements.dateFilter.value = '';
    filteredData = [];
    renderTable();
    updateStatistics();
}

function updateFilteredStatistics() {
    const dataToUse = filteredData.length > 0 ? filteredData : expenseData;
    const totalAmount = dataToUse.reduce((sum, expense) => sum + parseFloat(expense.amount || 0), 0);
    const paidCount = dataToUse.filter(expense => expense.status === 'Paid').length;
    const pendingCount = dataToUse.filter(expense => expense.status === 'Pending').length;
    const totalCount = dataToUse.length;

    const totalAmountEl = document.getElementById('totalAmount');
    const totalAmountInWordsEl = document.getElementById('totalAmountInWords');
    const paidCountEl = document.getElementById('paidCount');
    const pendingCountEl = document.getElementById('pendingCount');
    const totalCountEl = document.getElementById('totalCount');

    if (totalAmountEl) totalAmountEl.textContent = formatCurrency(totalAmount);
    if (totalAmountInWordsEl) totalAmountInWordsEl.textContent = numberToWords(totalAmount);
    if (paidCountEl) paidCountEl.textContent = paidCount;
    if (pendingCountEl) pendingCountEl.textContent = pendingCount;
    if (totalCountEl) totalCountEl.textContent = totalCount;
}


// Handle file upload
async function handleFileUpload(event) {
    const file = event.target.files[0];
    
    if (!file) {
        showToast('No file selected', 'warning');
        return;
    }

    // Validate file type
    const validTypes = [
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'application/vnd.ms-excel'
    ];
    
    if (!validTypes.includes(file.type)) {
        showToast('Invalid file type. Please upload an Excel file (.xlsx or .xls)', 'error');
        event.target.value = '';
        return;
    }

    // Show loading overlay
    showLoading();

    // Create FormData
    const formData = new FormData();
    formData.append('excelFile', file);

    try {
        // Send file to server
        const response = await fetch('/upload-excel', {
            method: 'POST',
            body: formData
        });

        const result = await response.json();

        if (response.ok && result.success) {
            // Success - update table with data
            expenseData = result.data;
            renderTable();
            updateStatistics();
            showToast(result.message || 'Excel file imported successfully!', 'success');
        } else {
            // Error from server
            showToast(result.error || 'Error processing file', 'error');
        }
    } catch (error) {
        console.error('Upload error:', error);
        showToast('Network error. Please try again.', 'error');
    } finally {
        hideLoading();
        // Reset file input
        event.target.value = '';
    }
}

// Render table with expense data
function renderTable() {
    const dataToRender = filteredData.length > 0 ? filteredData : expenseData;
    
    if (!dataToRender || dataToRender.length === 0) {
        showEmptyState();
        return;
    }

    hideEmptyState();
    
    // Clear existing table content
    tableBody.innerHTML = '';

    // Add rows for each expense
    dataToRender.forEach((expense, index) => {
        const originalIndex = expenseData.findIndex(exp => exp.srNo === expense.srNo);
        const row = createTableRow(expense, originalIndex);
        tableBody.appendChild(row);
    });
}

// Create table row element
function createTableRow(expense, index) {
    const amount = parseFloat(expense.amount || 0);
    const amountInWords = numberToWords(amount);
    
    const row = document.createElement('tr');
    row.innerHTML = `
        <td>
            <input type="checkbox" class="voucher-checkbox" data-srno="${expense.srNo}" onchange="toggleVoucherSelection(${expense.srNo})">
        </td>
        <td>${expense.srNo || index + 1}</td>
        <td contenteditable="true" data-field="date" data-index="${index}">${formatDate(expense.date)}</td>
        <td contenteditable="true" data-field="givenTo" data-index="${index}">${expense.givenTo || '-'}</td>
        <td contenteditable="true" data-field="amount" data-index="${index}">${formatCurrency(amount)}</td>
        <td>${amountInWords}</td>
        <td contenteditable="true" data-field="mode" data-index="${index}">${expense.mode || '-'}</td>
        <td contenteditable="true" data-field="description" data-index="${index}">${expense.description || '-'}</td>
        <td contenteditable="true" data-field="fund" data-index="${index}">${expense.fund || '-'}</td>
        <td>${createEditableStatusBadge(expense.status, index)}</td>
        <td>${createActionButtons(expense, index)}</td>
    `;
    
    // Add event listeners for inline editing
    addInlineEditListeners(row);
    
    return row;
}

// Create editable status badge
function createEditableStatusBadge(status, index) {
    const statusClass = status === 'Paid' ? 'status-paid' : 'status-pending';
    return `
        <select class="status-badge ${statusClass}" data-field="status" data-index="${index}" onchange="handleStatusChange(this, ${index})">
            <option value="Pending" ${status === 'Pending' ? 'selected' : ''}>Pending</option>
            <option value="Paid" ${status === 'Paid' ? 'selected' : ''}>Paid</option>
        </select>
    `;
}

// Add inline edit listeners
function addInlineEditListeners(row) {
    const editableCells = row.querySelectorAll('[contenteditable="true"]');
    
    editableCells.forEach(cell => {
        cell.addEventListener('blur', function() {
            handleInlineEdit(this);
        });
        
        cell.addEventListener('keypress', function(e) {
            if (e.key === 'Enter') {
                e.preventDefault();
                this.blur();
            }
        });
        
        // Add visual feedback
        cell.addEventListener('focus', function() {
            this.style.backgroundColor = 'rgba(255, 255, 255, 0.1)';
            this.style.outline = '2px solid #4CAF50';
        });
        
        cell.addEventListener('blur', function() {
            this.style.backgroundColor = '';
            this.style.outline = '';
        });
    });
}

// Handle inline edit
function handleInlineEdit(cell) {
    const index = parseInt(cell.dataset.index);
    const field = cell.dataset.field;
    let newValue = cell.textContent.trim();
    
    // Special handling for amount field
    if (field === 'amount') {
        newValue = newValue.replace('₹', '').replace(/,/g, '').trim();
        if (isNaN(newValue) || parseFloat(newValue) <= 0) {
            showToast('Please enter a valid amount', 'error');
            // Restore original value
            cell.textContent = formatCurrency(parseFloat(expenseData[index].amount || 0));
            return;
        }
        // Update the amount in words column in the same row
        const row = cell.parentElement;
        const amountInWordsCell = row.cells[4]; // Amount in Words column
        amountInWordsCell.textContent = numberToWords(parseFloat(newValue));
    }
    
    // Update data
    expenseData[index][field] = newValue;
    
  
    // Auto-save to Excel
    saveExpense(index);
}

// Handle status change
function handleStatusChange(select, index) {
    const newStatus = select.value;
    expenseData[index].status = newStatus;
    
    // Update badge styling
    select.className = `status-badge ${newStatus === 'Paid' ? 'status-paid' : 'status-pending'}`;
    
    // Auto-save to Excel
    saveExpense(index);
}

// Create status badge HTML
function createStatusBadge(status) {
    const statusClass = status === 'Paid' ? 'status-paid' : 'status-pending';
    return `<span class="status-badge ${statusClass}">${status || 'Pending'}</span>`;
}

// Create action buttons HTML
function createActionButtons(expense, index) {
    return `
        <div class="action-buttons">
            <button class="btn btn-success btn-sm" onclick="saveExpense(${index})">
                <i class="fas fa-save"></i> Save
            </button>
            <button class="btn btn-info btn-sm" onclick="generateVoucher(${index})">
                <i class="fas fa-receipt"></i> Voucher
            </button>
            <button class="btn btn-danger btn-sm" onclick="deleteExpense(${index})">
                <i class="fas fa-trash"></i> Delete
            </button>
        </div>
    `;
}

// Format date for display
function formatDate(dateString) {
    if (!dateString) return '-';
    
    try {
        const date = new Date(dateString);
        if (isNaN(date.getTime())) return dateString;
        
        return date.toLocaleDateString('en-US', {
            year: 'numeric',
            month: 'short',
            day: 'numeric'
        });
    } catch (error) {
        return dateString;
    }
}

// Save expense and sync with Excel
async function saveExpense(index) {
    try {
        const expense = expenseData[index];
        
        showLoading('Saving expense and updating Excel...');
        
        // Call server to update data
        const response = await fetch('/update', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                sheetName: currentSheet,
                updatedData: expenseData
            })
        });
        
        const result = await response.json();
        
        if (result.success) {
            showToast('Expense saved and Excel file updated!', 'success');
            console.log('Saved expense:', expense);
            
            // Smart Synchronization: Update bulk voucher data automatically
            synchronizeBulkVoucherData(expense, index);
        } else {
            showToast('Failed to save expense', 'error');
            console.error('Save error:', result.error);
        }
    } catch (error) {
        console.error('Error saving expense:', error);
        showToast('Error saving expense', 'error');
    } finally {
        hideLoading();
    }
}

// Generate premium bank-standard voucher with Tailwind CSS
function generateVoucher(index) {
    const expense = expenseData[index];
    currentVoucherData = expense;
    currentVoucherIndex = index;
    
    // Create premium bank-standard voucher HTML
    const voucherHTML = `
        <div class="print-area">
            <div class="bg-white relative overflow-hidden shadow-2xl border-4 border-black" id="premiumVoucher">
                <!-- Watermark Background -->
                <div class="absolute inset-0 opacity-5 pointer-events-none">
                    <div class="text-9xl font-bold text-blue-900 transform rotate-45 absolute top-1/2 left-1/2 -translate-x-1/2 -translate-y-1/2">
                        ASJ SAGWARA
                    </div>
                </div>
                
                <!-- Header Section -->
                <div class="relative z-10 text-center py-6 border-b-4 border-black">
                    <h1 class="font-playfair text-3xl font-bold text-gray-900 mb-2">
                        Anjuman E Saifee Jamaat, Sagwara
                    </h1>
                    <p class="text-3xl text-black mb-4">Shahrullah il Moazzam - 1447</p>
                    
                    <div class="relative inline-block">
                        <h2 class="font-playfair text-lg font-bold text-gray-900">EXPENSE VOUCHER</h2>
                        <!-- Double Underline -->
                        <div class="absolute bottom-0 left-0 right-0 h-1 bg-black"></div>
                        <div class="absolute -bottom-2 left-0 right-0 h-0.5 bg-black"></div>
                    </div>
                </div>
                
                <!-- Main Content -->
                <div class="relative z-10 p-8">
                    <!-- Grid Layout -->
                    <div class="grid grid-cols-4 gap-0 border-2 border-black">
                        <!-- S.No -->
                        <div class="border-r-2 border-b-2 border-black p-3 bg-gray-50">
                            <div class="font-bold text-sm text-gray-800">S.No</div>
                        </div>
                        <div class="border-r-2 border-b-2 border-black p-3">
                            <div class="text-center font-semibold" data-field="srNo">${expense.srNo || index + 1}</div>
                        </div>
                        
                        <!-- Date -->
                        <div class="border-r-2 border-b-2 border-black p-3 bg-gray-50">
                            <div class="font-bold text-sm text-gray-800">Date</div>
                        </div>
                        <div class="border-b-2 border-black p-3">
                            <div class="text-center font-semibold" data-field="date">${formatDate(expense.date)}</div>
                        </div>
                        
                        <!-- Name -->
                        <div class="border-r-2 border-b-2 border-black p-3 bg-gray-50">
                            <div class="font-bold text-sm text-gray-800">Name</div>
                        </div>
                        <div class="border-r-2 border-b-2 border-black p-3">
                            <div class="text-center font-semibold" data-field="givenTo">${expense.givenTo || '-'}</div>
                        </div>
                        
                        <!-- Mode of Payment -->
                        <div class="border-r-2 border-b-2 border-black p-3 bg-gray-50">
                            <div class="font-bold text-sm text-gray-800">Mode of Payment</div>
                        </div>
                        <div class="border-b-2 border-black p-3">
                            <div class="text-center font-semibold" data-field="mode">${expense.mode || '-'}</div>
                        </div>
                    </div>
                    
                    <!-- Expanded Description Section (Extra Large) -->
                    <div class="mt-6 border-2 border-black">
                        <div class="bg-gray-50 p-3 border-b-2 border-black flex justify-between items-center">
                            <div class="font-bold text-sm text-gray-800">Description</div>
                            <div class="flex gap-2">
                                <button class="btn btn-sm btn-primary" id="editDescBtn" onclick="toggleDescriptionEdit()">
                                    <i class="fas fa-edit"></i> Edit
                                </button>
                                <button class="btn btn-sm btn-success" id="saveDescBtn" onclick="saveDescriptionEdit()" style="display: none;">
                                    <i class="fas fa-save"></i> Save
                                </button>
                                <button class="btn btn-sm btn-secondary" id="cancelDescBtn" onclick="cancelDescriptionEdit()" style="display: none;">
                                    <i class="fas fa-times"></i> Cancel
                                </button>
                            </div>
                        </div>
                        <div class="p-8 min-h-[400px]">
                            <div id="descriptionDisplay" class="text-gray-700 text-base leading-relaxed" data-field="description">${expense.description || 'No description provided'}</div>
                            <textarea id="descriptionTextarea" class="w-full h-full p-2 border-2 border-blue-500 rounded text-gray-700 text-base leading-relaxed resize-none" style="display: none; min-height: 350px;" placeholder="Enter description here...">${expense.description || ''}</textarea>
                        </div>
                    </div>
                    
                    <!-- Total Amount Section (Direct - No separate Amount field) -->
                    <div class="mt-6 border-4 border-double border-black">
                        <div class="grid grid-cols-2 gap-0">
                            <div class="bg-gray-100 p-4 border-r-4 border-double border-black">
                                <div class="font-bold text-xl text-gray-900">TOTAL AMOUNT</div>
                            </div>
                            <div class="p-4 text-right bg-yellow-50">
                                <div class="text-3xl font-bold text-gray-900" data-field="amount" id="totalAmountDisplay">${formatCurrency(parseFloat(expense.amount || 0))}</div>
                            </div>
                        </div>
                    </div>
                    
                    <!-- Footer Signatures (Moved Lower) -->
                    <div class="mt-16 grid grid-cols-2 gap-8">
                        <div class="border-t-4 border-black pt-12">
                            <div class="text-center">
                                <div class="font-bold text-gray-800 mb-4">Approved By</div>
                                <div class="h-12 border-b-2 border-gray-400"></div>
                            </div>
                        </div>
                        <div class="border-t-4 border-black pt-12">
                            <div class="text-center">
                                <div class="font-bold text-gray-800 mb-4">Receiver Signature</div>
                                <div class="h-12 border-b-2 border-gray-400"></div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    `;
    
    // Show modal with voucher
    document.getElementById('voucherContent').innerHTML = voucherHTML;
    voucherModal.classList.add('active');
    
    // Reset edit state
    isVoucherEditMode = false;
    document.getElementById('saveVoucherBtn').style.display = 'none';
    
    // Initialize auto-calculation
    initializeVoucherCalculations();
}


// Disable voucher edit mode for premium design
function disableVoucherEdit() {
    const premiumVoucher = document.getElementById('premiumVoucher');
    const saveBtn = document.getElementById('saveVoucherBtn');
    
    if (!premiumVoucher) return;
    
    premiumVoucher.classList.remove('edit-mode');
    saveBtn.style.display = 'none';
    
    // Remove edit styling from all fields
    const editableFields = premiumVoucher.querySelectorAll('[data-field]');
    editableFields.forEach(field => {
        field.classList.remove('bg-yellow-50', 'border-2', 'border-blue-500', 'rounded', 'px-2', 'py-1');
        field.contentEditable = false;
    });
    
    // Remove edit indicator
    const editIndicator = document.getElementById('editIndicator');
    if (editIndicator) {
        editIndicator.remove();
    }
    
    isVoucherEditMode = false;
}

// Save voucher changes for premium design
async function saveVoucherChanges() {
    if (!currentVoucherData || currentVoucherIndex === undefined) {
        showToast('No voucher data to save', 'error');
        return;
    }
    
    try {
        showLoading('Saving premium voucher changes...');
        
        // Get updated values from premium voucher
        const premiumVoucher = document.getElementById('premiumVoucher');
        if (!premiumVoucher) {
            showToast('Voucher not found', 'error');
            return;
        }
        
        const updatedData = { ...currentVoucherData };
        
        // Update fields from premium voucher
        premiumVoucher.querySelectorAll('[data-field]').forEach(field => {
            const fieldName = field.dataset.field;
            let value = field.textContent.trim();
            
            // Special handling for amount field
            if (fieldName === 'amount') {
                value = value.replace('₹', '').replace(/,/g, '').trim();
                updatedData[fieldName] = parseFloat(value).toFixed(2);
            } else if (fieldName === 'date') {
                // Keep date as is
                updatedData[fieldName] = currentVoucherData[fieldName];
            } else if (fieldName === 'srNo') {
                // Don't update serial number
                updatedData[fieldName] = currentVoucherData[fieldName];
            } else {
                updatedData[fieldName] = value;
            }
        });
        
        // Update expense data
        expenseData[currentVoucherIndex] = updatedData;
        
        // Save to server
        const response = await fetch('/update', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                sheetName: currentSheet,
                updatedData: expenseData
            })
        });
        
        const result = await response.json();
        
        if (result.success) {
            // Update current voucher data
            currentVoucherData = updatedData;
            
            // Disable edit mode
            disableVoucherEdit();
            
            // Refresh table
            renderTable();
            updateStatistics();
            
            // Smart Synchronization: Update bulk voucher data automatically
            synchronizeBulkVoucherData(updatedData, currentVoucherIndex);
            
            showToast('Premium voucher changes saved successfully!', 'success');
        } else {
            showToast('Failed to save changes: ' + result.error, 'error');
        }
    } catch (error) {
        console.error('Error saving voucher changes:', error);
        showToast('Error saving changes', 'error');
    } finally {
        hideLoading();
    }
}

// Export voucher as PDF
function exportVoucherAsPDF() {
    if (!currentVoucherData) {
        showToast('No voucher to export', 'error');
        return;
    }
    
    // Create printable voucher content
    const voucherContent = document.getElementById('voucherContent').innerHTML;
    
    // Open in new window for PDF printing
    const printWindow = window.open('', '_blank');
    printWindow.document.write(`
        <html>
            <head>
                <title>Expense Voucher - ${currentVoucherData.srNo}</title>
                <style>
                    body { font-family: Arial, sans-serif; margin: 20px; color: #333; }
                    .voucher-template { 
                        background: #ffffff; 
                        color: #333; 
                        border-radius: 10px; 
                        padding: 30px; 
                        font-family: 'Arial', sans-serif; 
                        box-shadow: 0 4px 20px rgba(0, 0, 0, 0.1); 
                        max-width: 800px;
                        margin: 0 auto;
                    }
                    .voucher-header { 
                        text-align: center; 
                        margin-bottom: 30px; 
                        padding-bottom: 20px; 
                        border-bottom: 2px solid #2a5298; 
                    }
                    .voucher-header h2 { 
                        color: #2a5298; 
                        font-size: 2rem; 
                        margin: 0; 
                        font-weight: bold; 
                    }
                    .voucher-header h3 { 
                        color: #2a5298; 
                        font-size: 1.5rem; 
                        margin: 10px 0; 
                        font-weight: bold; 
                    }
                    .voucher-info { 
                        display: grid; 
                        grid-template-columns: repeat(2, 1fr); 
                        gap: 20px; 
                        margin-bottom: 30px; 
                    }
                    .voucher-field { 
                        display: flex; 
                        align-items: center; 
                        padding: 10px 0; 
                        border-bottom: 1px solid #ddd; 
                    }
                    .voucher-field label { 
                        font-weight: bold; 
                        color: #2a5298; 
                        min-width: 120px; 
                        margin-right: 15px; 
                    }
                    .voucher-description { margin: 30px 0; }
                    .voucher-description h4 { 
                        color: #2a5298; 
                        margin-bottom: 15px; 
                        font-weight: bold; 
                    }
                    .voucher-description .description-text { 
                        background: #f8f9fa; 
                        padding: 15px; 
                        border-radius: 8px; 
                        border: 1px solid #ddd; 
                        min-height: 80px; 
                    }
                    .voucher-amount-section { 
                        text-align: right; 
                        margin: 30px 0; 
                        padding: 20px; 
                        background: #f8f9fa; 
                        border-radius: 8px; 
                        border: 2px solid #2a5298; 
                    }
                    .voucher-amount-section .amount-label { 
                        font-size: 1.2rem; 
                        color: #666; 
                        margin-bottom: 10px; 
                    }
                    .voucher-amount-section .amount-value { 
                        font-size: 2rem; 
                        font-weight: bold; 
                        color: #2a5298; 
                    }
                    .voucher-signatures { 
                        display: grid; 
                        grid-template-columns: repeat(2, 1fr); 
                        gap: 40px; 
                        margin-top: 50px; 
                    }
                    .signature-field { 
                        text-align: center; 
                        padding: 20px; 
                        border-top: 2px solid #333; 
                    }
                    .signature-field .signature-label { 
                        font-weight: bold; 
                        margin-bottom: 30px; 
                        color: #2a5298; 
                    }
                    .voucher-footer { 
                        text-align: center; 
                        margin-top: 40px; 
                        padding-top: 20px; 
                        border-top: 1px solid #ddd; 
                        color: #666; 
                        font-style: italic; 
                    }
                    @media print { 
                        body { margin: 10px; }
                        .btn, button, [onclick*="Edit"], [onclick*="edit"], [id*="Btn"], [id*="edit"], [id*="save"], [id*="cancel"] {
                            display: none !important;
                        }
                        .flex.justify-between.items-center > div:last-child {
                            display: none !important;
                        }
                        textarea, input[type="text"], input[type="number"], select {
                            border: none !important;
                            background: transparent !important;
                            -webkit-appearance: none;
                            -moz-appearance: none;
                            appearance: none;
                        }
                    }
                </style>
            </head>
            <body>
                ${voucherContent}
                <script>
                    window.onload = function() {
                        setTimeout(() => {
                            window.print();
                            window.close();
                        }, 500);
                    }
                </script>
            </body>
        </html>
    `);
    printWindow.document.close();
    
    showToast('PDF export ready!', 'success');
}

// Print voucher
function printVoucher() {
    if (!currentVoucherData) return;
    
    const printContent = document.getElementById('voucherContent').innerHTML;
    const originalContent = document.body.innerHTML;
    
    // Create print-friendly version
    document.body.innerHTML = `
        <style>
            body { font-family: Arial, sans-serif; margin: 20px; }
            .voucher { max-width: 600px; margin: 0 auto; }
            .voucher-header { text-align: center; margin-bottom: 30px; padding-bottom: 20px; border-bottom: 2px solid #ddd; }
            .voucher-header h2 { color: #2a5298; margin-bottom: 10px; }
            .voucher-info { display: grid; grid-template-columns: repeat(2, 1fr); gap: 20px; margin-bottom: 20px; }
            .voucher-field { display: flex; justify-content: space-between; padding: 10px 0; border-bottom: 1px solid #eee; }
            .voucher-field label { font-weight: 600; color: #666; }
            .voucher-amount { text-align: center; font-size: 1.5rem; font-weight: 700; color: #4CAF50; margin: 20px 0; padding: 15px; background: #f8f9fa; border-radius: 8px; }
            .voucher-description { margin: 20px 0; }
            .voucher-footer { margin-top: 30px; text-align: center; color: #666; }
            
            /* Hide all edit buttons and controls during printing */
            @media print {
                .voucher-footer { position: fixed; bottom: 20px; }
                .btn, button, [onclick*="Edit"], [onclick*="edit"], [id*="Btn"], [id*="edit"], [id*="save"], [id*="cancel"] {
                    display: none !important;
                }
                .flex.justify-between.items-center > div:last-child {
                    display: none !important;
                }
                textarea, input[type="text"], input[type="number"], select {
                    border: none !important;
                    background: transparent !important;
                    -webkit-appearance: none;
                    -moz-appearance: none;
                    appearance: none;
                }
            }
        </style>
        ${printContent}
    `;
    
    window.print();
    document.body.innerHTML = originalContent;
    
    // Re-initialize event listeners after restoring content
    initializeEventListeners();
}

// Description edit functionality
let originalDescription = '';

function toggleDescriptionEdit() {
    const displayDiv = document.getElementById('descriptionDisplay');
    const textarea = document.getElementById('descriptionTextarea');
    const editBtn = document.getElementById('editDescBtn');
    const saveBtn = document.getElementById('saveDescBtn');
    const cancelBtn = document.getElementById('cancelDescBtn');
    
    // Store original description
    originalDescription = displayDiv.textContent.trim();
    
    // Show textarea, hide display
    displayDiv.style.display = 'none';
    textarea.style.display = 'block';
    textarea.value = originalDescription === 'No description provided' ? '' : originalDescription;
    
    // Show save/cancel buttons, hide edit button
    editBtn.style.display = 'none';
    saveBtn.style.display = 'inline-flex';
    cancelBtn.style.display = 'inline-flex';
    
    // Focus on textarea
    textarea.focus();
}

function saveDescriptionEdit() {
    const displayDiv = document.getElementById('descriptionDisplay');
    const textarea = document.getElementById('descriptionTextarea');
    const editBtn = document.getElementById('editDescBtn');
    const saveBtn = document.getElementById('saveDescBtn');
    const cancelBtn = document.getElementById('cancelDescBtn');
    
    const newDescription = textarea.value.trim();
    
    // Update display
    displayDiv.textContent = newDescription || 'No description provided';
    
    // Update expense data
    if (currentVoucherData && currentVoucherIndex !== null) {
        expenseData[currentVoucherIndex].description = newDescription || '';
        currentVoucherData.description = newDescription || '';
        
        // Update the main table as well
        updateExpenseInTable(currentVoucherIndex, 'description', newDescription || '');
        
        // Save to server
        saveExpenseData();
        
        showToast('Description updated successfully!', 'success');
    }
    
    // Show display, hide textarea
    displayDiv.style.display = 'block';
    textarea.style.display = 'none';
    
    // Show edit button, hide save/cancel buttons
    editBtn.style.display = 'inline-flex';
    saveBtn.style.display = 'none';
    cancelBtn.style.display = 'none';
}

function cancelDescriptionEdit() {
    const displayDiv = document.getElementById('descriptionDisplay');
    const textarea = document.getElementById('descriptionTextarea');
    const editBtn = document.getElementById('editDescBtn');
    const saveBtn = document.getElementById('saveDescBtn');
    const cancelBtn = document.getElementById('cancelDescBtn');
    
    // Reset textarea to original value
    textarea.value = originalDescription === 'No description provided' ? '' : originalDescription;
    
    // Show display, hide textarea
    displayDiv.style.display = 'block';
    textarea.style.display = 'none';
    
    // Show edit button, hide save/cancel buttons
    editBtn.style.display = 'inline-flex';
    saveBtn.style.display = 'none';
    cancelBtn.style.display = 'none';
}

function updateExpenseInTable(index, field, value) {
    // Update the main table if it exists
    const tableBody = document.getElementById('tableBody');
    if (tableBody) {
        const rows = tableBody.getElementsByTagName('tr');
        for (let i = 0; i < rows.length; i++) {
            const cells = rows[i].getElementsByTagName('td');
            if (cells.length > 0) {
                const srNoCell = cells[1]; // S.No column
                if (srNoCell && parseInt(srNoCell.textContent) === expenseData[index].srNo) {
                    // Find the description column (index 8)
                    const descCell = cells[8];
                    if (descCell && field === 'description') {
                        descCell.textContent = value || '-';
                        descCell.setAttribute('data-field', 'description');
                        descCell.setAttribute('data-index', index);
                        descCell.contentEditable = 'true';
                    }
                    break;
                }
            }
        }
    }
}

async function saveExpenseData() {
    // Use the updated saveDataToCurrentSheet function
    await saveDataToCurrentSheet();
}

// Close voucher modal
function closeVoucherModal() {
    voucherModal.classList.remove('active');
    currentVoucherData = null;
    currentVoucherIndex = null;
    isVoucherEditMode = false;
}

// Delete expense and sync with Excel
async function deleteExpense(index) {
    if (confirm('Are you sure you want to delete this expense record?')) {
        try {
            const expenseToDelete = expenseData[index];
            
            showLoading('Deleting expense and updating Excel...');
            
            // Call server to delete expense
            const response = await fetch(`/delete/${expenseToDelete.srNo}`, {
                method: 'DELETE'
            });
            
            const result = await response.json();
            
            if (result.success) {
                // Remove from local data
                expenseData.splice(index, 1);
                
                // Re-render table with updated data
                renderTable();
                updateStatistics();
                
                showToast('Expense deleted and Excel file updated!', 'success');
                console.log('Deleted expense:', expenseToDelete);
            } else {
                showToast('Failed to delete expense', 'error');
                console.error('Delete error:', result.error);
            }
        } catch (error) {
            console.error('Error deleting expense:', error);
            showToast('Error deleting expense', 'error');
        } finally {
            hideLoading();
        }
    }
}

// Clear all data
function clearAllData() {
    if (expenseData.length === 0) {
        showToast('No data to clear', 'warning');
        return;
    }
    
    if (confirm('Are you sure you want to clear all expense records?')) {
        expenseData = [];
        renderTable();
        updateStatistics();
        showToast('All data cleared successfully!', 'success');
    }
}

// Update statistics
function updateStatistics() {
    const totalAmount = expenseData.reduce((sum, expense) => sum + parseFloat(expense.amount || 0), 0);
    const paidCount = expenseData.filter(expense => expense.status === 'Paid').length;
    const pendingCount = expenseData.filter(expense => expense.status === 'Pending').length;
    const totalCount = expenseData.length;

    document.getElementById('totalAmount').textContent = formatCurrency(totalAmount);
    document.getElementById('totalAmountInWords').textContent = numberToWords(totalAmount);
    document.getElementById('paidCount').textContent = paidCount;
    document.getElementById('pendingCount').textContent = pendingCount;
    document.getElementById('totalCount').textContent = totalCount;
}

// Show empty state
function showEmptyState() {
    emptyState.style.display = 'block';
    tableBody.parentElement.style.display = 'none';
}

// Hide empty state
function hideEmptyState() {
    emptyState.style.display = 'none';
    tableBody.parentElement.style.display = 'table';
}

// Show loading overlay
function showLoading(message = 'Processing...') {
    const loadingMessage = document.getElementById('loadingMessage');
    if (loadingMessage) {
        loadingMessage.textContent = message;
    }
    loadingOverlay.classList.add('active');
}

// Hide loading overlay
function hideLoading() {
    loadingOverlay.classList.remove('active');
}

// Show toast notification
function showToast(message, type = 'success') {
    toastMessage.textContent = message;
    
    // Remove existing type classes
    toast.classList.remove('error', 'warning');
    
    // Add new type class if not success
    if (type !== 'success') {
        toast.classList.add(type);
    }
    
    // Show toast
    toast.classList.add('show');
    
    // Hide after 3 seconds
    setTimeout(() => {
        toast.classList.remove('show');
    }, 3000);
}

// Utility function to format currency with Indian Rupee symbol and commas
function formatCurrency(amount) {
    const formattedAmount = new Intl.NumberFormat('en-IN', {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2
    }).format(amount);
    return `₹${formattedAmount}`;
}

// Utility function to format number with commas (without currency symbol)
function formatNumber(amount) {
    return new Intl.NumberFormat('en-IN', {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2
    }).format(amount);
}

// Utility function to generate random ID
function generateId() {
    return Date.now().toString(36) + Math.random().toString(36).substr(2);
}

// Convert number to words (Indian system)
function numberToWords(num) {
    if (num === 0) return 'Zero';
    
    const ones = ['', 'One', 'Two', 'Three', 'Four', 'Five', 'Six', 'Seven', 'Eight', 'Nine'];
    const teens = ['Ten', 'Eleven', 'Twelve', 'Thirteen', 'Fourteen', 'Fifteen', 'Sixteen', 'Seventeen', 'Eighteen', 'Nineteen'];
    const tens = ['', '', 'Twenty', 'Thirty', 'Forty', 'Fifty', 'Sixty', 'Seventy', 'Eighty', 'Ninety'];
    const thousands = ['', 'Thousand', 'Lakh', 'Crore'];
    
    // Convert to integer and handle decimal part
    const integerPart = Math.floor(num);
    const decimalPart = Math.round((num - integerPart) * 100);
    
    let words = '';
    
    // Convert integer part to words
    if (integerPart > 0) {
        words = convertIntegerToWords(integerPart);
    }
    
    // Add decimal part if exists
    if (decimalPart > 0) {
        if (words) words += ' and ';
        words += convertIntegerToWords(decimalPart) + ' Paise';
    }
    
    return words;
}

function convertIntegerToWords(num) {
    if (num === 0) return '';
    
    const ones = ['', 'One', 'Two', 'Three', 'Four', 'Five', 'Six', 'Seven', 'Eight', 'Nine'];
    const teens = ['Ten', 'Eleven', 'Twelve', 'Thirteen', 'Fourteen', 'Fifteen', 'Sixteen', 'Seventeen', 'Eighteen', 'Nineteen'];
    const tens = ['', '', 'Twenty', 'Thirty', 'Forty', 'Fifty', 'Sixty', 'Seventy', 'Eighty', 'Ninety'];
    
    let words = '';
    
    if (num < 10) {
        words = ones[num];
    } else if (num < 20) {
        words = teens[num - 10];
    } else if (num < 100) {
        words = tens[Math.floor(num / 10)] + ' ' + ones[num % 10];
    } else if (num < 1000) {
        words = ones[Math.floor(num / 100)] + ' Hundred ' + convertIntegerToWords(num % 100);
    } else if (num < 100000) {
        words = convertIntegerToWords(Math.floor(num / 1000)) + ' Thousand ' + convertIntegerToWords(num % 1000);
    } else if (num < 10000000) {
        words = convertIntegerToWords(Math.floor(num / 100000)) + ' Lakh ' + convertIntegerToWords(num % 100000);
    } else {
        words = convertIntegerToWords(Math.floor(num / 10000000)) + ' Crore ' + convertIntegerToWords(num % 10000000);
    }
    
    return words.trim();
}

// Enhanced Export Functions
function getCurrentExportData() {
    return filteredData.length > 0 ? filteredData : expenseData;
}

// Export to CSV
function exportToCSV() {
    const dataToExport = getCurrentExportData();
    
    if (dataToExport.length === 0) {
        showToast('No data to export', 'warning');
        return;
    }

    const headers = ['Sr.No', 'Date', 'Given To', 'Amount', 'Mode', 'Description', 'Fund', 'Status'];
    const csvContent = [
        headers.join(','),
        ...dataToExport.map(expense => [
            expense.srNo || '',
            expense.date || '',
            `"${(expense.givenTo || '').replace(/"/g, '""')}"`,
            expense.amount || '',
            expense.mode || '',
            `"${(expense.description || '').replace(/"/g, '""')}"`,
            `"${(expense.fund || '').replace(/"/g, '""')}"`,
            expense.status || ''
        ].join(','))
    ].join('\n');

    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `expenses_${new Date().toISOString().split('T')[0]}.csv`;
    a.click();
    window.URL.revokeObjectURL(url);

    showToast(`Exported ${dataToExport.length} records to CSV!`, 'success');
}

// Export to Excel
function exportToExcel() {
    const dataToExport = getCurrentExportData();
    
    if (dataToExport.length === 0) {
        showToast('No data to export', 'warning');
        return;
    }

    // Create CSV content for Excel
    const headers = ['Sr.No', 'Date', 'Given To', 'Amount', 'Mode', 'Description', 'Fund', 'Status'];
    const csvContent = [
        headers.join('\t'),
        ...dataToExport.map(expense => [
            expense.srNo || '',
            expense.date || '',
            expense.givenTo || '',
            expense.amount || '',
            expense.mode || '',
            expense.description || '',
            expense.fund || '',
            expense.status || ''
        ].join('\t'))
    ].join('\n');

    const blob = new Blob([csvContent], { type: 'application/vnd.ms-excel;charset=utf-8;' });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `expenses_${new Date().toISOString().split('T')[0]}.xls`;
    a.click();
    window.URL.revokeObjectURL(url);

    showToast(`Exported ${dataToExport.length} records to Excel!`, 'success');
}

// Export to PDF
function exportToPDF() {
    const dataToExport = getCurrentExportData();
    
    if (dataToExport.length === 0) {
        showToast('No data to export', 'warning');
        return;
    }

    // Create printable HTML content
    const printContent = createPrintContent(dataToExport);
    
    // Open in new window for PDF printing
    const printWindow = window.open('', '_blank');
    printWindow.document.write(`
        <html>
            <head>
                <title>Expense Report</title>
                <style>
                    body { font-family: Arial, sans-serif; margin: 20px; color: #333; }
                    h1 { color: #2a5298; text-align: center; margin-bottom: 30px; }
                    table { width: 100%; border-collapse: collapse; margin-bottom: 20px; }
                    th, td { border: 1px solid #ddd; padding: 12px; text-align: left; }
                    th { background-color: #2a5298; color: white; font-weight: bold; }
                    tr:nth-child(even) { background-color: #f9f9f9; }
                    .summary { margin-top: 30px; padding: 20px; background: #f8f9fa; border-radius: 8px; }
                    .summary-item { display: flex; justify-content: space-between; margin-bottom: 10px; }
                    .summary-item strong { color: #2a5298; }
                    @media print { body { margin: 10px; } }
                </style>
            </head>
            <body>
                ${printContent}
                <script>
                    window.onload = function() {
                        setTimeout(() => {
                            window.print();
                            window.close();
                        }, 500);
                    }
                </script>
            </body>
        </html>
    `);
    printWindow.document.close();

    showToast(`PDF export ready for ${dataToExport.length} records!`, 'success');
}

// Print expenses
function printExpenses() {
    const dataToExport = getCurrentExportData();
    
    if (dataToExport.length === 0) {
        showToast('No data to print', 'warning');
        return;
    }

    const printContent = createPrintContent(dataToExport);
    
    // Store original content
    const originalContent = document.body.innerHTML;
    
    // Replace body with print content
    document.body.innerHTML = `
        <style>
            body { font-family: Arial, sans-serif; margin: 20px; color: #333; }
            h1 { color: #2a5298; text-align: center; margin-bottom: 30px; }
            table { width: 100%; border-collapse: collapse; margin-bottom: 20px; }
            th, td { border: 1px solid #ddd; padding: 12px; text-align: left; }
            th { background-color: #2a5298; color: white; font-weight: bold; }
            tr:nth-child(even) { background-color: #f9f9f9; }
            .summary { margin-top: 30px; padding: 20px; background: #f8f9fa; border-radius: 8px; }
            .summary-item { display: flex; justify-content: space-between; margin-bottom: 10px; }
            .summary-item strong { color: #2a5298; }
            @media print { body { margin: 10px; } }
        </style>
        ${printContent}
    `;
    
    // Print
    window.print();
    
    // Restore original content
    document.body.innerHTML = originalContent;
    
    // Re-initialize event listeners
    initializeEventListeners();
    loadExistingData();

    showToast(`Printed ${dataToExport.length} records!`, 'success');
}

// Create print content
function createPrintContent(data) {
    const totalAmount = data.reduce((sum, expense) => sum + parseFloat(expense.amount || 0), 0);
    const paidCount = data.filter(expense => expense.status === 'Paid').length;
    const pendingCount = data.filter(expense => expense.status === 'Pending').length;
    
    const tableRows = data.map(expense => `
        <tr>
            <td>${expense.srNo || ''}</td>
            <td>${formatDate(expense.date)}</td>
            <td>${expense.givenTo || '-'}</td>
            <td>${formatCurrency(parseFloat(expense.amount || 0))}</td>
            <td>${numberToWords(parseFloat(expense.amount || 0))}</td>
            <td>${expense.mode || '-'}</td>
            <td>${expense.description || '-'}</td>
            <td>${expense.fund || '-'}</td>
            <td>${expense.status || 'Pending'}</td>
        </tr>
    `).join('');
    
    return `
        <h1>Expense Management Report</h1>
        <p style="text-align: center; color: #666; margin-bottom: 30px;">
            Generated on ${formatDate(new Date().toISOString())} | 
            Total Records: ${data.length}
        </p>
        
        <table>
            <thead>
                <tr>
                    <th>Sr.No</th>
                    <th>Date</th>
                    <th>Given To</th>
                    <th>Amount</th>
                    <th>Amount in Words</th>
                    <th>Mode</th>
                    <th>Description</th>
                    <th>Fund</th>
                    <th>Status</th>
                </tr>
            </thead>
            <tbody>
                ${tableRows}
            </tbody>
        </table>
        
        <div class="summary">
            <h3 style="margin-top: 0; color: #2a5298;">Summary</h3>
            <div class="summary-item">
                <span>Total Amount:</span>
                <strong>${formatCurrency(totalAmount)}</strong>
            </div>
            <div class="summary-item">
                <span>Total Amount in Words:</span>
                <strong>${numberToWords(totalAmount)}</strong>
            </div>
            <div class="summary-item">
                <span>Paid Expenses:</span>
                <strong>${paidCount}</strong>
            </div>
            <div class="summary-item">
                <span>Pending Expenses:</span>
                <strong>${pendingCount}</strong>
            </div>
            <div class="summary-item">
                <span>Total Records:</span>
                <strong>${data.length}</strong>
            </div>
        </div>
    `;
}

// Search and Filter Functions
function applyFilters() {
    const searchTerm = searchInput.value.toLowerCase().trim();
    const statusValue = statusFilter.value;
    const modeValue = modeFilter.value;
    const givenToValue = givenToFilter.value;
    const dateValue = dateFilter.value;
    
    filteredData = expenseData.filter(expense => {
        // Search filter
        const matchesSearch = !searchTerm || 
            expense.givenTo?.toLowerCase().includes(searchTerm) ||
            expense.description?.toLowerCase().includes(searchTerm) ||
            expense.fund?.toLowerCase().includes(searchTerm) ||
            expense.mode?.toLowerCase().includes(searchTerm);
        
        // Status filter
        const matchesStatus = !statusValue || expense.status === statusValue;
        
        // Mode filter
        const matchesMode = !modeValue || expense.mode === modeValue;
        
        // Given To filter
        const matchesGivenTo = !givenToValue || expense.givenTo === givenToValue;
        
        // Date filter
        let matchesDate = !dateValue;
        if (dateValue && expense.date) {
            const expenseDate = new Date(expense.date).toISOString().split('T')[0];
            matchesDate = expenseDate === dateValue;
        }
        
        return matchesSearch && matchesStatus && matchesMode && matchesGivenTo && matchesDate;
    });
    
    renderTable();
    updateFilteredStatistics();
}

function clearAllFilters() {
    searchInput.value = '';
    statusFilter.value = '';
    modeFilter.value = '';
    givenToFilter.value = '';
    dateFilter.value = '';
    filteredData = [];
    renderTable();
    updateStatistics();
}

function updateFilteredStatistics() {
    const dataToUse = filteredData.length > 0 ? filteredData : expenseData;
    const totalAmount = dataToUse.reduce((sum, expense) => sum + parseFloat(expense.amount || 0), 0);
    const paidCount = dataToUse.filter(expense => expense.status === 'Paid').length;
    const pendingCount = dataToUse.filter(expense => expense.status === 'Pending').length;
    const totalCount = dataToUse.length;

    document.getElementById('totalAmount').textContent = formatCurrency(totalAmount);
    document.getElementById('totalAmountInWords').textContent = numberToWords(totalAmount);
    document.getElementById('paidCount').textContent = paidCount;
    document.getElementById('pendingCount').textContent = pendingCount;
    document.getElementById('totalCount').textContent = totalCount;
}

// Dropdown Recommendation Functions
function initializeDropdowns() {
    const givenToInput = document.getElementById('givenTo');
    const fundInput = document.getElementById('fund');
    const modeSelect = document.getElementById('mode');
    
    // Given To dropdown
    givenToInput.addEventListener('input', () => showDropdown('givenTo'));
    givenToInput.addEventListener('focus', () => showDropdown('givenTo'));
    
    // Fund dropdown  
    fundInput.addEventListener('input', () => showDropdown('fund'));
    fundInput.addEventListener('focus', () => showDropdown('fund'));
    
    // Mode dropdown (show on focus)
    modeSelect.addEventListener('focus', () => showDropdown('mode'));
    
    // Hide dropdowns on outside click
    document.addEventListener('click', (e) => {
        if (!e.target.closest('.dropdown-input')) {
            hideAllDropdowns();
        }
    });
}

function getUniqueValues(field) {
    const values = new Set();
    expenseData.forEach(expense => {
        const value = expense[field];
        if (value && value.trim()) {
            values.add(value.trim());
        }
    });
    
    // Add common modes if field is 'mode' and no data exists
    if (field === 'mode' && values.size === 0) {
        values.add('Cash');
        values.add('Card');
        values.add('Online');
        values.add('Cheque');
    }
    
    return Array.from(values).sort();
}

function showDropdown(field) {
    hideAllDropdowns();
    
    const input = field === 'mode' ? document.getElementById('mode') : document.getElementById(field);
    const dropdown = document.getElementById(`${field}Dropdown`);
    const searchValue = field === 'mode' ? '' : input.value.toLowerCase().trim();
    
    const uniqueValues = getUniqueValues(field);
    let filteredValues = uniqueValues;
    
    if (searchValue && field !== 'mode') {
        filteredValues = uniqueValues.filter(value => 
            value.toLowerCase().includes(searchValue)
        );
    }
    
    dropdown.innerHTML = '';
    
    if (filteredValues.length === 0) {
        const noResults = document.createElement('div');
        noResults.className = 'dropdown-item no-results';
        noResults.textContent = 'No suggestions found';
        dropdown.appendChild(noResults);
    } else {
        filteredValues.forEach(value => {
            const item = document.createElement('div');
            item.className = 'dropdown-item';
            item.textContent = value;
            
            item.addEventListener('click', () => {
                if (field === 'mode') {
                    document.getElementById('mode').value = value;
                } else {
                    input.value = value;
                }
                hideAllDropdowns();
                input.focus();
            });
            
            dropdown.appendChild(item);
        });
    }
    
    dropdown.classList.add('active');
}

function hideAllDropdowns() {
    document.querySelectorAll('.dropdown-list').forEach(dropdown => {
        dropdown.classList.remove('active');
    });
}

// Add Expense Modal Functions
function openAddExpenseModal() {
    // Set today's date as default
    document.getElementById('expenseDate').value = new Date().toISOString().split('T')[0];
    addExpenseModal.classList.add('active');
    
    // Initialize dropdowns with current data
    setTimeout(() => initializeDropdowns(), 100);
}

function closeAddExpenseModal() {
    addExpenseModal.classList.remove('active');
    // Reset form and hide dropdowns
    document.getElementById('addExpenseForm').reset();
    hideAllDropdowns();
}

// Add Sheet Modal Functions
function openAddSheetModal() {
    elements.sheetNameInput.value = '';
    elements.addSheetModal.classList.add('active');
    elements.sheetNameInput.focus();
}

async function createNewSheet() {
    try {
        const sheetName = elements.sheetNameInput.value.trim();
        
        if (!sheetName) {
            showToast('Please enter a sheet name', 'error');
            return;
        }
        
        showLoading('Creating new worksheet...');
        
        // Send to server
        const response = await fetch('/api/add-sheet', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({ sheetName })
        });
        
        const result = await response.json();
        
        if (result.success) {
            showToast(`Sheet "${result.sheet.name}" created successfully!`, 'success');
            elements.addSheetModal.classList.remove('active');
            elements.sheetNameInput.value = '';
            
            // Optional: Refresh data to show updated sheets
            setTimeout(() => {
                refreshData();
            }, 1000);
        } else {
            showToast(result.error || 'Failed to create sheet', 'error');
        }
    } catch (error) {
        console.error('Error creating sheet:', error);
        showToast('Error creating sheet: ' + error.message, 'error');
    } finally {
        hideLoading();
    }
}

async function addNewExpense() {
    try {
        // Get form values
        const expenseData = {
            sheetName: currentSheet,
            date: document.getElementById('expenseDate').value,
            givenTo: document.getElementById('givenTo').value.trim(),
            amount: parseFloat(document.getElementById('amount').value).toFixed(2),
            mode: document.getElementById('mode').value,
            description: document.getElementById('description').value.trim(),
            fund: document.getElementById('fund').value.trim(),
            status: document.getElementById('status').value
        };
        
        // Validation
        if (!expenseData.date || !expenseData.givenTo || !expenseData.amount || !expenseData.mode) {
            showToast('Please fill all required fields', 'error');
            return;
        }
        
        if (isNaN(expenseData.amount) || parseFloat(expenseData.amount) <= 0) {
            showToast('Please enter a valid amount', 'error');
            return;
        }
        
        showLoading('Adding expense and updating Excel...');
        
        // Send to server
        const response = await fetch('/add-expense', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(expenseData)
        });
        
        const result = await response.json();
        
        if (result.success) {
            // Close modal and reset form
            closeAddExpenseModal();
            
            // Reload data from server to get updated list
            await loadExistingData();
            
            // Clear filters to show new expense
            clearAllFilters();
            
            showToast(`Expense added successfully to sheet "${currentSheet}"!`, 'success');
        } else {
            showToast(result.error || 'Failed to add expense', 'error');
        }
    } catch (error) {
        console.error('Error adding expense:', error);
        showToast('Error adding expense', 'error');
    } finally {
        hideLoading();
    }
}

// Keyboard shortcuts
document.addEventListener('keydown', (e) => {
    // Ctrl/Cmd + I for import
    if ((e.ctrlKey || e.metaKey) && e.key === 'i') {
        e.preventDefault();
        importBtn.click();
    }
    
    // Ctrl/Cmd + A for add expense
    if ((e.ctrlKey || e.metaKey) && e.key === 'a') {
        e.preventDefault();
        openAddExpenseModal();
    }
    
    // Ctrl/Cmd + E for export
    if ((e.ctrlKey || e.metaKey) && e.key === 'e') {
        e.preventDefault();
        exportToCSV();
    }
    
    // Ctrl/Cmd + F for search focus
    if ((e.ctrlKey || e.metaKey) && e.key === 'f') {
        e.preventDefault();
        searchInput.focus();
    }
    
    // Delete key to clear all
    if (e.key === 'Delete' && e.shiftKey) {
        e.preventDefault();
        clearAllData();
    }
});

// Smart Bulk Voucher Synchronization System
function synchronizeBulkVoucherData(updatedData, voucherIndex) {
    try {
        // Find if this voucher is currently selected for bulk printing
        const srNo = updatedData.srNo || (voucherIndex + 1);
        
        if (selectedVouchers.has(srNo)) {
            // Update the bulk voucher cache with new data
            console.log(`🔄 Synchronizing bulk voucher data for SR No: ${srNo}`);
            
            // Update any open bulk voucher modal if it exists
            const bulkModal = document.getElementById('bulkVoucherModal');
            if (bulkModal && bulkModal.classList.contains('active')) {
                // Update the selection summary to reflect changes
                updateBulkSelectionSummary();
                
                // Show subtle notification that data was synchronized
                showBulkSyncNotification(srNo);
            }
            
            // If bulk voucher print window is open, update its content
            updateBulkPrintWindowIfOpen(updatedData, srNo);
            
            // Log the synchronization for debugging
            console.log('✅ Bulk voucher data synchronized successfully:', {
                srNo: srNo,
                updatedFields: Object.keys(updatedData),
                timestamp: new Date().toISOString()
            });
        }
    } catch (error) {
        console.error('❌ Error synchronizing bulk voucher data:', error);
        // Don't show error to user as it's a background operation
    }
}

// Show subtle notification for bulk synchronization
function showBulkSyncNotification(srNo) {
    // Create a subtle notification element
    const notification = document.createElement('div');
    notification.style.cssText = `
        position: fixed;
        top: 20px;
        right: 20px;
        background: linear-gradient(135deg, #10b981, #059669);
        color: white;
        padding: 12px 20px;
        border-radius: 8px;
        box-shadow: 0 4px 12px rgba(16, 185, 129, 0.3);
        z-index: 10000;
        font-size: 14px;
        font-weight: 500;
        animation: slideIn 0.3s ease-out;
        max-width: 300px;
    `;
    
    notification.innerHTML = `
        <div style="display: flex; align-items: center; gap: 10px;">
            <i class="fas fa-sync-alt" style="animation: spin 1s linear infinite;"></i>
            <span>Voucher #${srNo} updated in bulk selection</span>
        </div>
    `;
    
    // Add animation styles
    const style = document.createElement('style');
    style.textContent = `
        @keyframes slideIn {
            from { transform: translateX(100%); opacity: 0; }
            to { transform: translateX(0); opacity: 1; }
        }
        @keyframes spin {
            from { transform: rotate(0deg); }
            to { transform: rotate(360deg); }
        }
    `;
    document.head.appendChild(style);
    
    document.body.appendChild(notification);
    
    // Auto-remove after 3 seconds
    setTimeout(() => {
        notification.style.animation = 'slideIn 0.3s ease-out reverse';
        setTimeout(() => {
            notification.remove();
            style.remove();
        }, 300);
    }, 3000);
}

// Update bulk print window if it's open
function updateBulkPrintWindowIfOpen(updatedData, srNo) {
    // Check if there's a bulk print window open and update it
    // This is a more advanced feature for future implementation
    try {
        // For now, we'll just log that this would update open print windows
        console.log('📋 Bulk print window update check for SR No:', srNo);
    } catch (error) {
        console.log('Bulk print window update not available:', error.message);
    }
}

// ========== BULK VOUCHER FUNCTIONALITY ==========

// Show bulk voucher selection modal
function showBulkVoucherModal() {
    updateBulkSelectionSummary();
    
    // Show real-time synchronization status
    showBulkSyncStatus();
    
    elements.bulkVoucherModal.classList.add('active');
}

// Show real-time synchronization status in bulk modal
function showBulkSyncStatus() {
    const modalBody = elements.bulkVoucherModal.querySelector('.modal-body');
    if (!modalBody) return;
    
    // Check if any selected vouchers have been recently updated
    const selectedExpenses = expenseData.filter(expense => selectedVouchers.has(expense.srNo));
    
    if (selectedExpenses.length > 0) {
        // Add a sync status indicator
        let syncStatusDiv = document.getElementById('bulkSyncStatus');
        if (!syncStatusDiv) {
            syncStatusDiv = document.createElement('div');
            syncStatusDiv.id = 'bulkSyncStatus';
            syncStatusDiv.style.cssText = `
                background: linear-gradient(135deg, #f0fdf4, #dcfce7);
                border: 1px solid #86efac;
                border-radius: 8px;
                padding: 12px;
                margin-bottom: 20px;
                font-size: 14px;
                color: #166534;
                display: flex;
                align-items: center;
                gap: 10px;
            `;
            
            syncStatusDiv.innerHTML = `
                <i class="fas fa-sync-alt" style="color: #16a34a; animation: pulse 2s infinite;"></i>
                <span>
                    <strong>Smart Sync Active:</strong> Changes to individual vouchers will automatically update in bulk selection
                </span>
            `;
            
            // Insert at the beginning of modal body
            modalBody.insertBefore(syncStatusDiv, modalBody.firstChild);
            
            // Add pulse animation
            const style = document.createElement('style');
            style.textContent = `
                @keyframes pulse {
                    0%, 100% { opacity: 1; }
                    50% { opacity: 0.5; }
                }
            `;
            document.head.appendChild(style);
        }
    }
}

// Toggle voucher selection
function toggleVoucherSelection(srNo) {
    if (selectedVouchers.has(srNo)) {
        selectedVouchers.delete(srNo);
    } else {
        selectedVouchers.add(srNo);
    }
    updateSelectAllCheckbox();
    updateBulkSelectionSummary();
}

// Toggle all selections
function toggleAllSelections() {
    const dataToRender = filteredData.length > 0 ? filteredData : expenseData;
    const isChecked = elements.selectAllCheckbox.checked;
    
    if (isChecked) {
        dataToRender.forEach(expense => {
            selectedVouchers.add(expense.srNo);
        });
    } else {
        selectedVouchers.clear();
    }
    
    // Update all checkboxes
    const checkboxes = document.querySelectorAll('.voucher-checkbox');
    checkboxes.forEach(checkbox => {
        checkbox.checked = isChecked;
    });
    
    updateBulkSelectionSummary();
}

// Update select all checkbox state
function updateSelectAllCheckbox() {
    const dataToRender = filteredData.length > 0 ? filteredData : expenseData;
    const checkboxes = document.querySelectorAll('.voucher-checkbox');
    
    if (checkboxes.length === 0) {
        elements.selectAllCheckbox.checked = false;
        return;
    }
    
    const allChecked = Array.from(checkboxes).every(cb => cb.checked);
    elements.selectAllCheckbox.checked = allChecked;
}

// Update bulk selection summary
function updateBulkSelectionSummary() {
    const totalRecords = filteredData.length > 0 ? filteredData.length : expenseData.length;
    const selectedRecords = selectedVouchers.size;
    
    document.getElementById('totalRecords').textContent = totalRecords;
    document.getElementById('selectedRecords').textContent = selectedRecords;
}

// Select all expenses
function selectAllExpenses() {
    const dataToRender = filteredData.length > 0 ? filteredData : expenseData;
    dataToRender.forEach(expense => {
        selectedVouchers.add(expense.srNo);
    });
    
    // Update all checkboxes
    const checkboxes = document.querySelectorAll('.voucher-checkbox');
    checkboxes.forEach(checkbox => {
        checkbox.checked = true;
    });
    
    elements.selectAllCheckbox.checked = true;
    updateBulkSelectionSummary();
}

// Deselect all expenses
function deselectAllExpenses() {
    selectedVouchers.clear();
    
    // Update all checkboxes
    const checkboxes = document.querySelectorAll('.voucher-checkbox');
    checkboxes.forEach(checkbox => {
        checkbox.checked = false;
    });
    
    elements.selectAllCheckbox.checked = false;
    updateBulkSelectionSummary();
}

// Filter by status
function filterByStatus() {
    const filterPaid = document.getElementById('filterPaid').checked;
    const filterPending = document.getElementById('filterPending').checked;
    
    if (!filterPaid && !filterPending) {
        // No filters applied, show all
        applyFilters();
        return;
    }
    
    const dataToRender = filteredData.length > 0 ? filteredData : expenseData;
    const filtered = dataToRender.filter(expense => {
        if (filterPaid && expense.status === 'Paid') return true;
        if (filterPending && expense.status === 'Pending') return true;
        return false;
    });
    
    // Update filtered data temporarily
    const originalFiltered = [...filteredData];
    filteredData = filtered;
    renderTable();
    filteredData = originalFiltered;
}

// Filter by date range
function filterByDate() {
    const startDate = document.getElementById('startDate').value;
    const endDate = document.getElementById('endDate').value;
    
    if (!startDate && !endDate) {
        applyFilters();
        return;
    }
    
    const dataToRender = filteredData.length > 0 ? filteredData : expenseData;
    const filtered = dataToRender.filter(expense => {
        const expenseDate = new Date(expense.date);
        if (startDate && expenseDate < new Date(startDate)) return false;
        if (endDate && expenseDate > new Date(endDate)) return false;
        return true;
    });
    
    // Update filtered data temporarily
    const originalFiltered = [...filteredData];
    filteredData = filtered;
    renderTable();
    filteredData = originalFiltered;
}

// Generate bulk voucher HTML (original with Tailwind classes)
function generateBulkVoucherHTML(selectedExpenses) {
    let bulkHTML = '<div class="bulk-voucher-container">';
    
    selectedExpenses.forEach((expense, index) => {
        const amount = parseFloat(expense.amount || 0);
        const amountInWords = numberToWords(amount);
        
        bulkHTML += `
            <div class="voucher-page" style="page-break-after: always; margin-bottom: 40px;">
                <div class="bg-white relative overflow-hidden shadow-2xl border-4 border-black" style="margin-bottom: 20px;">
                    <!-- Watermark Background -->
                    <div class="absolute inset-0 opacity-5 pointer-events-none">
                        <div class="text-9xl font-bold text-blue-900 transform rotate-45 absolute top-1/2 left-1/2 -translate-x-1/2 -translate-y-1/2">
                            EXPENSE
                        </div>
                    </div>
                    
                    <!-- Header Section -->
                    <div class="bg-gradient-to-r from-blue-900 to-blue-700 text-white p-6 relative">
                        <div class="text-center">
                            <h1 class="text-4xl font-bold mb-2">Shahrullah il Moazzam - 1447</h1>
                            <p class="text-3xl text-black mb-4">Shahrullah il Moazzam - 1447</p>
                            
                            <div class="relative inline-block">
                                <h2 class="font-playfair text-lg font-bold text-gray-900">EXPENSE VOUCHER</h2>
                                <!-- Double Underline -->
                                <div class="absolute bottom-0 left-0 right-0 h-1 bg-black"></div>
                                <div class="absolute -bottom-2 left-0 right-0 h-0.5 bg-black"></div>
                            </div>
                        </div>
                    </div>
                    
                    <!-- Voucher Details -->
                    <div class="p-8">
                        <div class="grid grid-cols-2 gap-8 mb-8">
                            <div>
                                <p class="text-sm text-gray-600 mb-2">Voucher No:</p>
                                <p class="text-xl font-bold text-gray-900" data-field="srNo">${expense.srNo}</p>
                            </div>
                            <div>
                                <p class="text-sm text-gray-600 mb-2">Date:</p>
                                <p class="text-xl font-bold text-gray-900" data-field="date">${formatDate(expense.date)}</p>
                            </div>
                        </div>
                        
                        <div class="mb-8">
                            <p class="text-sm text-gray-600 mb-2">Given To:</p>
                            <p class="text-xl font-bold text-gray-900 border-b-2 border-gray-300 pb-2" data-field="givenTo">${expense.givenTo || '-'}</p>
                        </div>
                        
                        <div class="mb-8">
                            <p class="text-sm text-gray-600 mb-2">Amount:</p>
                            <div class="flex items-center justify-between bg-green-50 p-6 rounded-lg border-2 border-green-200">
                                <div>
                                    <p class="text-3xl font-bold text-green-800" data-field="amount">₹${amount.toFixed(2)}</p>
                                    <p class="text-lg text-green-600 mt-2">${amountInWords}</p>
                                </div>
                                <div class="text-right">
                                    <p class="text-sm text-gray-600">Payment Mode</p>
                                    <p class="text-xl font-bold text-gray-900" data-field="mode">${expense.mode || '-'}</p>
                                </div>
                            </div>
                        </div>
                        
                        <div class="mb-8">
                            <p class="text-sm text-gray-600 mb-2">Description:</p>
                            <p class="text-lg text-gray-800 bg-gray-50 p-4 rounded-lg border" data-field="description">${expense.description || 'No description provided'}</p>
                        </div>
                        
                        <div class="grid grid-cols-2 gap-8 mb-8">
                            <div>
                                <p class="text-sm text-gray-600 mb-2">Fund:</p>
                                <p class="text-lg font-bold text-gray-900 border-b border-gray-300 pb-1" data-field="fund">${expense.fund || '-'}</p>
                            </div>
                            <div>
                                <p class="text-sm text-gray-600 mb-2">Status:</p>
                                <span class="inline-block px-4 py-2 text-lg font-bold rounded-full ${expense.status === 'Paid' ? 'bg-green-100 text-green-800' : 'bg-yellow-100 text-yellow-800'}" data-field="status">
                                    ${expense.status || 'Pending'}
                                </span>
                            </div>
                        </div>
                        
                        <!-- Signature Section -->
                        <div class="border-t-4 border-black pt-6 mt-12">
                            <div class="grid grid-cols-3 gap-8 text-center">
                                <div>
                                    <div class="border-b-2 border-gray-400 pb-2 mb-2">
                                        <p class="text-sm text-gray-600">Prepared By</p>
                                    </div>
                                </div>
                                <div>
                                    <div class="border-b-2 border-gray-400 pb-2 mb-2">
                                        <p class="text-sm text-gray-600">Verified By</p>
                                    </div>
                                </div>
                                <div>
                                    <div class="border-b-2 border-gray-400 pb-2 mb-2">
                                        <p class="text-sm text-gray-600">Approved By</p>
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        `;
    });
    
    bulkHTML += '</div>';
    return bulkHTML;
}

// Generate printable bulk voucher HTML with inline styles matching single voucher exactly
function generatePrintableBulkVoucherHTML(selectedExpenses) {
    let bulkHTML = '<div class="bulk-voucher-container">';
    
    selectedExpenses.forEach((expense, index) => {
        const amount = parseFloat(expense.amount || 0);
        const amountInWords = numberToWords(amount);
        
        bulkHTML += `
            <div class="voucher-page">
                <div class="print-area">
                    <div class="bg-white relative overflow-hidden shadow-2xl border-4 border-black" id="premiumVoucher_${index}">
                        <!-- Watermark Background -->
                        <div class="absolute inset-0 opacity-5 pointer-events-none">
                            <div class="text-9xl font-bold text-blue-900 transform rotate-45 absolute top-1/2 left-1/2 -translate-x-1/2 -translate-y-1/2" style="font-size: 9rem; font-weight: 900; color: #1e3a8a; transform: rotate(45deg); position: absolute; top: 50%; left: 50%; transform: translate(-50%, -50%) rotate(45deg);">
                                ASJ SAGWARA
                            </div>
                        </div>
                        
                        <!-- Header Section -->
                        <div class="relative z-10 text-center py-6 border-b-4 border-black" style="position: relative; z-index: 10; text-align: center; padding: 1.5rem 0; border-bottom: 4px solid black;">
                            <h1 class="font-playfair text-3xl font-bold text-gray-900 mb-2" style="font-family: 'Playfair Display', serif; font-size: 1.875rem; font-weight: 700; color: #111827; margin-bottom: 0.5rem;">
                                Anjuman E Saifee Jamaat, Sagwara
                            </h1>
                            <p class="text-3xl text-black mb-4" style="font-size: 1.875rem; color: black; margin-bottom: 1rem;">Shahrullah il Moazzam - 1447</p>
                            
                            <div class="relative inline-block" style="position: relative; display: inline-block;">
                                <h2 class="font-playfair text-lg font-bold text-gray-900" style="font-family: 'Playfair Display', serif; font-size: 1.125rem; font-weight: 700; color: #111827;">EXPENSE VOUCHER</h2>
                                <!-- Double Underline -->
                                <div class="absolute bottom-0 left-0 right-0 h-1 bg-black" style="position: absolute; bottom: 0; left: 0; right: 0; height: 0.25rem; background-color: black;"></div>
                                <div class="absolute -bottom-2 left-0 right-0 h-0.5 bg-black" style="position: absolute; bottom: -0.5rem; left: 0; right: 0; height: 0.125rem; background-color: black;"></div>
                            </div>
                        </div>
                        
                        <!-- Main Content -->
                        <div class="relative z-10 p-8" style="position: relative; z-index: 10; padding: 2rem;">
                            <!-- Grid Layout -->
                            <div class="grid grid-cols-4 gap-0 border-2 border-black" style="display: grid; grid-template-columns: repeat(4, 1fr); gap: 0; border: 2px solid black;">
                                <!-- S.No -->
                                <div class="border-r-2 border-b-2 border-black p-3 bg-gray-50" style="border-right: 2px solid black; border-bottom: 2px solid black; padding: 0.75rem; background-color: #f9fafb;">
                                    <div class="font-bold text-sm text-gray-800" style="font-weight: 700; font-size: 0.875rem; color: #1f2937;">S.No</div>
                                </div>
                                <div class="border-r-2 border-b-2 border-black p-3" style="border-right: 2px solid black; border-bottom: 2px solid black; padding: 0.75rem;">
                                    <div class="text-center font-semibold" data-field="srNo" style="text-align: center; font-weight: 600;">${expense.srNo || index + 1}</div>
                                </div>
                                
                                <!-- Date -->
                                <div class="border-r-2 border-b-2 border-black p-3 bg-gray-50" style="border-right: 2px solid black; border-bottom: 2px solid black; padding: 0.75rem; background-color: #f9fafb;">
                                    <div class="font-bold text-sm text-gray-800" style="font-weight: 700; font-size: 0.875rem; color: #1f2937;">Date</div>
                                </div>
                                <div class="border-b-2 border-black p-3" style="border-bottom: 2px solid black; padding: 0.75rem;">
                                    <div class="text-center font-semibold" data-field="date" style="text-align: center; font-weight: 600;">${formatDate(expense.date)}</div>
                                </div>
                                
                                <!-- Name -->
                                <div class="border-r-2 border-b-2 border-black p-3 bg-gray-50" style="border-right: 2px solid black; border-bottom: 2px solid black; padding: 0.75rem; background-color: #f9fafb;">
                                    <div class="font-bold text-sm text-gray-800" style="font-weight: 700; font-size: 0.875rem; color: #1f2937;">Name</div>
                                </div>
                                <div class="border-r-2 border-b-2 border-black p-3" style="border-right: 2px solid black; border-bottom: 2px solid black; padding: 0.75rem;">
                                    <div class="text-center font-semibold" data-field="givenTo" style="text-align: center; font-weight: 600;">${expense.givenTo || '-'}</div>
                                </div>
                                
                                <!-- Mode of Payment -->
                                <div class="border-r-2 border-b-2 border-black p-3 bg-gray-50" style="border-right: 2px solid black; border-bottom: 2px solid black; padding: 0.75rem; background-color: #f9fafb;">
                                    <div class="font-bold text-sm text-gray-800" style="font-weight: 700; font-size: 0.875rem; color: #1f2937;">Mode of Payment</div>
                                </div>
                                <div class="border-b-2 border-black p-3" style="border-bottom: 2px solid black; padding: 0.75rem;">
                                    <div class="text-center font-semibold" data-field="mode" style="text-align: center; font-weight: 600;">${expense.mode || '-'}</div>
                                </div>
                            </div>
                            
                            <!-- Expanded Description Section (Extra Large) -->
                            <div class="mt-6 border-2 border-black" style="margin-top: 1.5rem; border: 2px solid black;">
                                <div class="bg-gray-50 p-3 border-b-2 border-black flex justify-between items-center" style="background-color: #f9fafb; padding: 0.75rem; border-bottom: 2px solid black; display: flex; justify-content: space-between; align-items: center;">
                                    <div class="font-bold text-sm text-gray-800" style="font-weight: 700; font-size: 0.875rem; color: #1f2937;">Description</div>
                                </div>
                                <div class="p-8 min-h-[400px]" style="padding: 2rem; min-height: 400px;">
                                    <div id="descriptionDisplay_${index}" class="text-gray-700 text-base leading-relaxed" data-field="description" style="color: #374151; font-size: 1rem; line-height: 1.625;">${expense.description || 'No description provided'}</div>
                                </div>
                            </div>
                            
                            <!-- Total Amount Section (Direct - No separate Amount field) -->
                            <div class="mt-6 border-4 border-double border-black" style="margin-top: 1.5rem; border: 4px double black;">
                                <div class="grid grid-cols-2 gap-0" style="display: grid; grid-template-columns: repeat(2, 1fr); gap: 0;">
                                    <div class="bg-gray-100 p-4 border-r-4 border-double border-black" style="background-color: #f3f4f6; padding: 1rem; border-right: 4px double black;">
                                        <div class="font-bold text-xl text-gray-900" style="font-weight: 700; font-size: 1.25rem; color: #111827;">TOTAL AMOUNT</div>
                                    </div>
                                    <div class="p-4 text-right bg-yellow-50" style="padding: 1rem; text-align: right; background-color: #fefce8;">
                                        <div class="text-3xl font-bold text-gray-900" data-field="amount" id="totalAmountDisplay_${index}" style="font-size: 1.875rem; font-weight: 700; color: #111827;">${formatCurrency(amount)}</div>
                                    </div>
                                </div>
                            </div>
                            
                            <!-- Footer Signatures (Moved Lower) -->
                            <div class="mt-16 grid grid-cols-2 gap-8" style="margin-top: 4rem; display: grid; grid-template-columns: repeat(2, 1fr); gap: 2rem;">
                                <div class="border-t-4 border-black pt-12" style="border-top: 4px solid black; padding-top: 3rem;">
                                    <div class="text-center" style="text-align: center;">
                                        <div class="font-bold text-gray-800 mb-4" style="font-weight: 700; color: #1f2937; margin-bottom: 1rem;">Approved By</div>
                                        <div class="h-12 border-b-2 border-gray-400" style="height: 3rem; border-bottom: 2px solid #9ca3af;"></div>
                                    </div>
                                </div>
                                <div class="border-t-4 border-black pt-12" style="border-top: 4px solid black; padding-top: 3rem;">
                                    <div class="text-center" style="text-align: center;">
                                        <div class="font-bold text-gray-800 mb-4" style="font-weight: 700; color: #1f2937; margin-bottom: 1rem;">Receiver Signature</div>
                                        <div class="h-12 border-b-2 border-gray-400" style="height: 3rem; border-bottom: 2px solid #9ca3af;"></div>
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        `;
    });
    
    bulkHTML += '</div>';
    return bulkHTML;
}

// Print bulk vouchers
function printBulkVouchers() {
    if (selectedVouchers.size === 0) {
        showToast('Please select at least one voucher to print', 'error');
        return;
    }
    
    const selectedExpenses = expenseData.filter(expense => selectedVouchers.has(expense.srNo));
    
    // Open in new window for better printing control
    const printWindow = window.open('', '_blank', 'width=800,height=600');
    
    // Generate complete HTML document with inline styles
    const printHTML = generatePrintableBulkVoucherHTML(selectedExpenses);
    
    printWindow.document.write(`
        <!DOCTYPE html>
        <html>
        <head>
            <title>Bulk Expense Vouchers - ${selectedVouchers.size} Vouchers</title>
            <style>
                @page { margin: 0.5in; }
                body { 
                    font-family: Arial, sans-serif; 
                    margin: 20px; 
                    color: #333;
                    background: white;
                }
                .bulk-voucher-container { max-width: 100%; }
                .voucher-page { 
                    margin-bottom: 40px; 
                    page-break-after: always;
                    background: white;
                    position: relative;
                }
                .voucher-page:last-child { page-break-after: auto; }
                .print-area { padding: 20px 0; }
                .bg-white { background-color: white; }
                .relative { position: relative; }
                .overflow-hidden { overflow: hidden; }
                .shadow-2xl { box-shadow: 0 25px 50px -12px rgba(0, 0, 0, 0.25); }
                .border-4 { border: 4px solid black; }
                .border-black { border-color: black; }
                .absolute { position: absolute; }
                .inset-0 { top: 0; right: 0; bottom: 0; left: 0; }
                .opacity-5 { opacity: 0.05; }
                .pointer-events-none { pointer-events: none; }
                .text-9xl { font-size: 9rem; }
                .font-bold { font-weight: 700; }
                .text-blue-900 { color: #1e3a8a; }
                .transform { transform: translate(-50%, -50%) rotate(45deg); }
                .z-10 { z-index: 10; }
                .text-center { text-align: center; }
                .py-6 { padding-top: 1.5rem; padding-bottom: 1.5rem; }
                .border-b-4 { border-bottom: 4px solid black; }
                .font-playfair { font-family: 'Playfair Display', serif; }
                .text-3xl { font-size: 1.875rem; }
                .mb-2 { margin-bottom: 0.5rem; }
                .text-gray-900 { color: #111827; }
                .mb-4 { margin-bottom: 1rem; }
                .text-black { color: black; }
                .inline-block { display: inline-block; }
                .text-lg { font-size: 1.125rem; }
                .p-8 { padding: 2rem; }
                .grid { display: grid; }
                .grid-cols-4 { grid-template-columns: repeat(4, 1fr); }
                .gap-0 { gap: 0; }
                .border-2 { border: 2px solid black; }
                .border-r-2 { border-right: 2px solid black; }
                .border-b-2 { border-bottom: 2px solid black; }
                .p-3 { padding: 0.75rem; }
                .bg-gray-50 { background-color: #f9fafb; }
                .text-sm { font-size: 0.875rem; }
                .text-gray-800 { color: #1f2937; }
                .font-semibold { font-weight: 600; }
                .mt-6 { margin-top: 1.5rem; }
                .flex { display: flex; }
                .justify-between { justify-content: space-between; }
                .items-center { align-items: center; }
                .min-h-400 { min-height: 400px; }
                .text-gray-700 { color: #374151; }
                .text-base { font-size: 1rem; }
                .leading-relaxed { line-height: 1.625; }
                .border-double { border-style: double; }
                .grid-cols-2 { grid-template-columns: repeat(2, 1fr); }
                .bg-gray-100 { background-color: #f3f4f6; }
                .p-4 { padding: 1rem; }
                .text-right { text-align: right; }
                .bg-yellow-50 { background-color: #fefce8; }
                .mt-16 { margin-top: 4rem; }
                .gap-8 { gap: 2rem; }
                .border-t-4 { border-top: 4px solid black; }
                .pt-12 { padding-top: 3rem; }
                .h-12 { height: 3rem; }
                .border-gray-400 { border-color: #9ca3af; }
                @media print { 
                    body { margin: 10px; }
                    .voucher-page { page-break-after: always; }
                    .voucher-page:last-child { page-break-after: auto; }
                }
                .no-print { display: none !important; }
            </style>
        </head>
        <body>
            ${printHTML}
        </body>
        </html>
    `);
    
    printWindow.document.close();
    printWindow.focus();
    
    // Wait for content to load then print
    setTimeout(() => {
        printWindow.print();
        printWindow.close();
    }, 500);
    
    showToast(`Successfully printed ${selectedVouchers.size} vouchers!`, 'success');
    elements.bulkVoucherModal.classList.remove('active');
}

// Export bulk vouchers as PDF
function exportBulkVouchersAsPDF() {
    if (selectedVouchers.size === 0) {
        showToast('Please select at least one voucher to export', 'error');
        return;
    }
    
    const selectedExpenses = expenseData.filter(expense => selectedVouchers.has(expense.srNo));
    
    // Open in new window for PDF printing
    const printWindow = window.open('', '_blank', 'width=800,height=600');
    
    // Generate complete HTML document with inline styles
    const printHTML = generatePrintableBulkVoucherHTML(selectedExpenses);
    
    printWindow.document.write(`
        <!DOCTYPE html>
        <html>
        <head>
            <title>Bulk Expense Vouchers - ${selectedVouchers.size} Vouchers</title>
            <style>
                @page { margin: 0.5in; }
                body { 
                    font-family: Arial, sans-serif; 
                    margin: 20px; 
                    color: #333;
                    background: white;
                }
                .bulk-voucher-container { max-width: 100%; }
                .voucher-page { 
                    margin-bottom: 40px; 
                    page-break-after: always;
                    background: white;
                    position: relative;
                }
                .voucher-page:last-child { page-break-after: auto; }
                .print-area { padding: 20px 0; }
                .bg-white { background-color: white; }
                .relative { position: relative; }
                .overflow-hidden { overflow: hidden; }
                .shadow-2xl { box-shadow: 0 25px 50px -12px rgba(0, 0, 0, 0.25); }
                .border-4 { border: 4px solid black; }
                .border-black { border-color: black; }
                .absolute { position: absolute; }
                .inset-0 { top: 0; right: 0; bottom: 0; left: 0; }
                .opacity-5 { opacity: 0.05; }
                .pointer-events-none { pointer-events: none; }
                .text-9xl { font-size: 9rem; }
                .font-bold { font-weight: 700; }
                .text-blue-900 { color: #1e3a8a; }
                .transform { transform: translate(-50%, -50%) rotate(45deg); }
                .z-10 { z-index: 10; }
                .text-center { text-align: center; }
                .py-6 { padding-top: 1.5rem; padding-bottom: 1.5rem; }
                .border-b-4 { border-bottom: 4px solid black; }
                .font-playfair { font-family: 'Playfair Display', serif; }
                .text-3xl { font-size: 1.875rem; }
                .mb-2 { margin-bottom: 0.5rem; }
                .text-gray-900 { color: #111827; }
                .mb-4 { margin-bottom: 1rem; }
                .text-black { color: black; }
                .inline-block { display: inline-block; }
                .text-lg { font-size: 1.125rem; }
                .p-8 { padding: 2rem; }
                .grid { display: grid; }
                .grid-cols-4 { grid-template-columns: repeat(4, 1fr); }
                .gap-0 { gap: 0; }
                .border-2 { border: 2px solid black; }
                .border-r-2 { border-right: 2px solid black; }
                .border-b-2 { border-bottom: 2px solid black; }
                .p-3 { padding: 0.75rem; }
                .bg-gray-50 { background-color: #f9fafb; }
                .text-sm { font-size: 0.875rem; }
                .text-gray-800 { color: #1f2937; }
                .font-semibold { font-weight: 600; }
                .mt-6 { margin-top: 1.5rem; }
                .flex { display: flex; }
                .justify-between { justify-content: space-between; }
                .items-center { align-items: center; }
                .min-h-400 { min-height: 400px; }
                .text-gray-700 { color: #374151; }
                .text-base { font-size: 1rem; }
                .leading-relaxed { line-height: 1.625; }
                .border-double { border-style: double; }
                .grid-cols-2 { grid-template-columns: repeat(2, 1fr); }
                .bg-gray-100 { background-color: #f3f4f6; }
                .p-4 { padding: 1rem; }
                .text-right { text-align: right; }
                .bg-yellow-50 { background-color: #fefce8; }
                .mt-16 { margin-top: 4rem; }
                .gap-8 { gap: 2rem; }
                .border-t-4 { border-top: 4px solid black; }
                .pt-12 { padding-top: 3rem; }
                .h-12 { height: 3rem; }
                .border-gray-400 { border-color: #9ca3af; }
                @media print { 
                    body { margin: 10px; }
                    .voucher-page { page-break-after: always; }
                    .voucher-page:last-child { page-break-after: auto; }
                }
                .no-print { display: none !important; }
            </style>
        </head>
        <body>
            <div style="text-align: center; margin-bottom: 30px; padding-bottom: 20px; border-bottom: 2px solid #ddd;">
                <h1 style="color: #1e40af; margin-bottom: 10px;">Bulk Expense Vouchers</h1>
                <p>Total Vouchers: ${selectedVouchers.size}</p>
                <p>Generated on: ${new Date().toLocaleDateString()}</p>
            </div>
            ${printHTML}
        </body>
        </html>
    `);
    
    printWindow.document.close();
    printWindow.focus();
    
    showToast(`Opened ${selectedVouchers.size} vouchers in new window for PDF export`, 'success');
    elements.bulkVoucherModal.classList.remove('active');
}
