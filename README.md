# Expense Management System

A full-stack web application for managing expenses with Excel import functionality. Built with Node.js, Express, and vanilla JavaScript.

## Features

- **Excel Import**: Upload .xlsx or .xls files to import expense data
- **Modern Dashboard**: Dark-themed admin interface with statistics cards
- **Data Table**: Responsive table with status badges and action buttons
- **Status Management**: Visual badges for Paid/Pending status
- **Action Buttons**: Save, Voucher generation, and Delete functionality
- **Printable Vouchers**: Generate and print expense vouchers
- **Responsive Design**: Works seamlessly on desktop and mobile devices
- **Real-time Statistics**: Live updates of total amounts and record counts

## Tech Stack

- **Backend**: Node.js with Express
- **Frontend**: HTML5, CSS3, Vanilla JavaScript
- **File Processing**: Multer for uploads, XLSX for Excel parsing
- **Styling**: Modern CSS with gradients and animations
- **Icons**: Font Awesome

## Project Structure

```
expense-management-system/
|-- server.js              # Express server and API endpoints
|-- package.json           # Dependencies and scripts
|-- uploads/               # Temporary file storage
|-- public/
|   |-- index.html         # Main HTML file
|   |-- style.css          # Styles and responsive design
|   |-- script.js          # Frontend JavaScript
|-- README.md              # This file
```

## Installation

1. **Install Dependencies**
   ```bash
   npm install
   ```

2. **Start the Server**
   ```bash
   npm start
   ```
   Or
   ```bash
   node server.js
   ```

3. **Open the Application**
   Navigate to `http://localhost:3000` in your browser

## Usage

### Importing Excel Files

1. Click the **"Import Excel"** button
2. Select an Excel file (.xlsx or .xls)
3. The file will be processed and data will appear in the table

### Expected Excel Format

Your Excel file should have the following headers:
- `Date` (any date format)
- `Given To` (recipient name)
- `Amount` (numeric value)
- `Mode` (payment method)
- `Description` (expense details)
- `Fund` (fund category)
- `Status` (Paid/Pending)

### Table Features

- **Status Badges**: 
  - Green for "Paid"
  - Orange for "Pending"
- **Action Buttons**:
  - **Save**: Logs expense data to console
  - **Voucher**: Opens printable expense voucher
  - **Delete**: Removes the record from the table

### Keyboard Shortcuts

- `Ctrl/Cmd + I`: Open import dialog
- `Ctrl/Cmd + E`: Export data to CSV
- `Shift + Delete`: Clear all data
- `Escape`: Close modal dialogs

## API Endpoints

### POST /upload-excel
Uploads and processes Excel files.

**Request**: `multipart/form-data` with file field named `excelFile`

**Response**:
```json
{
  "success": true,
  "data": [
    {
      "srNo": 1,
      "date": "2024-01-15",
      "givenTo": "John Doe",
      "amount": "150.00",
      "mode": "Cash",
      "description": "Office supplies",
      "fund": "General",
      "status": "Paid"
    }
  ],
  "message": "Successfully imported 1 records"
}
```

### GET /health
Health check endpoint.

**Response**:
```json
{
  "status": "Server is running",
  "timestamp": "2024-01-15T10:30:00.000Z"
}
```

## File Processing

- Uploaded files are temporarily stored in the `uploads/` directory
- Files are automatically deleted after processing
- Supports both .xlsx and .xls file formats
- Maximum file size: 5MB

## Error Handling

- Invalid file types are rejected
- Network errors show user-friendly messages
- Empty files are handled gracefully
- Malformed Excel data is processed with fallbacks

## Browser Compatibility

- Chrome 60+
- Firefox 55+
- Safari 12+
- Edge 79+

## Development

To run in development mode:
```bash
npm run dev
```

The server will start on `http://localhost:3000` and serve static files from the `public/` directory.

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly
5. Submit a pull request

## License

MIT License - feel free to use this project for personal or commercial purposes.

## Troubleshooting

### Common Issues

1. **File Upload Fails**
   - Check file format (.xlsx or .xls only)
   - Ensure file size is under 5MB
   - Verify Excel headers match expected format

2. **Server Won't Start**
   - Ensure all dependencies are installed: `npm install`
   - Check if port 3000 is available
   - Verify Node.js version (14+ recommended)

3. **Data Not Displaying**
   - Check browser console for errors
   - Verify Excel file has correct headers
   - Ensure data is not empty

### Support

For issues and questions, please check the browser console for error messages and verify all setup steps have been completed correctly.
