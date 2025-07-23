# Volume Sets Reverse Search Tool

A web-based tool for reverse searching volume sets from component ISBNs in DespatchReport Excel files. Upload your DespatchReport file and automatically find which volume sets your component ISBNs belong to, organized by order numbers with full order tracking.

## Features

- **Excel File Processing**: Automatically extracts component ISBNs from DespatchReport.xlsx files
- **Reverse Search**: Finds which volume sets contain your component ISBNs
- **Order Number Integration**: Associates each ISBN with both Order ID and Order No. (7000 series) from the DespatchReport
- **Multi-Order Support**: Handles files containing multiple orders with different Order Nos
- **Smart Grouping**: Groups results by volume set and order number for easy organization
- **PDF Export**: Generate professional PDF reports with detailed results including order information
- **Real-time Processing**: Instant results upon file upload with progress indicators

## How It Works

1. **Upload**: Select your DespatchReport Excel file (.xlsx or .xls)
2. **Automatic Extraction**: The tool scans column L for "VSet" entries and extracts ISBNs from column F
3. **Order Mapping**: Associates each ISBN with:
   - Order number from column D (e.g., 497833)
   - Order No. from column B (e.g., 7000770543)
4. **Volume Set Search**: Finds which volume sets contain each component ISBN
5. **Results Display**: Shows grouped results with both order numbers and complete set information
6. **PDF Export**: Generate and download a comprehensive report with order tracking

## File Format Requirements

The tool is designed to work with DespatchReport Excel files with the following structure:

- **Column B**: Order No. (7000 series numbers, e.g., 7000770543)
- **Column D**: Order numbers (e.g., 497833, 497588)
- **Column F**: Component ISBNs (13-digit ISBNs)
- **Column L**: Volume Set indicator (must contain "VSet")

### Example DespatchReport Structure:
```
Row | A                    | B          | C | D      | E   | F             | ... | L     |
----|---------------------|------------|---|--------|-----|---------------|-----|-------|
38  | Taylor & Francis UK | 7000770543 |   | 497833 | DPD |               | ... |       |
39  |                     |            |   |        |     | 9781032743721 | ... | VSet  |
40  |                     |            |   |        |     | 9781138676688 | ... | VSet  |
```

### Multi-Order File Support:
The tool handles files with multiple orders, where each section has its own Order No.:
- Order 497833 → Order No. 7000770543
- Order 497588 → Order No. 7000769173
- Order 497720 → Order No. 7000769977
- etc.

## Setup Instructions

### Prerequisites
- Modern web browser (Chrome, Firefox, Safari, Edge)
- Web server (local or remote) to serve the files

### Installation

1. **Download Files**: Get all the project files:
   - `index.html`
   - `script.js`
   - `vol_sets.json` (your volume sets database)

2. **Data Setup**: Ensure your `vol_sets.json` file contains volume set data in the following format:
   ```json
   [
     {
       "set_isbn": "9781234567890",
       "set_isbn_title": "Complete Works Series",
       "volume_count": 3,
       "volume_titles": "Volume 1 | Volume 2 | Volume 3",
       "associated_volumes": "9781111111111, 9782222222222, 9783333333333"
     }
   ]
   ```

3. **Web Server**: Serve the files from a web server:
   
   **Option 1 - Python (if installed):**
   ```bash
   # Python 3
   python -m http.server 8000
   
   # Python 2
   python -m SimpleHTTPServer 8000
   ```
   
   **Option 2 - Node.js (if installed):**
   ```bash
   npx serve .
   ```
   
   **Option 3 - PHP (if installed):**
   ```bash
   php -S localhost:8000
   ```

4. **Access**: Open your browser and navigate to `http://localhost:8000`

### File Structure
```
volume-sets-reverse-search/
├── index.html          # Main application interface
├── script.js           # Application logic and PDF export
├── vol_sets.json       # Volume sets database
├── README.md           # This file
└── favicon files       # Optional: favicon-32x32.png, apple-touch-icon.png
```

## Usage

### Step 1: Upload File
- Click "Choose File" and select your DespatchReport Excel file
- The file will be processed automatically upon selection
- Progress bar shows upload and processing status

### Step 2: Review Results
- **Summary**: View total ISBNs processed, found, and not found counts
- **Result Cards**: Each card shows:
  - Volume set information (title, ISBN, volume count)
  - Order badges showing both Order ID and Order No. (7000 series)
  - Found component ISBNs (highlighted in yellow in volume tables)
  - Complete list of all volumes in the set with titles

### Step 3: Export PDF (Optional)
- Click "Export PDF Report" to generate a professional report
- PDF includes:
  - All found sets with order information
  - Detailed volume tables with highlighted searched components
  - Not-found ISBNs grouped by order
  - Professional formatting with page numbers
- File is automatically downloaded with timestamp

## Result Display Details

### Found Volume Sets
Each result card displays:
- **Header**: "Found X component(s) in set"
- **Order Information**: 
  - Blue badge: "Order: 497833"
  - Gray badge: "Order No: 7000770543"
- **Set Details**:
  - Set title and ISBN
  - Volume count
  - Complete volume table with searched ISBNs highlighted
  
### Not Found ISBNs
- Grouped by order information
- Shows both Order ID and Order No.
- Lists all ISBNs that couldn't be matched to volume sets

## Technical Details

### Dependencies
The tool uses the following JavaScript libraries (loaded from CDN):
- **Bootstrap 5.3.2**: UI framework
- **Font Awesome 6.4.0**: Icons
- **SheetJS (XLSX) 0.18.5**: Excel file parsing
- **jsPDF 2.5.1**: PDF generation
- **jsPDF AutoTable 3.5.31**: PDF table formatting

### Browser Compatibility
- Chrome 60+
- Firefox 55+
- Safari 12+
- Edge 79+

### Data Processing
- Validates ISBN-13 checksums using mathematical verification
- Handles both string and numeric ISBN formats in Excel
- Groups results by both volume set ISBN and order information
- Maintains order number associations (both Order ID and Order No.) throughout processing
- Supports files with multiple orders and varying Order Nos

### Order Number Extraction
- **Order ID**: Extracted from column D (e.g., 497833, 497588)
- **Order No.**: Extracted from column B (7000 series, e.g., 7000770543)
- **Association**: Each ISBN is linked to the most recent order header above it
- **Multi-Order**: Handles files with dozens of different orders

## Troubleshooting

### Common Issues

**"No ISBN numbers found in the uploaded file"**
- Ensure column L contains "VSet" text exactly
- Verify column F contains valid 13-digit ISBNs
- Check that columns B and D contain order information
- Ensure the Excel file isn't corrupted or password-protected

**"Error loading data"**
- Ensure `vol_sets.json` file is present and properly formatted
- Check browser console (F12) for detailed error messages
- Verify web server is serving JSON files with correct MIME type

**PDF export not working**
- Ensure you have search results loaded first
- Check browser console for JavaScript errors
- Try refreshing the page and re-uploading your file
- Verify your browser allows file downloads

**Missing order information**
- Check that column B contains Order No. (7000 series)
- Verify column D contains Order numbers
- Ensure order headers appear before VSet rows in the file

### File Format Issues
- Only Excel files (.xlsx, .xls) are supported
- CSV files are not supported for this tool
- Ensure your DespatchReport follows the expected column structure
- Order information must be in the correct columns (B and D)
- VSet indicators must be in column L, ISBNs in column F

### Performance Notes
- Large files (1000+ orders) may take a few seconds to process
- PDF generation time depends on the number of results
- Browser may briefly freeze during large file processing (this is normal)

## License

This project is provided as-is for internal use. Modify and distribute as needed for your organization.

## Support

For technical issues or feature requests:
1. Check the browser console (F12) for error messages
2. Ensure all file format requirements are met
3. Verify your vol_sets.json database is properly formatted
4. Test with a smaller sample file if experiencing performance issues