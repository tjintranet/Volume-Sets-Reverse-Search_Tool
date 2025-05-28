# Volume Sets Reverse Search Tool

A web-based tool for reverse searching volume sets from component ISBNs in DespatchReport Excel files. Upload your DespatchReport file and automatically find which volume sets your component ISBNs belong to, organized by order numbers.

## Features

- **Excel File Processing**: Automatically extracts component ISBNs from DespatchReport.xlsx files
- **Reverse Search**: Finds which volume sets contain your component ISBNs
- **Order Number Integration**: Associates each ISBN with its order number from the DespatchReport
- **Smart Grouping**: Groups results by volume set and order number
- **PDF Export**: Generate professional PDF reports with detailed results
- **Real-time Processing**: Instant results upon file upload

## How It Works

1. **Upload**: Select your DespatchReport Excel file (.xlsx or .xls)
2. **Automatic Extraction**: The tool scans column L for "VSet" entries and extracts ISBNs from column F
3. **Order Mapping**: Associates each ISBN with its order number from column D
4. **Volume Set Search**: Finds which volume sets contain each component ISBN
5. **Results Display**: Shows grouped results with order numbers and complete set information
6. **PDF Export**: Generate and download a comprehensive report

## File Format Requirements

The tool is designed to work with DespatchReport Excel files with the following structure:

- **Column D**: Order numbers (e.g., 492990, 493012)
- **Column F**: Component ISBNs (13-digit ISBNs)
- **Column L**: Volume Set indicator (must contain "VSet")

### Example DespatchReport Structure:
```
Row | A                    | B          | C | D      | E   | F             | ... | L     |
----|---------------------|------------|---|--------|-----|---------------|-----|-------|
6   | Taylor & Francis UK | 7000746762 |   | 492990 | DPD |               | ... |       |
7   |                     |            |   |        |     | 9781138124011 | ... | VSet  |
8   |                     |            |   |        |     | 9781138339279 | ... | VSet  |
```

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

### Step 2: Review Results
- **Summary**: View total ISBNs processed, found, and not found
- **Result Cards**: Each card shows:
  - Volume set information
  - Order number (blue badge)
  - Found component ISBNs (highlighted in tables)
  - Complete list of all volumes in the set

### Step 3: Export PDF (Optional)
- Click "Export PDF Report" to generate a professional report
- PDF includes all found sets, order numbers, and detailed volume tables
- File is automatically downloaded with timestamp

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
- Validates ISBN-13 checksums
- Handles both string and numeric ISBN formats in Excel
- Groups results by volume set and order number
- Maintains order number associations throughout processing

## Troubleshooting

### Common Issues

**"No ISBN numbers found in the uploaded file"**
- Ensure column L contains "VSet" text
- Verify column F contains valid 13-digit ISBNs
- Check that the Excel file isn't corrupted

**"Error loading data"**
- Ensure `vol_sets.json` file is present and properly formatted
- Check browser console for detailed error messages
- Verify web server is serving JSON files with correct MIME type

**PDF export not working**
- Ensure you have search results loaded
- Check browser console for JavaScript errors
- Try refreshing the page and re-uploading your file

### File Format Issues
- Only Excel files (.xlsx, .xls) are supported
- CSV files are not supported for this tool
- Ensure your DespatchReport follows the expected column structure

## License

This project is provided as-is for internal use. Modify and distribute as needed for your organization.

## Support

For technical issues or feature requests, check the browser console for error messages and ensure all file format requirements are met.