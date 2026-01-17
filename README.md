# CUP Jackets Picklist Generator

A web application for generating PDF picklists and XML files for Cambridge University Press (CUP) book jacket jobs. The application reads Excel files containing job data, matches ISBNs against a database, filters for jobs with book jackets, and generates downloadable outputs with Code 128 barcodes.

## Features

- **Excel File Upload**: Upload Excel files (.xlsx, .xls) containing job information
- **Automatic ISBN Matching**: Matches ISBN codes from Excel against the CUP database (cup_data.json)
- **Jacket Filtering with Double Validation**: Automatically identifies and filters jobs using dual validation:
  - Excel must have Jacket Y/N = true
  - JSON database must have has_jacket = true
  - Both conditions must be met for a job to be processed
- **PDF Picklist Generation**: Creates a formatted PDF picklist with:
  - Job summary information (job number, order date, total jobs, total quantity)
  - Detailed table with Order No, ISBN, Title, Qty, Trim Size, Treatment
  - Code 128 barcodes for each ISBN for easy scanning
  - Automatic pagination with repeated headers
- **XML Export**: Generates individual XML files for each jacket job packaged in a ZIP file
- **Results Dashboard**: Shows total jobs, jacket jobs, and matched jobs with detailed table preview
- **Clean Title Processing**: Automatically removes "Cover" suffix from book titles
- **Responsive Design**: Works on desktop and mobile devices

## Setup

1. Ensure all files are in the same directory:
   - `index.html`
   - `styles.css`
   - `script.js`
   - `cup_data.json`

2. Open `index.html` in a modern web browser (Chrome, Firefox, Safari, Edge)

3. No server or installation required - runs entirely in the browser

## Required Files

### cup_data.json
The JSON database file must contain an array of book records with the following structure:

```json
[
  {
    "isbn": "9781108471060",
    "book_description": "Plays 1676–1678 Cover",
    "customer": "CUP",
    "binding_type": "Cased",
    "trim_height": "229",
    "trim_width": "152",
    "pagination": null,
    "spine_size": "49",
    "weight_gsm": "115",
    "pdf_url": "POD_Files/9781108471060_jkt.pdf",
    "stock_description": "Cover Paper",
    "cover_media_treatment": "Matt",
    "has_jacket": true
  }
]
```

### Excel File Format
The Excel file should contain the following columns:
- **Customer**: Customer name
- **Pace Job No**: Job number
- **Customer Order No.**: Order number
- **Title**: Book title
- **Code**: ISBN (matched against cup_data.json)
- **Qty**: Quantity required
- **Order Date**: Order date
- **Bind Method**: Binding method (e.g., Cased)
- **H**: Height dimension
- **W**: Width dimension
- **Spine**: Spine size
- **Jacket Y/N**: Jacket indicator (true/false, yes/no, or 1/0)
- **Text Paper**: Paper type
- **Job Status**: Current job status
- **Note**: Additional notes

## Usage

### 1. Upload Excel File
- Click "Choose file" or drag and drop an Excel file
- The application automatically processes the file and matches ISBNs

### 2. View Results
- **Summary Statistics**: 
  - Total Jobs: All jobs in the Excel file
  - Jobs with Jackets: Jobs where Jacket Y/N = true
  - Jobs Matched in Database: Jobs successfully matched in cup_data.json
- **Detailed Table**: Shows all matched jacket jobs with full specifications

### 3. Download PDF Picklist
- Click "Download PDF Picklist" to generate a formatted PDF
- PDF includes:
  - Header with job number, order date, generation timestamp
  - Total jacket jobs count and total quantity required
  - Table with all jacket jobs
  - Code 128 barcode for each ISBN (scannable)
  - Page numbers on each page
- File name format: `Jacket_Picklist_[JobNumber]_[Date].pdf`

### 4. Download XML Files
- Click "Download XML (ZIP)" to generate XML files
- Creates a ZIP file containing individual XML files for each jacket job
- Each XML includes complete job and specification details
- File name format: `Jacket_XML_[JobNumber]_[Date].zip`
- Individual XML files named: `[ISBN]_jacket.xml`

### 5. Clear Data
- Click "Clear All" to reset the application and start over

## PDF Picklist Details

### Header Information
- Job Number
- Order Date
- Generation timestamp
- Total Jacket Jobs (count)
- Total Jackets Required (sum of all quantities)

### Table Columns
1. **Order No**: Customer order number for tracking
2. **ISBN**: Book ISBN-13 number
3. **Title**: Book title (automatically cleaned - "Cover" suffix removed)
4. **Qty**: Quantity of jackets required
5. **Trim Size**: Book dimensions (Height × Width in mm)
6. **Treatment**: Cover finish (Matt/Gloss)
7. **Barcode**: Code 128 barcode of ISBN for scanning

### Features
- Automatic page breaks with header repetition
- Compact layout fitting ~25 jobs per page
- Professional formatting suitable for production use
- Scannable barcodes for inventory management

## XML Output Structure

Each XML file contains:

```xml
<?xml version="1.0" encoding="UTF-8"?>
<BookJacket>
    <JobInfo>
        <ISBN>9781108471060</ISBN>
        <PaceJobNo>511527</PaceJobNo>
        <CustomerOrderNo>5000069286</CustomerOrderNo>
        <OrderDate>27/11/2025</OrderDate>
        <Quantity>1</Quantity>
    </JobInfo>
    <BookDetails>
        <Title>Plays 1676–1678</Title>
        <Customer>CUP</Customer>
        <BindingType>Cased</BindingType>
    </BookDetails>
    <Specifications>
        <TrimHeight unit="mm">229</TrimHeight>
        <TrimWidth unit="mm">152</TrimWidth>
        <SpineSize unit="mm">49</SpineSize>
        <Pagination></Pagination>
    </Specifications>
    <Materials>
        <StockDescription>Cover Paper</StockDescription>
        <WeightGSM>115</WeightGSM>
        <CoverMediaTreatment>Matt</CoverMediaTreatment>
    </Materials>
    <Files>
        <PDFUrl>POD_Files/9781108471060_jkt.pdf</PDFUrl>
    </Files>
</BookJacket>
```

## Technical Details

### Technologies Used
- **HTML5**: Structure and layout
- **CSS3**: Styling with modern design patterns
- **Vanilla JavaScript**: Application logic (ES6+)
- **SheetJS (xlsx)**: Excel file parsing
- **jsPDF**: PDF generation
- **JSZip**: ZIP file creation for XML exports
- **JsBarcode**: Code 128 barcode generation

### Browser Compatibility
- Chrome/Edge (recommended)
- Firefox
- Safari
- Any modern browser with ES6+ support

### Performance
- Excel files: Optimized for up to 1000 rows
- PDF generation: Handles 100+ jacket jobs efficiently
- All processing done client-side (no server required)

## Data Processing

### Double Validation for Jacket Jobs
The application uses a **two-step validation process** to ensure data accuracy:

1. **Excel Validation**: Checks if "Jacket Y/N" column = true/yes/1
2. **Database Validation**: Verifies `has_jacket: true` in cup_data.json

**Both conditions must be met** for a job to be processed. This prevents:
- Processing non-jacket jobs accidentally marked in Excel
- Data mismatches between Excel and the master database
- Invalid jacket specifications being sent to production

**Console Logging**: The browser console will show:
- Number of jobs marked as jackets in Excel
- Number matched in database with `has_jacket = true`
- Warning list of any jobs filtered out due to database mismatch

**Example Scenario:**
- Excel shows ISBN 9780521127295 with "Jacket Y/N" = true
- Database shows `"has_jacket": false` for this ISBN
- Result: Job is **filtered out** and will not appear in results
- Console warning: "1 job(s) marked as jackets in Excel but has_jacket = false in database"

### Title Cleaning
The application automatically cleans book titles by:
- Removing " Cover" suffix (case insensitive)
- Trimming whitespace
- Preserving original formatting otherwise

**Example:**
- Input: "Plays 1676–1678 Cover"
- Output: "Plays 1676–1678"

### ISBN Matching
- Exact match required between Excel "Code" column and cup_data.json "isbn" field
- ISBNs should be 13-digit format (e.g., 9781108471060)
- No formatting (hyphens removed)

### Jacket Filtering
Jobs are identified as jacket jobs when "Jacket Y/N" column contains:
- "true" (case insensitive)
- "yes" (case insensitive)
- "1"

## Troubleshooting

### "Failed to load database" error
- Ensure `cup_data.json` is in the same directory as `index.html`
- Check that the JSON file is valid
- Verify file permissions allow reading

### Excel file not processing
- Verify file is .xlsx or .xls format
- Check that required columns exist
- Ensure "Code" column contains valid ISBNs
- Verify "Jacket Y/N" column has proper values

### No jobs matched
- Check ISBNs in Excel match exactly those in cup_data.json
- Verify "Jacket Y/N" column is marked correctly
- Ensure ISBNs don't have extra spaces or formatting

### PDF not downloading
- Check browser pop-up blocker settings
- Ensure JavaScript is enabled
- Try a different browser
- Check browser console for errors (F12)

### Barcodes not appearing in PDF
- Ensure JsBarcode library is loaded
- Check browser console for errors
- Verify ISBNs are valid (numeric only)

## Customization

### Styling
Edit `styles.css` to customize:
- Colors and theme
- Fonts and typography
- Layout and spacing
- Button styles
- Responsive breakpoints

### PDF Format
Edit the `generatePdfPicklist()` function in `script.js` to modify:
- Page layout and margins
- Header/footer content
- Column selection and widths
- Font sizes and styling
- Barcode size and position

### XML Structure
Edit the `createJobXml()` function in `script.js` to modify:
- XML element structure
- Data fields included
- Element naming conventions
- Data formatting

## File Structure

```
project/
├── index.html          # Main HTML file
├── styles.css          # Application styles
├── script.js           # Application logic
├── cup_data.json       # CUP book database
└── README.md          # This file
```

## Security Notes

- All processing happens client-side (browser)
- No data is sent to external servers
- Files are processed locally in browser memory
- Safe to use with confidential data

## Known Limitations

- Requires modern browser with ES6+ support
- File size limited by browser memory (typically 100MB+)
- PDF generation may be slow for 500+ jobs
- Excel files must follow expected format

## Support

For issues or questions:
1. Check browser console for error messages (F12)
2. Verify all files are present and properly formatted
3. Try with a smaller test file first
4. Check that cup_data.json contains matching ISBNs

## Version History

### Current Version
- Code 128 barcode generation for ISBNs
- Automatic title cleaning (removes "Cover" suffix)
- ZIP packaging for XML downloads
- Enhanced PDF layout with Treatment column
- Responsive upload interface
- Clear All functionality

## License

Copyright © 2025. All rights reserved.

## Credits

Built for Cambridge University Press jacket production workflow.
