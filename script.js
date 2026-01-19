// Global variables
let cupData = [];
let processedJobs = [];
let excelData = [];

// DOM elements
const fileInput = document.getElementById('fileInput');
const clearBtn = document.getElementById('clearBtn');
const statusSection = document.getElementById('statusSection');
const statusMessage = document.getElementById('statusMessage');
const resultsSection = document.getElementById('resultsSection');
const errorSection = document.getElementById('errorSection');
const errorMessage = document.getElementById('errorMessage');
const resetBtn = document.getElementById('resetBtn');
const downloadPdfBtn = document.getElementById('downloadPdfBtn');
const downloadXmlBtn = document.getElementById('downloadXmlBtn');
const resultsBody = document.getElementById('resultsBody');
const totalJobs = document.getElementById('totalJobs');
const jacketJobs = document.getElementById('jacketJobs');
const matchedJobs = document.getElementById('matchedJobs');

// Initialize application
async function init() {
    await loadCupData();
    setupEventListeners();
}

// Load CUP data from JSON
async function loadCupData() {
    try {
        const response = await fetch('cup_data.json');
        if (!response.ok) {
            throw new Error('Failed to load CUP data');
        }
        cupData = await response.json();
        console.log(`Loaded ${cupData.length} records from CUP database`);
    } catch (error) {
        console.error('Error loading CUP data:', error);
        showError('Failed to load database. Please ensure cup_data.json is in the same directory.');
    }
}

// Setup event listeners
function setupEventListeners() {
    // File upload events
    fileInput.addEventListener('change', handleFileSelect);
    
    // Clear button
    clearBtn.addEventListener('click', clearAll);
    
    // Download buttons
    downloadPdfBtn.addEventListener('click', generatePdfPicklist);
    downloadXmlBtn.addEventListener('click', generateXmlFiles);
    
    // Reset button
    resetBtn.addEventListener('click', resetApplication);
}

// Handle file selection
function handleFileSelect(e) {
    const file = e.target.files[0];
    if (file) {
        handleFile(file);
    }
}

// Handle file
function handleFile(file) {
    // Validate file type
    const validTypes = ['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'application/vnd.ms-excel'];
    if (!validTypes.includes(file.type) && !file.name.match(/\.(xlsx|xls)$/i)) {
        showError('Please upload a valid Excel file (.xlsx or .xls)');
        return;
    }
    
    // Process file
    processExcelFile(file);
}

// Remove file
function removeFile() {
    fileInput.value = '';
    resultsSection.style.display = 'none';
    errorSection.style.display = 'none';
    excelData = [];
    processedJobs = [];
}

// Clear all data and reset the form
function clearAll() {
    fileInput.value = '';
    resultsSection.style.display = 'none';
    errorSection.style.display = 'none';
    statusSection.style.display = 'none';
    excelData = [];
    processedJobs = [];
}

// Process Excel file
async function processExcelFile(file) {
    showStatus('Reading Excel file...');
    
    try {
        const data = await readExcelFile(file);
        excelData = data;
        
        showStatus('Matching jobs with database...');
        await new Promise(resolve => setTimeout(resolve, 500)); // Small delay for UX
        
        processJobs();
        
    } catch (error) {
        console.error('Error processing file:', error);
        showError(`Error processing file: ${error.message}`);
    }
}

// Read Excel file
function readExcelFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                const jsonData = XLSX.utils.sheet_to_json(firstSheet, { raw: false });
                resolve(jsonData);
            } catch (error) {
                reject(error);
            }
        };
        
        reader.onerror = () => reject(new Error('Failed to read file'));
        reader.readAsArrayBuffer(file);
    });
}

// Process jobs
function processJobs() {
    // Filter for jobs with jackets
    const jacketJobsData = excelData.filter(row => {
        const hasJacket = row['Jacket Y/N'];
        // Handle both boolean true and string values
        if (typeof hasJacket === 'boolean') {
            return hasJacket === true;
        }
        return hasJacket && (hasJacket.toLowerCase() === 'true' || hasJacket.toLowerCase() === 'yes' || hasJacket === '1');
    });
    
    console.log(`Excel jobs with Jacket Y/N = true: ${jacketJobsData.length}`);
    
    // Match with CUP data
    processedJobs = jacketJobsData.map(row => {
        const isbn = row['ISBN'] || row['Code'] ? (row['ISBN'] || row['Code']).toString().trim() : '';
        const cupRecord = cupData.find(item => item.isbn === isbn);
        
        return {
            excelData: row,
            cupData: cupRecord,
            isbn: isbn
        };
    }).filter(job => job.cupData && job.cupData.has_jacket === true); // Only keep jobs that matched in database AND have has_jacket = true
    
    console.log(`Jobs matched in database with has_jacket = true: ${processedJobs.length}`);
    
    // Log ISBNs not found in database
    const notFoundInDb = jacketJobsData.filter(row => {
        const isbn = row['ISBN'] || row['Code'] ? (row['ISBN'] || row['Code']).toString().trim() : '';
        const cupRecord = cupData.find(item => item.isbn === isbn);
        return !cupRecord;
    });
    
    if (notFoundInDb.length > 0) {
        console.error(`${notFoundInDb.length} jacket job(s) NOT FOUND in database:`);
        notFoundInDb.forEach(row => {
            const isbn = row['ISBN'] || row['Code'] ? (row['ISBN'] || row['Code']).toString().trim() : '';
            console.error(`  - ISBN: ${isbn}, Title: ${row['Title']}`);
        });
        console.error('These ISBNs need to be added to cup_data.json');
    }
    
    // Log any jobs that were filtered out due to has_jacket = false
    const filteredOut = jacketJobsData.filter(row => {
        const isbn = row['ISBN'] || row['Code'] ? (row['ISBN'] || row['Code']).toString().trim() : '';
        const cupRecord = cupData.find(item => item.isbn === isbn);
        return cupRecord && cupRecord.has_jacket === false;
    });
    
    if (filteredOut.length > 0) {
        console.warn(`${filteredOut.length} job(s) marked as jackets in Excel but has_jacket = false in database:`);
        filteredOut.forEach(row => {
            console.warn(`  - ISBN: ${row['ISBN'] || row['Code']}, Title: ${row['Title']}`);
        });
    }
    
    // Update UI
    displayResults();
}

// Display results
function displayResults() {
    const totalJobsCount = excelData.length;
    const jacketJobsCount = excelData.filter(row => {
        const hasJacket = row['Jacket Y/N'];
        // Handle both boolean true and string values
        if (typeof hasJacket === 'boolean') {
            return hasJacket === true;
        }
        return hasJacket && (hasJacket.toLowerCase() === 'true' || hasJacket.toLowerCase() === 'yes' || hasJacket === '1');
    }).length;
    const matchedJobsCount = processedJobs.length;
    
    // Update summary
    totalJobs.textContent = totalJobsCount;
    jacketJobs.textContent = jacketJobsCount;
    matchedJobs.textContent = matchedJobsCount;
    
    // If no matches found, show helpful error message
    if (matchedJobsCount === 0 && jacketJobsCount > 0) {
        hideStatus();
        showError(`Found ${jacketJobsCount} jacket job(s) in Excel, but none matched in the database.\n\nPossible reasons:\n• ISBNs don't exist in cup_data.json\n• Database records have has_jacket = false\n\nCheck browser console (F12) for details.`);
        return;
    }
    
    if (matchedJobsCount === 0 && jacketJobsCount === 0) {
        hideStatus();
        showError('No jacket jobs found in Excel file. Make sure the "Jacket Y/N" column contains True values.');
        return;
    }
    
    // Populate table
    resultsBody.innerHTML = '';
    processedJobs.forEach(job => {
        const row = document.createElement('tr');
        const title = cleanTitle(job.cupData.book_description || job.excelData.Title || 'N/A');
        const jacketRoute = getJacketRoute(job.cupData.trim_height, job.cupData.trim_width);
        row.innerHTML = `
            <td>${job.cupData.isbn}</td>
            <td>${title}</td>
            <td>${job.cupData.customer || 'N/A'}</td>
            <td>${job.excelData.Qty || '1'}</td>
            <td>${job.cupData.trim_height || 'N/A'} × ${job.cupData.trim_width || 'N/A'} mm</td>
            <td>${job.cupData.spine_size || 'N/A'} mm</td>
            <td>${job.cupData.cover_media_treatment || 'N/A'}</td>
            <td><strong>${jacketRoute}</strong></td>
        `;
        resultsBody.appendChild(row);
    });
    
    // Show results
    hideStatus();
    resultsSection.style.display = 'block';
    
    // Enable download buttons if we have results
    downloadPdfBtn.disabled = processedJobs.length === 0;
    downloadXmlBtn.disabled = processedJobs.length === 0;
}

// Generate PDF Picklist
async function generatePdfPicklist() {
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({
        orientation: 'landscape',
        unit: 'mm',
        format: 'a4'
    });
    
    // Get job number from first row
    const jobNumber = excelData[0]['Pace Job No'] || 'Unknown';
    const orderDate = excelData[0]['Order Date'] || new Date().toLocaleDateString();
    
    // Calculate total quantity
    const totalQuantity = processedJobs.reduce((sum, job) => {
        const qty = parseInt(job.excelData.Qty) || 1;
        return sum + qty;
    }, 0);
    
    // Title
    doc.setFontSize(18);
    doc.setFont(undefined, 'bold');
    doc.text('CUP Book Jacket Picklist', 14, 20);
    
    // Job details
    doc.setFontSize(10);
    doc.setFont(undefined, 'normal');
    doc.text(`Job Number: ${jobNumber}`, 14, 30);
    doc.text(`Order Date: ${orderDate}`, 14, 36);
    doc.text(`Generated: ${new Date().toLocaleString()}`, 14, 42);
    doc.text(`Total Jacket Jobs: ${processedJobs.length}`, 14, 48);
    doc.text(`Total Jackets Required: ${totalQuantity}`, 14, 54);
    
    // Table headers (landscape orientation allows more room)
    let yPosition = 64;
    doc.setFontSize(9);
    doc.setFont(undefined, 'bold');
    doc.text('Order No', 14, yPosition);
    doc.text('ISBN', 42, yPosition);
    doc.text('Title', 75, yPosition);
    doc.text('Qty', 155, yPosition);
    doc.text('Trim Size', 167, yPosition);
    doc.text('Treatment', 195, yPosition);
    doc.text('Jacket Route', 220, yPosition);
    doc.text('Barcode', 250, yPosition);
    
    // Draw line under headers
    doc.line(14, yPosition + 2, 283, yPosition + 2);
    
    // Table content
    yPosition += 8;
    doc.setFont(undefined, 'normal');
    
    for (let i = 0; i < processedJobs.length; i++) {
        const job = processedJobs[i];
        
        // Check if we need a new page (landscape has more vertical space: ~185mm usable)
        if (yPosition > 185) {
            doc.addPage();
            yPosition = 20;
            
            // Repeat headers on new page
            doc.setFont(undefined, 'bold');
            doc.text('Order No', 14, yPosition);
            doc.text('ISBN', 42, yPosition);
            doc.text('Title', 75, yPosition);
            doc.text('Qty', 155, yPosition);
            doc.text('Trim Size', 167, yPosition);
            doc.text('Treatment', 195, yPosition);
            doc.text('Jacket Route', 220, yPosition);
            doc.text('Barcode', 250, yPosition);
            doc.line(14, yPosition + 2, 283, yPosition + 2);
            yPosition += 8;
            doc.setFont(undefined, 'normal');
        }
        
        const orderNo = job.excelData['Customer Order No.'] || 'N/A';
        const isbn = job.cupData.isbn || 'N/A';
        const rawTitle = job.cupData.book_description || job.excelData.Title || 'N/A';
        const cleanedTitle = cleanTitle(rawTitle);
        const title = truncateText(cleanedTitle, 50); // More room in landscape for title
        const qty = job.excelData.Qty || '1';
        const trimSize = `${job.cupData.trim_height}×${job.cupData.trim_width}`;
        const treatment = truncateText(job.cupData.cover_media_treatment || 'N/A', 12);
        const jacketRoute = getJacketRoute(job.cupData.trim_height, job.cupData.trim_width);
        
        doc.text(orderNo.toString(), 14, yPosition);
        doc.text(isbn, 42, yPosition);
        doc.text(title, 75, yPosition);
        doc.text(qty.toString(), 155, yPosition);
        doc.text(trimSize, 167, yPosition);
        doc.text(treatment, 195, yPosition);
        
        // Add Jacket Route with bold font
        doc.setFont(undefined, 'bold');
        doc.text(jacketRoute, 220, yPosition);
        doc.setFont(undefined, 'normal');
        
        // Generate and add barcode
        try {
            const barcodeDataUrl = await generateBarcode(isbn);
            if (barcodeDataUrl) {
                // Add barcode image to PDF (positioned in the barcode column)
                doc.addImage(barcodeDataUrl, 'PNG', 248, yPosition - 4, 25, 6);
            }
        } catch (error) {
            console.error('Error generating barcode for ISBN:', isbn, error);
        }
        
        yPosition += 10; // Increased spacing to accommodate barcode
    }
    
    // Footer (landscape A4 height is 210mm)
    const pageCount = doc.internal.getNumberOfPages();
    for (let i = 1; i <= pageCount; i++) {
        doc.setPage(i);
        doc.setFontSize(8);
        doc.text(`Page ${i} of ${pageCount}`, 14, 203);
    }
    
    // Save PDF
    doc.save(`Jacket_Picklist_${jobNumber}_${new Date().toISOString().split('T')[0]}.pdf`);
}

// Generate Code 128 barcode
function generateBarcode(text) {
    return new Promise((resolve, reject) => {
        try {
            // Create a temporary canvas element
            const canvas = document.createElement('canvas');
            
            // Generate the barcode on the canvas
            JsBarcode(canvas, text, {
                format: 'CODE128',
                width: 2,
                height: 40,
                displayValue: false, // Don't show text below barcode
                margin: 2
            });
            
            // Convert canvas to data URL
            const dataUrl = canvas.toDataURL('image/png');
            resolve(dataUrl);
        } catch (error) {
            console.error('Barcode generation error:', error);
            resolve(null); // Return null if barcode generation fails
        }
    });
}

// Generate XML files
async function generateXmlFiles() {
    const jobNumber = excelData[0]['Pace Job No'] || 'Unknown';
    
    // Create a new JSZip instance
    const zip = new JSZip();
    
    // Add each XML file to the zip
    processedJobs.forEach((job, index) => {
        const xml = createJobXml(job);
        const filename = `${job.cupData.isbn}_jacket.xml`;
        zip.file(filename, xml);
    });
    
    // Generate the zip file
    try {
        const content = await zip.generateAsync({type: 'blob'});
        
        // Create download link
        const url = URL.createObjectURL(content);
        const a = document.createElement('a');
        a.href = url;
        a.download = `Jacket_XML_${jobNumber}_${new Date().toISOString().split('T')[0]}.zip`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
        
        // Show confirmation
        alert(`Generated ZIP file with ${processedJobs.length} XML files`);
    } catch (error) {
        console.error('Error generating ZIP file:', error);
        alert('Error generating ZIP file. Please try again.');
    }
}

// Create XML for a job
function createJobXml(job) {
    const cupData = job.cupData;
    const excelData = job.excelData;
    const rawTitle = cupData.book_description || excelData['Title'] || '';
    const cleanedTitle = cleanTitle(rawTitle);
    const jacketRoute = getJacketRoute(cupData.trim_height, cupData.trim_width);
    
    return `<?xml version="1.0" encoding="UTF-8"?>
<BookJacket>
    <JobInfo>
        <ISBN>${cupData.isbn}</ISBN>
        <PaceJobNo>${excelData['Pace Job No'] || ''}</PaceJobNo>
        <CustomerOrderNo>${excelData['Customer Order No.'] || ''}</CustomerOrderNo>
        <OrderDate>${excelData['Order Date'] || ''}</OrderDate>
        <Quantity>${excelData['Qty'] || '1'}</Quantity>
    </JobInfo>
    <BookDetails>
        <Title>${escapeXml(cleanedTitle)}</Title>
        <Customer>${escapeXml(cupData.customer || '')}</Customer>
        <BindingType>${escapeXml(cupData.binding_type || '')}</BindingType>
    </BookDetails>
    <Specifications>
        <TrimHeight unit="mm">${cupData.trim_height || ''}</TrimHeight>
        <TrimWidth unit="mm">${cupData.trim_width || ''}</TrimWidth>
        <SpineSize unit="mm">${cupData.spine_size || ''}</SpineSize>
        <Pagination>${cupData.pagination || ''}</Pagination>
        <JacketRoute>${jacketRoute}</JacketRoute>
    </Specifications>
    <Materials>
        <StockDescription>${escapeXml(cupData.stock_description || '')}</StockDescription>
        <WeightGSM>${cupData.weight_gsm || ''}</WeightGSM>
        <CoverMediaTreatment>${escapeXml(cupData.cover_media_treatment || '')}</CoverMediaTreatment>
    </Materials>
    <Files>
        <PDFUrl>${escapeXml(cupData.pdf_url || '')}</PDFUrl>
    </Files>
</BookJacket>`;
}

// Escape XML special characters
function escapeXml(str) {
    if (!str) return '';
    return str.toString()
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&apos;');
}

// Truncate text
function truncateText(text, maxLength) {
    if (!text) return '';
    if (text.length <= maxLength) return text;
    return text.substring(0, maxLength - 3) + '...';
}

// Remove 'Cover' from end of title
function cleanTitle(title) {
    if (!title) return '';
    // Remove ' Cover' from the end (case insensitive)
    return title.replace(/\s+Cover\s*$/i, '').trim();
}

// Determine jacket route based on trim size
function getJacketRoute(trimHeight, trimWidth) {
    // Convert to numbers for comparison
    const height = parseFloat(trimHeight);
    const width = parseFloat(trimWidth);
    
    // If trim is 280mm × 216mm, route to Indigo, otherwise Ricoh
    if (height === 280 && width === 216) {
        return 'Indigo';
    }
    return 'Ricoh';
}

// Show status
function showStatus(message) {
    statusMessage.textContent = message;
    statusSection.style.display = 'block';
    resultsSection.style.display = 'none';
    errorSection.style.display = 'none';
}

// Hide status
function hideStatus() {
    statusSection.style.display = 'none';
}

// Show error
function showError(message) {
    errorMessage.textContent = message;
    errorSection.style.display = 'block';
    statusSection.style.display = 'none';
    resultsSection.style.display = 'none';
}

// Reset application
function resetApplication() {
    removeFile();
    errorSection.style.display = 'none';
}

// Initialize when DOM is ready
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
} else {
    init();
}