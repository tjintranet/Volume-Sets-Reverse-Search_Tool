// Initialize empty array for volume sets data
let volumeSetsData = [];
let uploadedISBNs = [];
let currentFile = null;
let searchResults = []; // Store results for PDF export

// Event Listeners for File Upload
function initializeUploadEvents() {
    const fileInput = document.getElementById('file-input');
    const exportBtn = document.getElementById('export-pdf-btn');

    fileInput.addEventListener('change', (e) => {
        if (e.target.files.length > 0) {
            handleFileUpload(e.target.files[0]);
        }
    });

    exportBtn.addEventListener('click', exportToPDF);
}

// Data Fetching
async function fetchData() {
    try {
        const response = await fetch('vol_sets.json');
        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }
        volumeSetsData = await response.json();
        console.log('Loaded volume sets data:', volumeSetsData);
    } catch (error) {
        console.error('Error fetching data:', error);
        showAlert('Error loading data. Please try again later.', 'danger');
    }
}

// Input Handlers - removed since no single search needed

// File Upload Handlers
function handleFileUpload(file) {
    const validTypes = ['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 
                       'application/vnd.ms-excel'];
    
    if (!validTypes.includes(file.type) && !file.name.match(/\.(xlsx|xls)$/i)) {
        showAlert('Please upload a valid Excel (.xlsx, .xls) file.', 'warning');
        return;
    }

    currentFile = file;
    showProgress(0);
    
    // Simulate upload progress and then process immediately
    let progress = 0;
    const interval = setInterval(() => {
        progress += 20;
        showProgress(progress);
        if (progress >= 100) {
            clearInterval(interval);
            setTimeout(() => {
                hideProgress();
                processUploadedFile();
            }, 300);
        }
    }, 50);
}

function showProgress(percent) {
    const progress = document.getElementById('upload-progress');
    const progressBar = progress.querySelector('.progress-bar');
    progress.style.display = 'block';
    progressBar.style.width = `${percent}%`;
    progressBar.textContent = `${percent}%`;
}

function hideProgress() {
    document.getElementById('upload-progress').style.display = 'none';
}

function clearUploadedFile() {
    currentFile = null;
    uploadedISBNs = [];
    searchResults = [];
    document.getElementById('bulk-summary').style.display = 'none';
    document.getElementById('bulk-results').innerHTML = '';
    document.getElementById('file-input').value = '';
    document.getElementById('export-pdf-btn').style.display = 'none';
}

// File Processing
async function processUploadedFile() {
    if (!currentFile) return;

    try {
        await parseExcelFile(currentFile);
        
        if (uploadedISBNs.length > 0) {
            performBulkSearch();
        } else {
            showAlert('No ISBN numbers found in the uploaded file.', 'warning');
        }
    } catch (error) {
        console.error('Error processing file:', error);
        showAlert('Error processing file. Please check the file format and try again.', 'danger');
    }
}

// Remove CSV parsing and single search functions

function parseExcelFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = function(e) {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];
                
                // Use sheet_to_json with header: 1 to get raw array data like in analysis
                const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                
                console.log('Excel file parsed, total rows:', jsonData.length);
                console.log('First few rows:', jsonData.slice(0, 3));
                
                // Pass the raw array data directly - now returns objects with isbn and orderNumber
                uploadedISBNs = extractISBNsFromArrayData(jsonData);
                resolve();
            } catch (error) {
                reject(error);
            }
        };
        reader.onerror = reject;
        reader.readAsArrayBuffer(file);
    });
}

function extractISBNsFromArrayData(jsonData) {
    const results = []; // Changed to store both ISBN and order number
    
    console.log('Processing Excel array data, total rows:', jsonData.length);
    
    // First, find all order headers
    const orderHeaders = [];
    jsonData.forEach((row, index) => {
        if (row[3] && row[3].toString().match(/^\d+$/) && 
            (row[0] || row[1]) && row.filter(cell => cell !== null && cell !== undefined).length >= 3) {
            orderHeaders.push({
                rowIndex: index,
                orderNumber: row[3].toString()
            });
        }
    });
    
    console.log(`Found ${orderHeaders.length} order headers`);
    
    // Now process VSet rows and assign order numbers
    jsonData.forEach((row, rowIndex) => {
        // Column L is index 11, Column F is index 5
        const columnL = row[11]; // Column L 
        const columnF = row[5];  // Column F
        
        // Check if column L contains "VSet"
        if (columnL && columnL.toString().includes('VSet')) {
            console.log(`Row ${rowIndex}: Found VSet in column L, Column F value:`, columnF);
            
            if (columnF) {
                // Handle both string and number formats
                let isbn = columnF.toString().replace(/[-\s]/g, '');
                
                console.log(`Processing ISBN: "${isbn}" (length: ${isbn.length})`);
                
                // Check if it's a valid 13-digit ISBN
                if (isbn.length === 13 && /^\d{13}$/.test(isbn) && isValidISBN13(isbn)) {
                    // Find the most recent order header before this VSet row
                    let assignedOrder = null;
                    for (let i = orderHeaders.length - 1; i >= 0; i--) {
                        if (orderHeaders[i].rowIndex < rowIndex) {
                            assignedOrder = orderHeaders[i].orderNumber;
                            break;
                        }
                    }
                    
                    // Check if this ISBN+order combination already exists
                    const existingIndex = results.findIndex(r => r.isbn === isbn && r.orderNumber === assignedOrder);
                    if (existingIndex === -1) {
                        results.push({
                            isbn: isbn,
                            orderNumber: assignedOrder
                        });
                        console.log(`Added ISBN: ${isbn} with Order: ${assignedOrder || 'NONE'}`);
                    }
                } else {
                    console.log(`Invalid ISBN: "${isbn}" - length: ${isbn.length}, digits only: ${/^\d{13}$/.test(isbn)}, valid checksum: ${isValidISBN13(isbn)}`);
                }
            }
        }
    });
    
    console.log('Final extracted results:', results);
    return results;
}

// Remove CSV extraction method since only Excel files are used

// Bulk Search - Search multiple component ISBNs to find their sets
function performBulkSearch() {
    if (!volumeSetsData || volumeSetsData.length === 0) {
        showAlert('Error: Volume sets data not loaded. Please refresh the page and try again.', 'danger');
        return;
    }

    const results = uploadedISBNs.map(item => {
        const componentISBN = item.isbn;
        const orderNumber = item.orderNumber;
        
        // Find the set that contains this component ISBN
        const result = volumeSetsData.find(set => {
            const associatedISBNs = set.associated_volumes.split(',').map(isbn => isbn.trim());
            return associatedISBNs.includes(componentISBN);
        });
        
        return { 
            isbn: componentISBN, 
            orderNumber: orderNumber,
            result: result 
        };
    });

    // Store results for PDF export
    searchResults = results;

    displayBulkResults(results);
    showBulkSummary(results);
}

function showBulkSummary(results) {
    const found = results.filter(r => r.result).length;
    const notFound = results.length - found;
    
    const summaryContent = document.getElementById('summary-content');
    summaryContent.innerHTML = `
        <div class="row text-center">
            <div class="col-4">
                <div class="fs-4 fw-bold text-primary">${results.length}</div>
                <small>Total ISBNs</small>
            </div>
            <div class="col-4">
                <div class="fs-4 fw-bold text-success">${found}</div>
                <small>Found</small>
            </div>
            <div class="col-4">
                <div class="fs-4 fw-bold text-warning">${notFound}</div>
                <small>Not Found</small>
            </div>
        </div>
    `;
    
    document.getElementById('bulk-summary').style.display = 'block';
    document.getElementById('export-pdf-btn').style.display = 'inline-block';
}

function displayBulkResults(results) {
    const bulkResults = document.getElementById('bulk-results');
    
    // Group results by set ISBN and order number
    const groupedResults = {};
    const notFoundResults = [];
    
    results.forEach(({ isbn: componentISBN, orderNumber, result }) => {
        if (result) {
            const setISBN = result.set_isbn;
            const groupKey = `${setISBN}_${orderNumber || 'NO_ORDER'}`;
            
            if (!groupedResults[groupKey]) {
                groupedResults[groupKey] = {
                    setData: result,
                    orderNumber: orderNumber,
                    componentISBNs: []
                };
            }
            groupedResults[groupKey].componentISBNs.push(componentISBN);
        } else {
            notFoundResults.push({ isbn: componentISBN, orderNumber });
        }
    });
    
    let html = '<h5 class="mb-3">Bulk Search Results:</h5>';
    
    // Display grouped results (found sets)
    Object.entries(groupedResults).forEach(([groupKey, { setData, orderNumber, componentISBNs }]) => {
        const volumeTitles = setData.volume_titles.split(' | ');
        const allISBNs = setData.associated_volumes.split(',').map(isbn => isbn.trim());
        
        // Create table rows for all volumes, highlighting the searched ones
        const volumeRows = allISBNs.map((isbn, index) => {
            const title = index < volumeTitles.length ? volumeTitles[index] : 'Unknown Title';
            const isSearchedISBN = componentISBNs.includes(isbn);
            
            return `
                <tr ${isSearchedISBN ? 'class="table-warning"' : ''}>
                    <td class="fw-bold" style="width: 175px;">
                        ${isbn}
                        ${isSearchedISBN ? '<i class="fas fa-arrow-left text-warning ms-2" title="Searched ISBN"></i>' : ''}
                    </td>
                    <td>${title}</td>
                </tr>
            `;
        }).join('');

        html += `
            <div class="card bulk-result-card border-success">
                <div class="card-body">
                    <div class="d-flex justify-content-between align-items-start mb-3">
                        <div>
                            <h6 class="card-title text-success mb-1">
                                <i class="fas fa-check-circle me-2"></i>Found ${componentISBNs.length} component${componentISBNs.length > 1 ? 's' : ''} in set
                            </h6>
                            ${orderNumber ? `<span class="badge bg-info text-dark">Order: ${orderNumber}</span>` : ''}
                        </div>
                        <span class="badge bg-success">${setData.volume_count} volumes</span>
                    </div>
                    <p class="mb-2"><strong>Set Title:</strong> ${setData.set_isbn_title}</p>
                    <p class="mb-3"><strong>Set ISBN:</strong> ${setData.set_isbn}</p>
                    <div class="mt-3">
                        <strong>All Volumes in Set:</strong>
                        <table class="table table-striped table-bordered table-sm mt-2">
                            <tbody>
                                ${volumeRows}
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>
        `;
    });
    
    // Display not found results grouped by order
    if (notFoundResults.length > 0) {
        // Group not found by order number
        const notFoundByOrder = {};
        notFoundResults.forEach(({ isbn, orderNumber }) => {
            const orderKey = orderNumber || 'NO_ORDER';
            if (!notFoundByOrder[orderKey]) {
                notFoundByOrder[orderKey] = [];
            }
            notFoundByOrder[orderKey].push(isbn);
        });
        
        Object.entries(notFoundByOrder).forEach(([orderKey, isbnList]) => {
            html += `
                <div class="card bulk-result-card border-warning">
                    <div class="card-body">
                        <div class="d-flex justify-content-between align-items-start mb-2">
                            <h6 class="card-title text-warning mb-1">
                                <i class="fas fa-exclamation-triangle me-2"></i>Component ISBNs not found (${isbnList.length})
                            </h6>
                            ${orderKey !== 'NO_ORDER' ? `<span class="badge bg-warning text-dark">Order: ${orderKey}</span>` : ''}
                        </div>
                        <div class="row">
                            ${isbnList.map(isbn => 
                                `<div class="col-md-6 col-lg-4 mb-1">
                                    <span class="isbn-chip not-found">${isbn}</span>
                                </div>`
                            ).join('')}
                        </div>
                        <p class="card-text text-muted mt-2">These component ISBNs were not found in any volume set</p>
                    </div>
                </div>
            `;
        });
    }
    
    bulkResults.innerHTML = html;
}

// PDF Export Function
function exportToPDF() {
    if (!searchResults || searchResults.length === 0) {
        showAlert('No search results to export.', 'warning');
        return;
    }

    const { jsPDF } = window.jspdf;
    const doc = new jsPDF();

    // Document setup
    const pageWidth = doc.internal.pageSize.width;
    const margin = 20;
    const contentWidth = pageWidth - (margin * 2);
    
    // Header
    doc.setFontSize(20);
    doc.setFont('helvetica', 'bold');
    doc.text('Volume Sets Reverse Search Report', margin, 25);
    
    // File info
    doc.setFontSize(12);
    doc.setFont('helvetica', 'normal');
    const fileName = currentFile ? currentFile.name : 'Unknown File';
    const reportDate = new Date().toLocaleDateString();
    doc.text(`File: ${fileName}`, margin, 40);
    doc.text(`Report Date: ${reportDate}`, margin, 50);
    
    let yPosition = 70;
    
    // Group results for PDF
    const groupedResults = {};
    const notFoundResults = [];
    
    searchResults.forEach(({ isbn: componentISBN, orderNumber, result }) => {
        if (result) {
            const setISBN = result.set_isbn;
            const groupKey = `${setISBN}_${orderNumber || 'NO_ORDER'}`;
            
            if (!groupedResults[groupKey]) {
                groupedResults[groupKey] = {
                    setData: result,
                    orderNumber: orderNumber,
                    componentISBNs: []
                };
            }
            groupedResults[groupKey].componentISBNs.push(componentISBN);
        } else {
            notFoundResults.push({ isbn: componentISBN, orderNumber });
        }
    });
    
    // Found Results Section
    if (Object.keys(groupedResults).length > 0) {
        doc.setFontSize(16);
        doc.setFont('helvetica', 'bold');
        doc.text('Found Volume Sets:', margin, yPosition);
        yPosition += 15;
        
        Object.entries(groupedResults).forEach(([groupKey, { setData, orderNumber, componentISBNs }]) => {
            // Check if we need a new page
            if (yPosition > 220) {
                doc.addPage();
                yPosition = 25;
            }
            
            // Set header
            doc.setFontSize(12);
            doc.setFont('helvetica', 'bold');
            const setHeader = `Set: ${setData.set_isbn_title}`;
            doc.text(setHeader, margin, yPosition);
            yPosition += 10;
            
            doc.setFont('helvetica', 'normal');
            doc.text(`Set ISBN: ${setData.set_isbn}`, margin, yPosition);
            yPosition += 8;
            
            if (orderNumber) {
                doc.text(`Order: ${orderNumber}`, margin, yPosition);
                yPosition += 8;
            }
            
            doc.text(`Found Components: ${componentISBNs.join(', ')}`, margin, yPosition);
            yPosition += 8;
            
            doc.text(`Total Volumes in Set: ${setData.volume_count}`, margin, yPosition);
            yPosition += 15;
            
            // Volume table
            const volumeTitles = setData.volume_titles.split(' | ');
            const allISBNs = setData.associated_volumes.split(',').map(isbn => isbn.trim());
            
            const tableData = allISBNs.map((isbn, index) => {
                const title = index < volumeTitles.length ? volumeTitles[index] : 'Unknown Title';
                const isSearched = componentISBNs.includes(isbn) ? '✓' : '';
                return [isbn, title.substring(0, 50) + (title.length > 50 ? '...' : ''), isSearched];
            });
            
            doc.autoTable({
                startY: yPosition,
                head: [['ISBN', 'Title', 'Found']],
                body: tableData,
                margin: { left: margin, right: margin },
                styles: { fontSize: 8, cellPadding: 2 },
                headStyles: { fillColor: [76, 175, 80], textColor: 255 },
                columnStyles: {
                    0: { cellWidth: 35 },
                    1: { cellWidth: 90 },
                    2: { cellWidth: 15, halign: 'center' }
                },
                didParseCell: function (data) {
                    if (data.column.index === 2 && data.cell.text[0] === '✓') {
                        data.cell.styles.fillColor = [255, 235, 59];
                        data.cell.styles.textColor = [0, 0, 0];
                    }
                }
            });
            
            yPosition = doc.lastAutoTable.finalY + 15;
        });
    }
    
    // Not Found Results Section
    if (notFoundResults.length > 0) {
        // Check if we need a new page
        if (yPosition > 200) {
            doc.addPage();
            yPosition = 25;
        }
        
        doc.setFontSize(16);
        doc.setFont('helvetica', 'bold');
        doc.text('Component ISBNs Not Found:', margin, yPosition);
        yPosition += 15;
        
        // Group not found by order
        const notFoundByOrder = {};
        notFoundResults.forEach(({ isbn, orderNumber }) => {
            const orderKey = orderNumber || 'No Order';
            if (!notFoundByOrder[orderKey]) {
                notFoundByOrder[orderKey] = [];
            }
            notFoundByOrder[orderKey].push(isbn);
        });
        
        Object.entries(notFoundByOrder).forEach(([orderKey, isbnList]) => {
            doc.setFontSize(12);
            doc.setFont('helvetica', 'bold');
            doc.text(`${orderKey}:`, margin, yPosition);
            yPosition += 8;
            
            doc.setFont('helvetica', 'normal');
            const isbnText = isbnList.join(', ');
            const lines = doc.splitTextToSize(isbnText, contentWidth);
            doc.text(lines, margin, yPosition);
            yPosition += lines.length * 6 + 10;
        });
    }
    
    // Footer on last page
    const pageCount = doc.internal.getNumberOfPages();
    for (let i = 1; i <= pageCount; i++) {
        doc.setPage(i);
        doc.setFontSize(8);
        doc.setFont('helvetica', 'normal');
        doc.text(`Page ${i} of ${pageCount}`, pageWidth - margin, doc.internal.pageSize.height - 10, { align: 'right' });
        doc.text('Generated by Volume Sets Reverse Search Tool', margin, doc.internal.pageSize.height - 10);
    }
    
    // Save the PDF
    const timestamp = new Date().toISOString().slice(0, 19).replace(/:/g, '-');
    const filename = `volume-sets-report-${timestamp}.pdf`;
    doc.save(filename);
}

// Utility Functions
function isValidISBN13(isbn) {
    isbn = isbn.replace(/[-\s]/g, '');
    if (isbn.length !== 13 || !/^\d{13}$/.test(isbn)) return false;
    let sum = 0;
    for (let i = 0; i < 12; i++) {
        sum += (i % 2 === 0) ? parseInt(isbn[i]) : parseInt(isbn[i]) * 3;
    }
    let checksum = (10 - (sum % 10)) % 10;
    return isbn[12] === checksum.toString();
}

function showInvalidISBNFeedback() {
    // Removed since no single search needed
}

function showAlert(message, type) {
    const alertDiv = document.createElement('div');
    alertDiv.className = `alert alert-${type} alert-dismissible fade show`;
    alertDiv.innerHTML = `
        ${message}
        <button type="button" class="btn-close" data-bs-dismiss="alert" aria-label="Close"></button>
    `;
    const cardBody = document.querySelector('.card-body');
    cardBody.insertBefore(alertDiv, cardBody.firstChild);
}

function clearSearch() {
    // Removed since no single search needed
}

// Initialize
document.addEventListener('DOMContentLoaded', async () => {
    await fetchData();
    initializeUploadEvents();
    console.log('Volume Sets Reverse Search Tool initialized');
});