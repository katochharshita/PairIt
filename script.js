let currentData = [];
let currentColumn = 1;
let currentPairs = [];

// 4 predefined twist options
const TWIST_OPTIONS = [
    "Share your favorite childhood memory",
    "Discuss a book or movie that changed your perspective",
    "Talk about your dream travel destination",
    "Share something you're passionate about"
];

document.getElementById('excelFile').addEventListener('change', handleFileUpload);
document.getElementById('columnNumber').addEventListener('input', handleColumnChange);
document.getElementById('generatePairs').addEventListener('click', generatePairs);
document.getElementById('regeneratePairs').addEventListener('click', generatePairs);
document.getElementById('downloadPDF').addEventListener('click', downloadPDF);

function handleFileUpload(event) {
    const file = event.target.files[0];
    if (!file) return;
    
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            
            // Get the first sheet
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            
            // Convert to JSON
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
            
            // Store the data
            currentData = jsonData;
            
            // Show file info
            const fileInfo = document.getElementById('fileInfo');
            fileInfo.textContent = `File loaded: ${file.name} | Sheet: ${firstSheetName} | Rows: ${jsonData.length}`;
            fileInfo.classList.remove('hidden');
            
            // Enable generate button
            document.getElementById('generatePairs').disabled = false;
            
            // Hide error if any
            document.getElementById('errorMessage').classList.add('hidden');
            
        } catch (error) {
            showError('Error reading file: ' + error.message);
        }
    };
    
    reader.readAsArrayBuffer(file);
}

function handleColumnChange(event) {
    currentColumn = parseInt(event.target.value) || 1;
    if (currentColumn < 1) {
        currentColumn = 1;
        event.target.value = 1;
    }
}

function generatePairs() {
    if (currentData.length === 0) {
        showError('Please upload an Excel file first.');
        return;
    }
    
    // Extract names from the specified column (convert to 0-based index)
    const columnIndex = currentColumn - 1;
    const names = [];
    
    for (let i = 0; i < currentData.length; i++) {
        const cellValue = currentData[i][columnIndex];
        if (cellValue && String(cellValue).trim() !== '') {
            names.push(String(cellValue).trim());
        }
    }
    
    if (names.length < 2) {
        showError(`Not enough names found in column ${currentColumn}. Need at least 2 names.`);
        return;
    }
    
    // Shuffle the names array
    const shuffled = [...names].sort(() => Math.random() - 0.5);
    
    // Create pairs with random twists
    const pairs = [];
    for (let i = 0; i < shuffled.length; i += 2) {
        if (i + 1 < shuffled.length) {
            // Randomly assign a twist to this pair
            const randomTwist = TWIST_OPTIONS[Math.floor(Math.random() * TWIST_OPTIONS.length)];
            pairs.push({
                people: [shuffled[i], shuffled[i + 1]],
                twist: randomTwist
            });
        } else {
            // If odd number of people, the last person can be paired with a random person
            // or we can add them to an existing pair to make a group of 3
            if (pairs.length > 0) {
                pairs[pairs.length - 1].people.push(shuffled[i]);
            } else {
                const randomTwist = TWIST_OPTIONS[Math.floor(Math.random() * TWIST_OPTIONS.length)];
                pairs.push({
                    people: [shuffled[i]],
                    twist: randomTwist
                });
            }
        }
    }
    
    // Store pairs globally for PDF generation
    currentPairs = pairs;
    
    // Display pairs
    displayPairs(pairs);
    
    // Hide error if any
    document.getElementById('errorMessage').classList.add('hidden');
}

function cleanTwistText(twist) {
    // Remove any leading unwanted characters like "'("
    let clean = twist.trim();
    // Remove "'(" at the start
    clean = clean.replace(/^'\s*\(/g, '');
    // Remove "'" at the start
    clean = clean.replace(/^'/g, '');
    // Remove "(" at the start
    clean = clean.replace(/^\(/g, '');
    return clean.trim();
}

function displayPairs(pairs) {
    const container = document.getElementById('pairsContainer');
    container.innerHTML = '';
    
    pairs.forEach((pair, index) => {
        const pairCard = document.createElement('div');
        pairCard.className = 'pair-card';
        
        const cleanTwist = cleanTwistText(pair.twist);
        
        let namesHTML = '';
        if (pair.people.length === 2) {
            namesHTML = `
                <div class="pair-number">Pair #${index + 1}</div>
                <div class="pair-names">
                    <div>${pair.people[0]}</div>
                    <div class="pair-separator">☕</div>
                    <div>${pair.people[1]}</div>
                </div>
                <div class="pair-twist">✨ ${cleanTwist}</div>
            `;
        } else if (pair.people.length === 3) {
            namesHTML = `
                <div class="pair-number">Group #${index + 1}</div>
                <div class="pair-names">
                    <div>${pair.people[0]}</div>
                    <div class="pair-separator">☕</div>
                    <div>${pair.people[1]}</div>
                    <div class="pair-separator">☕</div>
                    <div>${pair.people[2]}</div>
                </div>
                <div class="pair-twist">✨ ${cleanTwist}</div>
            `;
        } else {
            namesHTML = `
                <div class="pair-number">Person #${index + 1}</div>
                <div class="pair-names">${pair.people[0]}</div>
                <div class="pair-twist">✨ ${cleanTwist}</div>
            `;
        }
        
        pairCard.innerHTML = namesHTML;
        container.appendChild(pairCard);
    });
    
    document.getElementById('pairsSection').classList.remove('hidden');
}

function showError(message) {
    const errorDiv = document.getElementById('errorMessage');
    errorDiv.textContent = message;
    errorDiv.classList.remove('hidden');
}

function downloadPDF() {
    if (currentPairs.length === 0) {
        showError('No pairs to download. Please generate pairs first.');
        return;
    }
    
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF();
    
    // Set up colors
    const primaryColor = [102, 126, 234];
    const secondaryColor = [118, 75, 162];
    const accentColor = [255, 193, 7];
    
    // Header with gradient effect
    doc.setFillColor(...primaryColor);
    doc.rect(0, 0, 210, 40, 'F');
    
    // Decorative accent line
    doc.setFillColor(...accentColor);
    doc.rect(0, 38, 210, 2, 'F');
    
    // Title
    doc.setTextColor(255, 255, 255);
    doc.setFontSize(28);
    doc.setFont(undefined, 'bold');
    doc.text('Find your ESN Date', 105, 22, { align: 'center' });
    
    // Subtitle
    doc.setFontSize(11);
    doc.setFont(undefined, 'normal');
    doc.text('Meet your Coffee Date for the month and have fun!!', 105, 32, { align: 'center' });
    
    // Reset text color
    doc.setTextColor(0, 0, 0);
    
    let yPosition = 55;
    const pageHeight = 280;
    const margin = 15;
    const cardHeight = 45;
    const cardSpacing = 15;
    
    currentPairs.forEach((pair, index) => {
        // Check if we need a new page
        if (yPosition + cardHeight > pageHeight - margin) {
            doc.addPage();
            yPosition = 20;
        }
        
        // Card shadow effect (light gray rectangle behind)
        doc.setFillColor(240, 240, 240);
        doc.roundedRect(margin + 2, yPosition - 10, 180, cardHeight + 2, 5, 5, 'F');
        
        // Main card background with gradient-like effect
        doc.setFillColor(255, 255, 255);
        doc.setDrawColor(...primaryColor);
        doc.setLineWidth(0.5);
        doc.roundedRect(margin, yPosition - 12, 180, cardHeight, 5, 5, 'FD');
        
        // Left accent bar
        doc.setFillColor(...primaryColor);
        doc.roundedRect(margin, yPosition - 12, 4, cardHeight, 0, 0, 'F');
        
        // Pair number - just show # and number
        doc.setTextColor(100, 100, 100);
        doc.setFontSize(12);
        doc.setFont(undefined, 'bold');
        doc.text(`#${index + 1}`, margin + 10, yPosition - 2);
        
        // Names with better styling
        doc.setFontSize(16);
        doc.setTextColor(0, 0, 0);
        doc.setFont(undefined, 'bold');
        const namesText = pair.people.join('  &  ');
        doc.text(namesText, margin + 10, yPosition + 8);
        
        // Coffee emoji separator line
        doc.setDrawColor(200, 200, 200);
        doc.setLineWidth(0.3);
        doc.line(margin + 10, yPosition + 12, margin + 170, yPosition + 12);
        
        // Twist with cleaned text
        doc.setFontSize(10);
        doc.setTextColor(...primaryColor);
        doc.setFont(undefined, 'italic');
        const cleanTwist = cleanTwistText(pair.twist);
        // Split long text into multiple lines if needed
        const maxWidth = 160;
        const twistLines = doc.splitTextToSize(cleanTwist, maxWidth);
        doc.text(twistLines, margin + 10, yPosition + 20);
        
        yPosition += cardHeight + cardSpacing;
    });
    
    // Add footer with better styling
    const totalPages = doc.internal.pages.length - 1;
    for (let i = 1; i <= totalPages; i++) {
        doc.setPage(i);
        // Footer line
        doc.setDrawColor(200, 200, 200);
        doc.setLineWidth(0.5);
        doc.line(20, 275, 190, 275);
        
        doc.setFontSize(9);
        doc.setTextColor(150, 150, 150);
        doc.setFont(undefined, 'normal');
        doc.text(`Page ${i} of ${totalPages}`, 105, 282, { align: 'center' });
    }
    
    // Generate filename with timestamp
    const timestamp = new Date().toISOString().slice(0, 10).replace(/-/g, '');
    const filename = `ESN_Date_Pairs_${timestamp}.pdf`;
    
    // Save the PDF
    doc.save(filename);
}

