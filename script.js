let currentData = [];
let currentColumn = 1;

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
    
    // Display pairs
    displayPairs(pairs);
    
    // Hide error if any
    document.getElementById('errorMessage').classList.add('hidden');
}

function displayPairs(pairs) {
    const container = document.getElementById('pairsContainer');
    container.innerHTML = '';
    
    pairs.forEach((pair, index) => {
        const pairCard = document.createElement('div');
        pairCard.className = 'pair-card';
        
        let namesHTML = '';
        if (pair.people.length === 2) {
            namesHTML = `
                <div class="pair-number">Pair #${index + 1}</div>
                <div class="pair-names">
                    <div>${pair.people[0]}</div>
                    <div class="pair-separator">☕</div>
                    <div>${pair.people[1]}</div>
                </div>
                <div class="pair-twist">✨ ${pair.twist}</div>
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
                <div class="pair-twist">✨ ${pair.twist}</div>
            `;
        } else {
            namesHTML = `
                <div class="pair-number">Person #${index + 1}</div>
                <div class="pair-names">${pair.people[0]}</div>
                <div class="pair-twist">✨ ${pair.twist}</div>
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

