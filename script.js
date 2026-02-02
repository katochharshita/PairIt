let currentData = [];
let currentColumn = 1;
let currentPairs = [];

// List of conversation prompts/questions
const TWIST_OPTIONS = [
    "Send your buddy a song and explain why it fits your mood lately.",
    "Each share one assumption people often make about you — and whether it's true.",
    "Describe a place that feels like \"home\" to you and why.",
    "Share one thing you're trying to unlearn.",
    "Finish this sentence together: \"Lately, I've been thinking a lot about…\"",
    "Share a photo from your camera roll that represents your week — explain why.",
    "Each describe a moment recently when you felt unexpectedly grateful.",
    "Recommend a habit you tried and kept (or tried and dropped) — what happened?",
    "Finish this sentence: \"Right now, I'm spending a lot of energy on…\"",
    "Share one thing you're better at than you were a year ago.",
    "Pick a word that describes how this month feels for you — explain it.",
    "Each share a piece of advice you'd give your past self from 2–3 years ago.",
    "Send your buddy a link (article, video, post) that stuck with you recently and say why.",
    "Describe a time you surprised yourself — good or bad.",
    "Share one boundary you've learned to set (or are learning to set).",
    "Finish this sentence honestly: \"Something I don't say out loud often is…\"",
    "Each name one thing that reliably improves your mood — even a little.",
    "Describe a place, activity, or routine where you feel most at ease.",
    "Share one question you're currently trying to answer in your life.",
    "Each share something you're intentionally saying \"no\" to lately.",
    "Describe what you wish people understood better about your work or daily life.",
    "Each share one thing you're hopeful about, even if it feels uncertain.",
    "Finish this sentence: \"I feel most supported when people…\"",
    "Share one small change that would make your next month noticeably better.",
    "Show each other something on your phone that makes you smile (photo, note, playlist, meme).",
    "Share one small win from the past week (nothing has to be impressive).",
    "Teach your buddy something tiny (a shortcut, tip, phrase, or fun fact).",
    "Describe your ideal lazy day in three steps.",
    "Exchange one recommendation (podcast, YouTube channel, app, food spot, book, or habit).",
    "Set a 60-second timer and rant about something harmless you love (coffee, dogs, stationery, niche hobby).",
    "Ask each other one question you've always wanted to ask new people but rarely do.",
    "Share one goal you're working toward right now — big or small.",
    "Agree on one thing you'll both try before your next catch-up (something you've been planning to).",
    "Swap a productivity hack or life shortcut you actually use.",
    "Each name one thing you're currently obsessed with (food, show, tool, song, hobby).",
    "Describe your perfect weekend morning in under 30 seconds.",
    "Send a GIF or emoji that matches your current mood — explain if you want.",
    "Name that one app or tool you'd be most annoyed to lose.",
    "Play \"This or That\" for at least 3 rounds",
    "Each share one thing that reliably makes your day better.",
    "Show a note, quote, list, or reminder you keep coming back to.",
    "Describe a food you could eat every week without getting bored.",
    "Each say one thing you're looking forward to this week.",
    "Teach each other a word, phrase, or saying you like (from any language or context).",
    "Share a playlist name or song title that fits your vibe lately."
];

document.getElementById('excelFile').addEventListener('change', handleFileUpload);
document.getElementById('columnNumber').addEventListener('input', handleColumnChange);
document.getElementById('generatePairs').addEventListener('click', generatePairs);
document.getElementById('regeneratePairs').addEventListener('click', generatePairs);
document.getElementById('downloadImage').addEventListener('click', downloadImage);

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

function downloadImage() {
    if (currentPairs.length === 0) {
        showError('No pairs to download. Please generate pairs first.');
        return;
    }
    
    // Create a temporary container for the image
    const tempContainer = document.createElement('div');
    tempContainer.style.position = 'absolute';
    tempContainer.style.left = '-9999px';
    tempContainer.style.width = '800px';
    tempContainer.style.padding = '40px';
    tempContainer.style.background = 'white';
    tempContainer.style.fontFamily = '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Oxygen, Ubuntu, Cantarell, sans-serif';
    
    // Header
    const header = document.createElement('div');
    header.style.background = 'linear-gradient(135deg, #667eea 0%, #764ba2 100%)';
    header.style.padding = '30px';
    header.style.borderRadius = '10px 10px 0 0';
    header.style.marginBottom = '30px';
    header.style.textAlign = 'center';
    
    const title = document.createElement('h1');
    title.textContent = 'Find your ESN Date';
    title.style.color = 'white';
    title.style.margin = '0 0 10px 0';
    title.style.fontSize = '32px';
    title.style.fontWeight = 'bold';
    
    const subtitle = document.createElement('p');
    subtitle.textContent = 'Meet your Coffee Date for the month and have fun!!';
    subtitle.style.color = 'rgba(255, 255, 255, 0.9)';
    subtitle.style.margin = '0';
    subtitle.style.fontSize = '14px';
    
    header.appendChild(title);
    header.appendChild(subtitle);
    tempContainer.appendChild(header);
    
    // Pairs container
    const pairsContainer = document.createElement('div');
    pairsContainer.style.display = 'grid';
    pairsContainer.style.gridTemplateColumns = 'repeat(2, 1fr)';
    pairsContainer.style.gap = '15px';
    
    currentPairs.forEach((pair, index) => {
        const pairCard = document.createElement('div');
        pairCard.style.background = 'linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%)';
        pairCard.style.padding = '20px';
        pairCard.style.borderRadius = '10px';
        pairCard.style.borderLeft = '4px solid #667eea';
        pairCard.style.boxShadow = '0 4px 6px rgba(0, 0, 0, 0.1)';
        
        const pairNumber = document.createElement('div');
        pairNumber.textContent = `#${index + 1}`;
        pairNumber.style.fontSize = '14px';
        pairNumber.style.color = '#666';
        pairNumber.style.fontWeight = 'bold';
        pairNumber.style.marginBottom = '10px';
        
        const names = document.createElement('div');
        names.textContent = pair.people.join('  &  ');
        names.style.fontSize = '16px';
        names.style.color = '#333';
        names.style.fontWeight = 'bold';
        names.style.marginBottom = '12px';
        
        const separator = document.createElement('div');
        separator.style.height = '1px';
        separator.style.background = '#ddd';
        separator.style.marginBottom = '12px';
        
        const twist = document.createElement('div');
        const cleanTwist = cleanTwistText(pair.twist);
        twist.textContent = cleanTwist;
        twist.style.fontSize = '11px';
        twist.style.color = '#667eea';
        twist.style.fontStyle = 'italic';
        
        pairCard.appendChild(pairNumber);
        pairCard.appendChild(names);
        pairCard.appendChild(separator);
        pairCard.appendChild(twist);
        pairsContainer.appendChild(pairCard);
    });
    
    tempContainer.appendChild(pairsContainer);
    document.body.appendChild(tempContainer);
    
    // Capture as image
    html2canvas(tempContainer, {
        backgroundColor: '#ffffff',
        scale: 2,
        logging: false,
        useCORS: true
    }).then(canvas => {
        // Convert to image and download
        canvas.toBlob(function(blob) {
            const url = URL.createObjectURL(blob);
            const link = document.createElement('a');
            link.href = url;
            const timestamp = new Date().toISOString().slice(0, 10).replace(/-/g, '');
            link.download = `ESN_Date_Pairs_${timestamp}.png`;
            link.click();
            URL.revokeObjectURL(url);
        }, 'image/png');
        
        // Clean up
        document.body.removeChild(tempContainer);
    }).catch(error => {
        showError('Error generating image: ' + error.message);
        document.body.removeChild(tempContainer);
    });
}

