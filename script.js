let currentData = [];
let currentColumn = 1;
let currentPairs = [];

// List of conversation prompts/questions
const TWIST_OPTIONS = [
    "Colour Doppelgängers: Find 3 things you have in common, each matching a colour: red = something you both love, blue = something that calms you, green = something you both want to try. Share the traits with your selfie in the group.",
    "The \"We Should Actually Do This\" Dare: Talk until you find one activity, place, or idea that makes you both say \"Wait, we should actually do this,\" and Share the idea with your selfies in the group.",
    "Create Your Pair Flag: Choose 2 colours, 1 symbol, and 1 slogan that represent your pair. Example: blue + Orange, Sunset at the Sea, \"calm chaos.\" Share your flag idea with your selfie in the group.",
    "ESN Event From Your Pockets: Pick one random object you have with you right now and use it as inspiration for a funny ESN event idea. Example: headphones -> silent disco night. Share your ideas with the selfie on the group.",
    "The Tiny Tradition Dare: Invent a tiny tradition your pair would start if you had to meet every month. Example: trying a new drink, rating cafes, or taking the same selfie pose. Share your tradition with a selfie in the group.",
    "The Common Chaos Scale: Rate yourselves from 1 to 10 on chaos, chill, social energy, and planning skills, then share where you match the most with a selfie on the group.",
    "The Rainbow Menu: Create a fake menu for your pair with one food you both like, one drink you both like, and one activity you both enjoy. Share the menu with a selfie on the group.",
    "The 10-Minute Event Planner: Invent an event in 10 minutes that includes one colour, one snack, one activity, and one weird rule. Share the idea with the selfie in the group.",
    "The Photo Mission: Send a photo of something that represents your pair's vibe, an object, place, drink, meme, screenshot (ideally with a selfie) on the group.",
    "The Common Playlist Cover: Pick one song each for self and your pair, then give your 4-song playlist a name and a colour. Example: \"Purple Panic\" or \"Golden Hour Energy.\" Make sure to share the playlist with your selfie on the group.",
    "The Buddy Brand: Imagine your pair is a brand and decide your colour, slogan, mascot, and main personality trait. Example: green, \"good vibes only,\" turtle mascot, grounded energy. Do advertise your brand with a selfie on the group.",
    "The Event No One Asked For: Invent a ridiculous ESN event that sounds silly but you would secretly attend. Example: PowerPoint Night About Things We Care Too Much About. Share the idea on the group with your selfie.",
    "The Comfort Combo: Find your shared comfort combo: comfort food, comfort place, and comfort activity. Share your comfort find with a selfie on the group.",
    "The Newcomer Magic Trick: Discuss what helps you feel comfortable in a new setting, then invent one small \"magic trick\" that would help newcomers feel welcome at an ESN event. Share the trick with a selfie on the group.",
    "The \"Same But Different\" Dare: Find 3 things you both like, but for completely different reasons. Example: both like walking - one for calm, one for exploring. Share what makes you same yet different with a selfie on the group.",
    "The Buddy Bingo: Create a 3-box bingo card of things you have in common and send it to the group so others can see if they match your bingo too.",
    "The Common Superpower: Find one \"superpower\" you both have or both appreciate in others. Example: making people laugh, staying calm, organizing things, noticing details, or finding good food. Share your discovered super powers with a selfie on the group.",
    "The Unofficial ESN Award: Create an unofficial award your pair would win together. Example: \"Most Likely to Turn Coffee Into a 3-Hour Talk\" or \"Best Chill Vibes Committee.\" Do share the award you are winning with a selfie on the group."
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
    
    // Shuffle names and prompts using Fisher-Yates for better randomness each run
    const shuffled = shuffleArray(names);
    const shuffledTwists = shuffleArray(TWIST_OPTIONS);
    
    // Create pairs with random twists
    const pairs = [];
    for (let i = 0; i < shuffled.length; i += 2) {
        if (i + 1 < shuffled.length) {
            // Assign a randomized twist from the shuffled prompt list
            const randomTwist = shuffledTwists[(i / 2) % shuffledTwists.length];
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
                const randomTwist = shuffledTwists[(i / 2) % shuffledTwists.length];
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

function shuffleArray(input) {
    const arr = [...input];
    for (let i = arr.length - 1; i > 0; i--) {
        const j = Math.floor(Math.random() * (i + 1));
        [arr[i], arr[j]] = [arr[j], arr[i]];
    }
    return arr;
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
